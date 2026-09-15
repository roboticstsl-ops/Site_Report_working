package com.roboticstsl.sitereport

import android.annotation.SuppressLint
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Bundle
import android.os.Handler
import android.os.Looper
import android.provider.MediaStore
import android.provider.OpenableColumns
import android.util.Base64
import android.util.Log
import android.webkit.JavascriptInterface
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebResourceError
import android.webkit.WebResourceRequest
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.FileProvider
import com.chaquo.python.Python
import com.chaquo.python.android.AndroidPlatform
import org.json.JSONObject
import java.io.File

/**
 * Loads the existing Flask app (app.py, unmodified) running on-device via
 * Chaquopy, inside a WebView -- the same UI, forms, and PDF/DOCX-generation
 * logic as the server deployment, just packaged for offline use. The only
 * thing that needs the internet is the "send to Telegram" button.
 */
class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var pendingFileCallback: ValueCallback<Array<Uri>>? = null
    private var pendingCameraUri: Uri? = null
    private var loadAttempts = 0
    private val maxLoadAttempts = 25 // ~10s of retrying at 400ms apart before giving up

    private val filePickerLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        val callback = pendingFileCallback
        pendingFileCallback = null
        if (callback == null) return@registerForActivityResult

        val data = result.data
        val results: Array<Uri>? = when {
            result.resultCode != RESULT_OK -> null
            data?.dataString != null -> arrayOf(Uri.parse(data.dataString))
            data?.clipData != null -> {
                val clip = data.clipData!!
                Array(clip.itemCount) { i -> clip.getItemAt(i).uri }
            }
            pendingCameraUri != null -> arrayOf(pendingCameraUri!!)
            else -> null
        }
        callback.onReceiveValue(results)
    }

    private val cameraPermissionLauncher = registerForActivityResult(
        ActivityResultContracts.RequestPermission()
    ) { /* if denied, the chooser below still offers the gallery */ }

    // "Save to Device" -- the web page hands us a base64 payload + suggested
    // filename via AndroidBridge.saveFile(); this launches Android's real
    // "choose where to save" picker (Storage Access Framework) and writes
    // the bytes to whatever the user picked, then reports the result back
    // into the page via window.onAndroidSaveResult(ok, displayName).
    private var pendingSaveBytes: ByteArray? = null
    private val createDocumentLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument("*/*")
    ) { uri: Uri? ->
        val bytes = pendingSaveBytes
        pendingSaveBytes = null
        if (uri == null || bytes == null) {
            notifySaveResult(false, null)
            return@registerForActivityResult
        }
        try {
            contentResolver.openOutputStream(uri)?.use { it.write(bytes) } ?: throw java.io.IOException("no output stream")
            notifySaveResult(true, queryDisplayName(uri))
        } catch (e: Exception) {
            Log.e("SiteReport", "save to device failed", e)
            notifySaveResult(false, null)
        }
    }

    private fun notifySaveResult(ok: Boolean, displayName: String?) {
        val nameJs = if (displayName != null) JSONObject.quote(displayName) else "null"
        webView.post {
            webView.evaluateJavascript(
                "window.onAndroidSaveResult && window.onAndroidSaveResult($ok, $nameJs)", null
            )
        }
    }

    private fun queryDisplayName(uri: Uri): String? {
        return try {
            contentResolver.query(uri, arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)?.use { c ->
                if (c.moveToFirst()) c.getString(0) else null
            }
        } catch (e: Exception) {
            null
        }
    }

    inner class AndroidBridge {
        @JavascriptInterface
        fun saveFile(base64Data: String, filename: String, @Suppress("UNUSED_PARAMETER") mimeType: String) {
            val bytes = try {
                Base64.decode(base64Data, Base64.DEFAULT)
            } catch (e: Exception) {
                Log.e("SiteReport", "bad base64 from web page", e)
                notifySaveResult(false, null)
                return
            }
            pendingSaveBytes = bytes
            runOnUiThread { createDocumentLauncher.launch(filename) }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        webView = WebView(this)
        setContentView(webView)
        setupWebView()

        try {
            if (!Python.isStarted()) {
                Python.start(AndroidPlatform(this))
            }
            val python = Python.getInstance()
            val environ = python.getModule("os").get("environ")!!
            // app.py's own temp/report dirs default to '/tmp', which isn't
            // writable on Android -- point tempfile.gettempdir() at the
            // app's actual writable cache dir instead. See app.py.
            environ.callAttr("__setitem__", "TMPDIR", cacheDir.absolutePath)
            // Same app.py as the server deployment; this flag only skips the
            // LibreOffice PDF step (unavailable on Android) -- see app.py.
            environ.callAttr("__setitem__", "OFFLINE_NO_PDF", "1")
            environ.callAttr("__setitem__", "TELEGRAM_BOT_TOKEN", BuildConfig.TELEGRAM_BOT_TOKEN)
            environ.callAttr("__setitem__", "TELEGRAM_CHAT_ID", BuildConfig.TELEGRAM_CHAT_ID)
            // Importing this starts the embedded Flask server (module-level
            // code runs once per process; safe to call again on Activity
            // recreation).
            python.getModule("android_entry")
        } catch (t: Throwable) {
            // Show the real error instead of a silent crash -- there's no
            // device/emulator in the build pipeline to catch this earlier.
            Log.e("SiteReport", "Python/Flask startup failed", t)
            webView.loadData(
                "<pre style=\"white-space:pre-wrap;font-family:monospace;padding:16px\">" +
                    "Startup failed:\n${Log.getStackTraceString(t).replace("<", "&lt;")}</pre>",
                "text/html", "UTF-8"
            )
            return
        }

        if (ActivityCompat.checkSelfPermission(this, android.Manifest.permission.CAMERA)
            != PackageManager.PERMISSION_GRANTED
        ) {
            cameraPermissionLauncher.launch(android.Manifest.permission.CAMERA)
        }

        loadAppWithRetry()
    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun setupWebView() {
        webView.settings.javaScriptEnabled = true
        webView.settings.domStorageEnabled = true
        webView.settings.allowFileAccess = true
        webView.settings.mediaPlaybackRequiresUserGesture = false
        webView.addJavascriptInterface(AndroidBridge(), "AndroidBridge")

        webView.webViewClient = object : WebViewClient() {
            override fun onReceivedError(
                view: WebView?, request: WebResourceRequest?, error: WebResourceError?
            ) {
                // The embedded Flask server can take a few hundred ms to bind
                // on first launch; retry (connection-refused only, bounded)
                // instead of showing a dead page.
                val connectionError = error?.errorCode == WebViewClient.ERROR_CONNECT ||
                    error?.errorCode == WebViewClient.ERROR_HOST_LOOKUP ||
                    error?.errorCode == WebViewClient.ERROR_TIMEOUT
                if (request?.isForMainFrame == true && connectionError && loadAttempts < maxLoadAttempts) {
                    loadAttempts++
                    Handler(Looper.getMainLooper()).postDelayed({ loadAppWithRetry() }, 400)
                }
            }

            override fun onPageFinished(view: WebView?, url: String?) {
                loadAttempts = 0
            }
        }

        webView.webChromeClient = object : WebChromeClient() {
            override fun onShowFileChooser(
                view: WebView?,
                filePathCallback: ValueCallback<Array<Uri>>?,
                fileChooserParams: FileChooserParams?
            ): Boolean {
                pendingFileCallback?.onReceiveValue(null)
                pendingFileCallback = filePathCallback

                val captureEnabled = fileChooserParams?.isCaptureEnabled == true

                fun buildCameraIntent(): android.content.Intent? {
                    val photoFile = File(
                        File(cacheDir, "camera").apply { mkdirs() },
                        "capture_${System.currentTimeMillis()}.jpg"
                    )
                    val photoUri = FileProvider.getUriForFile(
                        this@MainActivity, "$packageName.fileprovider", photoFile
                    )
                    pendingCameraUri = photoUri
                    val intent = android.content.Intent(MediaStore.ACTION_IMAGE_CAPTURE).apply {
                        putExtra(MediaStore.EXTRA_OUTPUT, photoUri)
                        addFlags(android.content.Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
                    }
                    return if (intent.resolveActivity(packageManager) != null) intent else null
                }

                // The "Take Image" quick-capture buttons use <input capture>,
                // which is exactly the signal that the page wants the camera
                // and nothing else -- skip the gallery-vs-camera chooser
                // dialog entirely and jump straight to the camera app.
                if (captureEnabled) {
                    val cameraIntent = buildCameraIntent()
                    if (cameraIntent != null) {
                        filePickerLauncher.launch(cameraIntent)
                        return true
                    }
                    // No camera app available -- fall through to the normal
                    // gallery chooser below rather than doing nothing.
                }

                val intents = mutableListOf<android.content.Intent>()
                if (fileChooserParams?.mode == FileChooserParams.MODE_OPEN) {
                    buildCameraIntent()?.let { intents.add(it) }
                }

                val galleryIntent = android.content.Intent(android.content.Intent.ACTION_GET_CONTENT).apply {
                    type = "image/*"
                    putExtra(android.content.Intent.EXTRA_ALLOW_MULTIPLE, true)
                    addCategory(android.content.Intent.CATEGORY_OPENABLE)
                }

                val chooser = android.content.Intent.createChooser(galleryIntent, "Add photo").apply {
                    if (intents.isNotEmpty()) {
                        putExtra(android.content.Intent.EXTRA_INITIAL_INTENTS, intents.toTypedArray())
                    }
                }
                filePickerLauncher.launch(chooser)
                return true
            }
        }
    }

    private fun loadAppWithRetry() {
        webView.loadUrl("http://127.0.0.1:5000/")
    }

    override fun onBackPressed() {
        if (webView.canGoBack()) webView.goBack() else super.onBackPressed()
    }
}
