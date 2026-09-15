package com.roboticstsl.sitereport

import android.annotation.SuppressLint
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Bundle
import android.os.Handler
import android.os.Looper
import android.provider.MediaStore
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

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        if (!Python.isStarted()) {
            Python.start(AndroidPlatform(this))
        }
        val python = Python.getInstance()
        val environ = python.getModule("os").get("environ")!!
        // Same app.py as the server deployment; this flag only skips the
        // LibreOffice PDF step (unavailable on Android) -- see app.py.
        environ.callAttr("__setitem__", "OFFLINE_NO_PDF", "1")
        environ.callAttr("__setitem__", "TELEGRAM_BOT_TOKEN", BuildConfig.TELEGRAM_BOT_TOKEN)
        environ.callAttr("__setitem__", "TELEGRAM_CHAT_ID", BuildConfig.TELEGRAM_CHAT_ID)
        // Importing this starts the embedded Flask server (module-level code
        // runs once per process; safe to call again on Activity recreation).
        python.getModule("android_entry")

        if (ActivityCompat.checkSelfPermission(this, android.Manifest.permission.CAMERA)
            != PackageManager.PERMISSION_GRANTED
        ) {
            cameraPermissionLauncher.launch(android.Manifest.permission.CAMERA)
        }

        webView = WebView(this)
        setContentView(webView)
        setupWebView()
        loadAppWithRetry()
    }

    @SuppressLint("SetJavaScriptEnabled")
    private fun setupWebView() {
        webView.settings.javaScriptEnabled = true
        webView.settings.domStorageEnabled = true
        webView.settings.allowFileAccess = true
        webView.settings.mediaPlaybackRequiresUserGesture = false

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
                val intents = mutableListOf<android.content.Intent>()

                if (captureEnabled || fileChooserParams?.mode == FileChooserParams.MODE_OPEN) {
                    val photoFile = File(
                        File(cacheDir, "camera").apply { mkdirs() },
                        "capture_${System.currentTimeMillis()}.jpg"
                    )
                    val photoUri = FileProvider.getUriForFile(
                        this@MainActivity, "$packageName.fileprovider", photoFile
                    )
                    pendingCameraUri = photoUri
                    val cameraIntent = android.content.Intent(MediaStore.ACTION_IMAGE_CAPTURE).apply {
                        putExtra(MediaStore.EXTRA_OUTPUT, photoUri)
                        addFlags(android.content.Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
                    }
                    if (cameraIntent.resolveActivity(packageManager) != null) {
                        intents.add(cameraIntent)
                    }
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
