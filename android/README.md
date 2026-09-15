# Android build

Wraps the existing Flask app (`app.py`, `templates/`, `static/`,
`WorkingTemplate.docx` — all at the repo root, unchanged) in a WebView, with
the Flask server itself running on-device via [Chaquopy](https://chaquo.com/chaquopy/)
(Python embedded in the APK) instead of on a remote host. Same UI, same form,
same DOCX generation — fully offline. The only network call the app ever
makes is "Send to Telegram", which posts the finished report to a Telegram
bot/chat.

## Why not a PDF on Android?

The server deployment converts DOCX → PDF with LibreOffice headless.
LibreOffice can't run on Android, so the Android build sets `OFFLINE_NO_PDF=1`
(see `app.py`) and only produces the `.docx` — sent to Telegram / saved
locally instead of a PDF. Nothing else about `app.py` changed.

## How the pieces fit together

- `app/src/main/python/` — **generated**, not committed. Gradle's
  `syncPythonSource` task mirrors `app.py`, `templates/`, `static/`, and
  `WorkingTemplate.docx` in from the repo root before every build, so the
  repo root stays the single source of truth.
- `app/src/main/python_native/android_entry.py` — the only Android-specific
  Python file. Starts `app.py`'s Flask app on `127.0.0.1:5000` in a
  background thread when the app launches.
- `MainActivity.kt` — starts Chaquopy, sets `OFFLINE_NO_PDF` /
  `TELEGRAM_BOT_TOKEN` / `TELEGRAM_CHAT_ID` as environment variables the
  Python side reads, then points a `WebView` at `http://127.0.0.1:5000/`.
  Also wires up the system camera/gallery picker for the form's photo
  inputs (`<input type=file>` needs this explicitly in a WebView).

## Building

CI (`.github/workflows/build-apk.yml`) builds this on every push to `main`
that touches the app or `android/`, or on demand from the Actions tab —
download the APK from that run's **Artifacts**. It's a debug build
(self-signed by the Android debug keystore), meant for direct sideload:
enable "install unknown apps" for your file manager / browser, open the
downloaded APK, install.

To build locally instead: `cd android && ./gradlew :app:assembleDebug`
(needs Android SDK + `ANDROID_HOME` set; run `gradle wrapper` first if
`gradlew` isn't present yet).

## Telegram bot token

Set as GitHub Actions repository secrets (`TELEGRAM_BOT_TOKEN`,
`TELEGRAM_CHAT_ID`), never committed to the repo — Gradle bakes them into
`BuildConfig` at build time from those secrets. Update them any time via
repo Settings → Secrets and variables → Actions.
