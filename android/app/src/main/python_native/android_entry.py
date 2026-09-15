"""Entry point imported once by MainActivity on app start. Starts the
existing Flask app (app.py, synced in unmodified from the repo root) on
loopback in a background thread -- the WebView then just loads it like any
other local page. Nothing about app.py's own logic changes; this file only
starts it."""
import threading

from werkzeug.serving import make_server

from app import app as flask_app

PORT = 5000

_server = make_server("127.0.0.1", PORT, flask_app, threaded=True)
threading.Thread(target=_server.serve_forever, daemon=True).start()
