import os
import subprocess
import sys
import threading
import time
import functools
import http.server
import socketserver

ROOT = os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), "app")
NAME = os.path.splitext(os.path.basename(sys.executable))[0]

BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


LAST = [time.time()]


class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        LAST[0] = time.time()
        if self.path.startswith("/hb"):
            self.send_response(204)
            self.end_headers()
            return
        super().do_GET()


def main():
    handler = functools.partial(Quiet, directory=ROOT)
    srv = socketserver.ThreadingTCPServer(("127.0.0.1", 0), handler)
    srv.daemon_threads = True
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    url = f"http://127.0.0.1:{srv.server_address[1]}/index.html"
    profile = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "OrionCoreCalc", NAME)
    os.makedirs(profile, exist_ok=True)
    exe = next((b for b in BROWSERS if os.path.exists(b)), None)
    if not exe:
        import webbrowser
        webbrowser.open(url)
        threading.Event().wait()
        return
    subprocess.run([exe, f"--app={url}", f"--user-data-dir={profile}", "--window-size=1440,900",
                    "--no-first-run", "--no-default-browser-check"])
    # The browser may hand the window to an existing process and exit at once;
    # keep serving until the page stops sending its heartbeat.
    LAST[0] = time.time()
    while time.time() - LAST[0] < 20:
        time.sleep(2)


if __name__ == "__main__":
    main()
