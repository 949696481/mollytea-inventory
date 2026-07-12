import http.client
import os
import socket
import subprocess
import time

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = 8765
URL = f"http://localhost:{PORT}/index.html"
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"


def server_running():
    try:
        conn = http.client.HTTPConnection("localhost", PORT, timeout=0.5)
        conn.request("GET", "/index.html")
        resp = conn.getresponse()
        return resp.status == 200
    except (OSError, socket.error):
        return False


if not server_running():
    subprocess.Popen(
        ["pythonw", "-m", "http.server", str(PORT)],
        cwd=APP_DIR,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    for _ in range(20):
        time.sleep(0.25)
        if server_running():
            break

subprocess.Popen([EDGE_PATH, f"--app={URL}"])
