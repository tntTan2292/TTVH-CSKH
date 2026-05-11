import http.server
import socketserver
import os

PORT = 8088
DIRECTORY = r"d:\Antigravity - Project - TTVH\CSKH\backend"

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

if __name__ == "__main__":
    print(f"FRONTEND DASHBOARD starting at http://localhost:{PORT}")
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        httpd.serve_forever()
