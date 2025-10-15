#!/usr/bin/env python3
"""
Simple HTTP server for the ULLTRA dashboard
Run this to avoid CORS issues when testing locally
"""

import http.server
import socketserver
import webbrowser
import os
from pathlib import Path

PORT = 8000

class CORSHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

def main():
    # Change to the directory containing this script
    os.chdir(Path(__file__).parent)
    
    with socketserver.TCPServer(("", PORT), CORSHTTPRequestHandler) as httpd:
        print(f"ULLTRA Dashboard Server")
        print(f"Serving at http://localhost:{PORT}")
        print(f"Open your browser to: http://localhost:{PORT}")
        print("Press Ctrl+C to stop the server")
        
        # Try to open browser automatically
        try:
            webbrowser.open(f'http://localhost:{PORT}')
        except:
            pass
            
        httpd.serve_forever()

if __name__ == "__main__":
    main()