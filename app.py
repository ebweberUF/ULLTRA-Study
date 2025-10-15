#!/usr/bin/env python3
"""
ULLTRA Study Dashboard - Desktop Application
PyInstaller-ready application with embedded web server and system tray
"""

import sys
import os
import threading
import webbrowser
import socket
from pathlib import Path
import http.server
import socketserver
from urllib.parse import urlparse

# GUI imports
try:
    import tkinter as tk
    from tkinter import messagebox, Menu
    import tkinter.ttk as ttk
    GUI_AVAILABLE = True
except ImportError:
    GUI_AVAILABLE = False

# Try to import system tray support
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

class CORSHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    """HTTP request handler with CORS support and API endpoints"""

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, format, *args):
        # Suppress HTTP server logs for cleaner output
        return

    def do_POST(self):
        """Handle POST requests for API endpoints"""
        from sharepoint_api import get_sharepoint_manager
        import json

        # SharePoint API endpoints
        if self.path == '/api/sharepoint/auth/start':
            self.handle_sharepoint_auth_start()
        elif self.path == '/api/sharepoint/auth/status':
            self.handle_sharepoint_auth_status()
        elif self.path == '/api/sharepoint/auth/logout':
            self.handle_sharepoint_logout()
        elif self.path == '/api/sharepoint/events':
            self.handle_sharepoint_events()
        else:
            self.send_error(404, "Endpoint not found")

    def handle_sharepoint_auth_start(self):
        """Start SharePoint device code authentication"""
        import json
        from sharepoint_api import get_sharepoint_manager

        try:
            # Read request body for configuration
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length > 0:
                body = self.rfile.read(content_length)
                config = json.loads(body.decode('utf-8'))
            else:
                config = {}

            # Get SharePoint configuration
            site_url = config.get('site_url', 'https://uflorida.sharepoint.com/sites/PRICE')
            list_name = config.get('list_name', 'PRICECalendar')

            # Get or create SharePoint manager
            sp_manager = get_sharepoint_manager(site_url, list_name)

            # Start device code flow
            result = sp_manager.start_device_code_flow()

            # Send response
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(result).encode())

        except Exception as e:
            self.send_error(500, f"Authentication error: {str(e)}")

    def handle_sharepoint_auth_status(self):
        """Check SharePoint authentication status"""
        import json
        from sharepoint_api import get_sharepoint_manager

        try:
            sp_manager = get_sharepoint_manager()
            status = sp_manager.get_auth_status()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(status).encode())

        except Exception as e:
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'authenticated': False,
                'error': 'Not initialized'
            }).encode())

    def handle_sharepoint_logout(self):
        """Logout from SharePoint"""
        import json
        from sharepoint_api import get_sharepoint_manager

        try:
            sp_manager = get_sharepoint_manager()
            sp_manager.logout()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            self.send_error(500, f"Logout error: {str(e)}")

    def handle_sharepoint_events(self):
        """Fetch SharePoint calendar events"""
        import json
        from sharepoint_api import get_sharepoint_manager

        try:
            sp_manager = get_sharepoint_manager()

            if not sp_manager.is_authenticated():
                self.send_response(401)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'error': 'Not authenticated'
                }).encode())
                return

            # Fetch events
            events = sp_manager.get_calendar_events()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'success': True,
                'events': events,
                'count': len(events)
            }).encode())

        except Exception as e:
            self.send_error(500, f"Error fetching events: {str(e)}")

class ULLTRADashboard:
    """Main application class for ULLTRA Dashboard"""
    
    def __init__(self):
        self.port = self.find_free_port()
        self.server = None
        self.server_thread = None
        self.root = None
        self.tray = None
        self.running = False
        
        # Determine the base directory (for PyInstaller compatibility)
        if getattr(sys, 'frozen', False):
            # Running as PyInstaller executable
            self.base_dir = Path(sys._MEIPASS)
        else:
            # Running as Python script
            self.base_dir = Path(__file__).parent
            
        print(f"Base directory: {self.base_dir}")
        
    def find_free_port(self, start_port=8000):
        """Find a free port starting from start_port"""
        port = start_port
        while port < start_port + 100:
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(('localhost', port))
                    return port
            except OSError:
                port += 1
        raise Exception("Could not find a free port")
    
    def start_server(self):
        """Start the HTTP server in a separate thread"""
        try:
            # Change to the base directory where web files are located
            os.chdir(self.base_dir)
            
            # Verify web files exist
            required_files = ['index.html', 'styles.css', 'script.js']
            missing_files = [f for f in required_files if not (self.base_dir / f).exists()]
            
            if missing_files:
                raise Exception(f"Missing required files: {', '.join(missing_files)}")
            
            # Create and start server
            self.server = socketserver.TCPServer(("127.0.0.1", self.port), CORSHTTPRequestHandler)
            self.server_thread = threading.Thread(target=self.server.serve_forever, daemon=True)
            self.server_thread.start()
            self.running = True
            
            print(f"Server started on http://127.0.0.1:{self.port}")
            return True
            
        except Exception as e:
            print(f"Error starting server: {e}")
            if GUI_AVAILABLE:
                messagebox.showerror("Server Error", f"Failed to start server: {e}")
            return False
    
    def stop_server(self):
        """Stop the HTTP server"""
        if self.server:
            self.server.shutdown()
            self.server.server_close()
            self.running = False
            print("Server stopped")
    
    def open_dashboard(self):
        """Open the dashboard in the default web browser"""
        url = f"http://127.0.0.1:{self.port}"
        try:
            webbrowser.open(url)
            print(f"Dashboard opened: {url}")
        except Exception as e:
            print(f"Error opening browser: {e}")
            if GUI_AVAILABLE:
                messagebox.showwarning("Browser Error", f"Could not open browser automatically.\nPlease visit: {url}")
    
    def create_tray_icon(self):
        """Create system tray icon"""
        if not TRAY_AVAILABLE:
            return None
            
        # Create a simple icon
        image = Image.new('RGB', (64, 64), color='blue')
        draw = ImageDraw.Draw(image)
        draw.text((10, 20), 'U', fill='white', anchor='mm')
        
        # Create menu
        menu = pystray.Menu(
            pystray.MenuItem("Open Dashboard", self.open_dashboard),
            pystray.MenuItem("Show Window", self.show_window),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self.quit_app)
        )
        
        # Create tray icon
        self.tray = pystray.Icon("ULLTRA Dashboard", image, menu=menu)
        return self.tray
    
    def create_gui(self):
        """Create the main GUI window"""
        if not GUI_AVAILABLE:
            return None
            
        self.root = tk.Tk()
        self.root.title("ULLTRA Study Dashboard")
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="ULLTRA Study Dashboard", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Subtitle
        subtitle_label = ttk.Label(main_frame, 
                                  text="Photobiomodulation for TMD Pain Management")
        subtitle_label.pack(pady=(0, 20))
        
        # Server status
        self.status_label = ttk.Label(main_frame, text="Server Status: Starting...", 
                                     font=('Arial', 10))
        self.status_label.pack(pady=(0, 10))
        
        # URL display
        self.url_label = ttk.Label(main_frame, text="", 
                                  font=('Arial', 9), foreground='blue')
        self.url_label.pack(pady=(0, 20))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        # Open Dashboard button
        self.open_btn = ttk.Button(button_frame, text="Open Dashboard", 
                                  command=self.open_dashboard, width=15)
        self.open_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Minimize to tray button (if available)
        if TRAY_AVAILABLE:
            self.minimize_btn = ttk.Button(button_frame, text="Minimize to Tray", 
                                          command=self.minimize_to_tray, width=15)
            self.minimize_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Exit button
        self.exit_btn = ttk.Button(button_frame, text="Exit", 
                                  command=self.quit_app, width=10)
        self.exit_btn.pack(side=tk.LEFT)
        
        # Info text
        info_text = tk.Text(main_frame, height=6, width=50, wrap=tk.WORD, 
                           font=('Arial', 9))
        info_text.pack(pady=(20, 0))
        info_text.insert(tk.END, 
            "This application runs a local web server for the ULLTRA study dashboard. "
            "Click 'Open Dashboard' to view enrollment statistics and study metrics in your web browser. "
            "\n\nThe dashboard automatically fetches data from REDCap and provides cached access for improved performance."
        )
        info_text.config(state=tk.DISABLED)
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        return self.root
    
    def update_status(self, status_text, url_text=""):
        """Update the GUI status labels"""
        if self.root and self.status_label:
            self.status_label.config(text=f"Server Status: {status_text}")
            self.url_label.config(text=url_text)
            self.root.update_idletasks()
    
    def show_window(self):
        """Show the main window"""
        if self.root:
            self.root.deiconify()
            self.root.lift()
    
    def minimize_to_tray(self):
        """Minimize to system tray"""
        if self.tray and self.root:
            self.root.withdraw()
    
    def on_window_close(self):
        """Handle window close event"""
        if TRAY_AVAILABLE and self.tray:
            # Minimize to tray instead of closing
            self.minimize_to_tray()
        else:
            # No tray support, actually quit
            self.quit_app()
    
    def quit_app(self):
        """Quit the application"""
        self.stop_server()
        
        if self.tray:
            self.tray.stop()
        
        if self.root:
            self.root.quit()
            self.root.destroy()
        
        sys.exit(0)
    
    def run(self):
        """Main application entry point"""
        print("Starting ULLTRA Dashboard...")
        
        # Start the web server
        if not self.start_server():
            print("Failed to start server. Exiting.")
            return
        
        self.update_status("Running", f"http://127.0.0.1:{self.port}")
        
        # Create system tray icon
        if TRAY_AVAILABLE:
            self.tray = self.create_tray_icon()
        
        # Create GUI if available
        if GUI_AVAILABLE:
            self.root = self.create_gui()
            
            # Auto-open dashboard after a short delay
            self.root.after(1000, self.open_dashboard)
            
            # Start tray icon in separate thread if available
            if self.tray:
                tray_thread = threading.Thread(target=self.tray.run, daemon=True)
                tray_thread.start()
            
            # Run GUI main loop
            try:
                self.root.mainloop()
            except KeyboardInterrupt:
                pass
        else:
            # No GUI, just run server and open browser
            self.open_dashboard()
            print("Press Ctrl+C to stop the server")
            try:
                while self.running:
                    threading.Event().wait(1)
            except KeyboardInterrupt:
                pass
        
        # Cleanup
        self.quit_app()

def main():
    """Application entry point"""
    try:
        app = ULLTRADashboard()
        app.run()
    except Exception as e:
        print(f"Application error: {e}")
        if GUI_AVAILABLE:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Application Error", f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()