"""
SharePoint API Handler for ULLTRA Dashboard
Uses OAuth2 Device Code Flow for authentication
"""

import json
import threading
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Any

try:
    from office365.sharepoint.client_context import ClientContext
    import msal
    OFFICE365_AVAILABLE = True
except ImportError:
    OFFICE365_AVAILABLE = False
    print("Warning: Office365-REST-Python-Client not installed. SharePoint integration disabled.")


class SharePointManager:
    """Manages SharePoint authentication and data access"""

    # Microsoft's public Office client ID - no registration required
    CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    TENANT_ID = "common"  # Works across all Microsoft 365 tenants

    def __init__(self, site_url: str, list_name: str):
        self.site_url = site_url
        self.list_name = list_name
        self.ctx: Optional[ClientContext] = None
        self.auth_status = {
            'authenticated': False,
            'device_code': None,
            'user_code': None,
            'verification_url': None,
            'message': None,
            'expires_in': None,
            'error': None
        }
        self.access_token = None
        self.token_expires_at = None

    def start_device_code_flow(self) -> Dict[str, Any]:
        """
        Initiate device code flow authentication.
        Returns device code info for user to complete authentication.
        """
        if not OFFICE365_AVAILABLE:
            return {
                'success': False,
                'error': 'Office365-REST-Python-Client library not installed. Run: pip install Office365-REST-Python-Client'
            }

        try:
            # Use MSAL directly to get device code
            app = msal.PublicClientApplication(
                client_id=self.CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{self.TENANT_ID}"
            )

            # Initiate device flow - this returns the device code info
            flow = app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])

            if "user_code" not in flow:
                raise Exception("Failed to create device flow")

            # Store device code info
            self.auth_status['user_code'] = flow['user_code']
            self.auth_status['device_code'] = flow['device_code']
            self.auth_status['verification_url'] = flow['verification_uri']
            self.auth_status['message'] = flow['message']
            self.auth_status['expires_in'] = flow.get('expires_in', 900)

            # Start authentication in background thread
            def authenticate():
                try:
                    # Wait for user to authenticate
                    result = app.acquire_token_by_device_flow(flow)

                    if "access_token" in result:
                        # Now use the token with Office365 library
                        self.access_token = result['access_token']

                        # Create SharePoint context with the token
                        self.ctx = ClientContext(self.site_url)
                        self.ctx.with_access_token(lambda: result['access_token'])

                        # Test connection
                        web = self.ctx.web
                        self.ctx.load(web)
                        self.ctx.execute_query()

                        self.auth_status['authenticated'] = True
                        self.auth_status['message'] = f'Successfully connected to: {web.properties.get("Title", "SharePoint")}'
                        print(f"[SharePoint] Authentication successful")
                    else:
                        error_msg = result.get('error_description', 'Authentication failed')
                        self.auth_status['authenticated'] = False
                        self.auth_status['error'] = error_msg
                        print(f"[SharePoint] Authentication failed: {error_msg}")

                except Exception as e:
                    self.auth_status['authenticated'] = False
                    self.auth_status['error'] = str(e)
                    print(f"[SharePoint] Authentication failed: {e}")

            # Start authentication in background
            auth_thread = threading.Thread(target=authenticate, daemon=True)
            auth_thread.start()

            return {
                'success': True,
                'user_code': flow['user_code'],
                'device_code': flow['device_code'],
                'verification_url': flow['verification_uri'],
                'expires_in': flow.get('expires_in', 900),
                'message': flow['message'],
                'instructions': [
                    f"Go to: {flow['verification_uri']}",
                    f"Enter the code: {flow['user_code']}",
                    'Sign in with your Microsoft 365 credentials',
                    'Grant permissions when prompted',
                    'Return here - authentication will complete automatically'
                ]
            }

        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }

    def get_auth_status(self) -> Dict[str, Any]:
        """Get current authentication status"""
        return {
            'authenticated': self.auth_status['authenticated'],
            'message': self.auth_status['message'],
            'error': self.auth_status['error']
        }

    def is_authenticated(self) -> bool:
        """Check if currently authenticated"""
        return self.ctx is not None and self.auth_status['authenticated']

    def logout(self):
        """Clear authentication"""
        self.ctx = None
        self.access_token = None
        self.token_expires_at = None
        self.auth_status = {
            'authenticated': False,
            'device_code': None,
            'user_code': None,
            'verification_url': None,
            'message': None,
            'expires_in': None,
            'error': None
        }

    def get_calendar_events(self) -> List[Dict[str, Any]]:
        """
        Fetch calendar events from SharePoint list
        Returns list of events with standardized field names
        """
        if not self.is_authenticated():
            raise Exception("Not authenticated. Please authenticate first.")

        try:
            # Get the list
            target_list = self.ctx.web.lists.get_by_title(self.list_name)

            # Fetch list items
            items = target_list.items
            self.ctx.load(items)
            self.ctx.execute_query()

            # Transform items to calendar events
            events = []
            for item in items:
                try:
                    event = self._transform_list_item_to_event(item)
                    if event:
                        events.append(event)
                except Exception as e:
                    print(f"Error transforming item: {e}")
                    continue

            return events

        except Exception as e:
            raise Exception(f"Failed to fetch calendar events: {str(e)}")

    def _transform_list_item_to_event(self, item) -> Optional[Dict[str, Any]]:
        """
        Transform SharePoint list item to standardized event format
        Adjust field mappings based on your actual SharePoint list structure
        """
        try:
            properties = item.properties

            # Common SharePoint list field mappings
            # Adjust these based on your actual column names
            event = {
                'id': properties.get('ID'),
                'title': properties.get('Title', ''),
                'participant': properties.get('Participant') or properties.get('ParticipantID') or properties.get('Subject', ''),
                'date': properties.get('EventDate') or properties.get('StartDate') or properties.get('Date'),
                'time': properties.get('EventTime') or properties.get('Time', ''),
                'type': properties.get('Category') or properties.get('EventType') or properties.get('Type', 'general'),
                'description': properties.get('Description') or properties.get('Notes') or properties.get('Body', ''),
                'location': properties.get('Location', ''),
                'status': properties.get('Status', ''),
                # Keep all raw properties for debugging
                'raw': properties
            }

            return event

        except Exception as e:
            print(f"Error transforming item: {e}")
            return None


# Global SharePoint manager instance
_sharepoint_manager: Optional[SharePointManager] = None


def get_sharepoint_manager(site_url: str = None, list_name: str = None) -> SharePointManager:
    """Get or create SharePoint manager instance"""
    global _sharepoint_manager

    if _sharepoint_manager is None:
        if not site_url or not list_name:
            raise ValueError("SharePoint site URL and list name required")
        _sharepoint_manager = SharePointManager(site_url, list_name)

    return _sharepoint_manager
