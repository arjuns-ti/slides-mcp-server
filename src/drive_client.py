#!/usr/bin/env python3
"""
Google Drive and Slides API client with OAuth authentication.
"""
import os
import sys
import socket
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

# Google Drive and Slides API scopes
SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/presentations'  # Read and write access to presentations
]


def log_message(message: str):
    """
    Log a message to logs.txt if logging is enabled.
    
    Args:
        message: The message to log
    """
    enable_logging = os.getenv('ENABLE_LOGGING', 'false').lower() == 'true'
    if not enable_logging:
        return
    
    try:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        
        with open('logs.txt', 'a') as f:
            f.write(log_entry)
    except Exception:
        # Silently fail if logging fails
        pass


@contextmanager
def suppress_output():
    """
    Context manager to suppress all stdout and stderr output.
    Redirects everything to devnull to prevent console output.
    """
    devnull = open(os.devnull, 'w')
    old_stdout = sys.stdout
    old_stderr = sys.stderr
    try:
        sys.stdout = devnull
        sys.stderr = devnull
        yield
    finally:
        sys.stdout = old_stdout
        sys.stderr = old_stderr
        devnull.close()


def get_drive_service():
    """
    Get authenticated Google Drive service.
    
    Returns:
        Google Drive service object with credentials attached
        
    Raises:
        ValueError: If credentials are not configured properly
    """
    oauth_client_secret = os.getenv('OAUTH_CLIENT_SECRET')
    oauth_client_token = os.getenv('OAUTH_CLIENT_TOKEN')
    
    log_message("Attempting to get Google Drive service")
    
    if not oauth_client_secret:
        log_message("ERROR: OAUTH_CLIENT_SECRET not set")
        raise ValueError("OAUTH_CLIENT_SECRET environment variable not set")
    
    if not oauth_client_token:
        log_message("ERROR: OAUTH_CLIENT_TOKEN not set")
        raise ValueError("OAUTH_CLIENT_TOKEN environment variable not set")
    
    creds = None
    token_path = Path(oauth_client_token)
    
    # Load existing token
    if token_path.exists():
        log_message(f"Token file found at {token_path}")
        log_message(f"Loading token from {token_path}")
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    else:
        log_message(f"Token file not found at {token_path}")
        log_message("Will need to authenticate with Google OAuth")
    
    # Refresh or get new credentials
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            log_message("Refreshing expired credentials")
            with suppress_output():
                creds.refresh(Request())
        else:
            log_message("Getting new credentials via OAuth flow")
            log_message("Token file not present, initiating OAuth authentication")
            
            # Find an available port for the OAuth callback
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.bind(('localhost', 0))
            port = sock.getsockname()[1]
            sock.close()
            
            log_message(f"Starting local OAuth server on port {port}")
            
            try:
                from http.server import HTTPServer, BaseHTTPRequestHandler
                from urllib.parse import urlparse, parse_qs
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    oauth_client_secret, SCOPES)
                
                flow.redirect_uri = f'http://localhost:{port}/'
                
                # Step 1: Generate and log the authorization URL
                auth_url, state = flow.authorization_url(
                    access_type='offline',
                    include_granted_scopes='true',
                    prompt='consent'
                )
                
                log_message("=" * 80)
                log_message("AUTHORIZATION REQUIRED")
                log_message("=" * 80)
                log_message(f"Please open this URL in your browser:")
                log_message("")
                log_message(auth_url)
                log_message("")
                log_message("=" * 80)
                
                # Step 2: Create simple HTTP server to receive callback
                log_message(f"Waiting for OAuth callback on port {port}...")
                
                callback_data = {'code': None, 'error': None}
                
                class CallbackHandler(BaseHTTPRequestHandler):
                    def log_message(self, format, *args):
                        # Suppress HTTP server logs
                        pass
                    
                    def do_GET(self):
                        query_params = parse_qs(urlparse(self.path).query)
                        
                        if 'code' in query_params:
                            callback_data['code'] = query_params['code'][0]
                            callback_data['state'] = query_params.get('state', [None])[0]
                            self.send_response(200)
                            self.send_header('Content-type', 'text/html')
                            self.end_headers()
                            self.wfile.write(b'<html><body><h1>Authentication successful!</h1><p>You can close this window.</p></body></html>')
                        elif 'error' in query_params:
                            callback_data['error'] = query_params['error'][0]
                            self.send_response(400)
                            self.send_header('Content-type', 'text/html')
                            self.end_headers()
                            self.wfile.write(b'<html><body><h1>Authentication failed</h1></body></html>')
                        else:
                            self.send_response(400)
                            self.end_headers()
                
                server = HTTPServer(('localhost', port), CallbackHandler)
                
                # Handle one request then shut down
                server.handle_request()
                
                if callback_data['error']:
                    raise Exception(f"OAuth error: {callback_data['error']}")
                
                if not callback_data['code']:
                    raise Exception("No authorization code received")
                
                # Verify state matches
                if callback_data.get('state') != state:
                    raise Exception("State mismatch - possible CSRF attack")
                
                log_message("OAuth callback received successfully")
                log_message("Exchanging authorization code for tokens...")
                
                # Step 4: Exchange code for tokens
                flow.fetch_token(code=callback_data['code'])
                creds = flow.credentials
                
                log_message("Access token obtained successfully")
            except Exception as e:
                log_message(f"OAuth error: {str(e)}")
                raise
        
        # Save the credentials to the token file
        log_message(f"Saving credentials to {token_path}")
        token_path.parent.mkdir(parents=True, exist_ok=True)
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
        log_message(f"Token saved successfully to {token_path}")
    
    log_message("Successfully authenticated with Google Drive and Slides APIs")
    with suppress_output():
        service = build('drive', 'v3', credentials=creds)
    
    # Store credentials for later use by Slides API
    service._credentials = creds
    return service


def get_slides_service(credentials):
    """
    Get authenticated Google Slides service.
    
    Args:
        credentials: OAuth credentials
        
    Returns:
        Google Slides service object
    """
    with suppress_output():
        return build('slides', 'v1', credentials=credentials)

