"""Authentication module for ManageBac Gradebook Scraper."""

import time
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

from .exceptions import AuthenticationError, SessionExpiredError


class Authenticator:
    """Handles ManageBac authentication and CSRF token extraction."""
    
    def __init__(self, school_code: str, email: str, password: str):
        """Initialize with ManageBac credentials.
        
        Args:
            school_code: The school's ManageBac subdomain code
            email: User's email address
            password: User's password
        """
        self.school_code = school_code
        self.email = email
        self.password = password
        self.base_url = f"https://{school_code}.managebac.cn"
        self._session: requests.Session | None = None
    
    def login(self) -> requests.Session:
        """Authenticate with ManageBac and return authenticated session.
        
        Returns:
            requests.Session: Authenticated session for subsequent requests
            
        Raises:
            AuthenticationError: If authentication fails
        """
        tqdm.write(f"Logging in to {self.base_url}...")
        
        try:
            # Use a session to maintain cookies across requests
            session = requests.Session()
            
            # Add common headers to mimic a browser
            session.headers.update({
                "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
            })
            
            # Get login page to extract CSRF token
            login_url = f"{self.base_url}/login"
            tqdm.write(f"Fetching login page: {login_url}")
            response = session.get(login_url, timeout=30)
            response.raise_for_status()
            
            csrf_token = self._extract_csrf_token(response.text)
            if not csrf_token:
                raise AuthenticationError(
                    "Failed to extract CSRF token from login page. "
                    "The login page structure may have changed."
                )
            
            tqdm.write(f"CSRF token extracted successfully")
            
            # Find the login form action URL
            soup = BeautifulSoup(response.text, "lxml")
            form = soup.find("form")
            if form and form.get("action"):
                form_action = form.get("action")
                if form_action.startswith("/"):
                    session_url = f"{self.base_url}{form_action}"
                elif form_action.startswith("http"):
                    session_url = form_action
                else:
                    session_url = f"{self.base_url}/{form_action}"
            else:
                # Default to /session
                session_url = f"{self.base_url}/session"
            
            tqdm.write(f"Form action URL: {session_url}")

            # Build payload from form inputs
            payload = {
                "authenticity_token": csrf_token,
            }
            
            # Find all input fields in the form
            if form:
                for input_field in form.find_all("input"):
                    name = input_field.get("name")
                    if name and name != "authenticity_token":
                        if name in ["login", "email", "user[email]", "session[email]"]:
                            payload[name] = self.email
                        elif name in ["password", "user[password]", "session[password]"]:
                            payload[name] = self.password
                        elif input_field.get("value"):
                            payload[name] = input_field.get("value")
            
            # Fallback if form parsing didn't find email/password fields
            if not any(k for k in payload.keys() if "email" in k.lower() or k == "login"):
                payload["login"] = self.email
            if not any(k for k in payload.keys() if "password" in k.lower()):
                payload["password"] = self.password
            
            tqdm.write(f"Payload fields: {list(payload.keys())}")
            
            headers = {
                "Content-Type": "application/x-www-form-urlencoded",
                "Referer": login_url,
                "Origin": self.base_url,
            }
            
            tqdm.write(f"Submitting credentials to: {session_url}")
            login_response = session.post(
                session_url,
                data=payload,
                headers=headers,
                allow_redirects=True,
                timeout=30
            )
            
            tqdm.write(f"Login response status: {login_response.status_code}")
            tqdm.write(f"Final URL: {login_response.url}")
            
            # Check for 404 - wrong endpoint
            if login_response.status_code == 404:
                raise AuthenticationError(
                    f"Login endpoint not found (404). The URL {session_url} may be incorrect."
                )
            
            # Check if we ended up on a dashboard/home page (successful login)
            # or still on login page (failed)
            if "/login" not in login_response.url:
                tqdm.write(f"Login successful! Got {len(session.cookies)} cookies")
                self._session = session
                return session
            
            # Check response for error messages
            error_msg = self._extract_error_message(login_response.text)
            tqdm.write(f"Login failed - still on login page")
            raise AuthenticationError(
                f"Authentication failed: {error_msg or 'Invalid credentials or unexpected response'}"
            )
            
        except requests.RequestException as e:
            raise AuthenticationError(
                f"Network error during authentication: {str(e)}"
            ) from e
    
    def _extract_csrf_token(self, html: str) -> str | None:
        """Extract CSRF token from login page meta tag.
        
        Args:
            html: HTML content of the login page
            
        Returns:
            The CSRF token string, or None if not found
        """
        soup = BeautifulSoup(html, "lxml")
        meta_tag = soup.find("meta", attrs={"name": "csrf-token"})
        if meta_tag and meta_tag.get("content"):
            content = meta_tag["content"]
            if isinstance(content, str):
                return content
        return None
    
    def _extract_error_message(self, html: str) -> str | None:
        """Extract error message from login response.
        
        Args:
            html: HTML content of the response
            
        Returns:
            Error message string, or None if not found
        """
        soup = BeautifulSoup(html, "lxml")
        # Look for common error message containers
        error_div = soup.find("div", class_="alert-danger")
        if error_div:
            return error_div.get_text(strip=True)
        error_div = soup.find("div", class_="error")
        if error_div:
            return error_div.get_text(strip=True)
        return None



class SessionManager:
    """Manages authenticated sessions with retry logic and auto-reauthentication."""
    
    def __init__(self, session: requests.Session, authenticator: Authenticator):
        """Initialize with authenticated session and authenticator reference.
        
        Args:
            session: Authenticated requests.Session object
            authenticator: Authenticator instance for re-authentication
        """
        self.session: requests.Session = session
        self.authenticator: Authenticator = authenticator
    
    def get(self, url: str, retry_count: int = 3) -> requests.Response:
        """Make authenticated GET request with retry logic.
        
        Implements exponential backoff and automatic re-authentication
        on session expiration.
        
        Args:
            url: URL to request
            retry_count: Maximum number of retry attempts (default: 3)
            
        Returns:
            requests.Response: The response object
            
        Raises:
            requests.RequestException: If all retries fail
            SessionExpiredError: If re-authentication fails
        """
        last_exception = None
        
        for attempt in range(retry_count):
            try:
                # Add common headers
                headers = {
                    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                    "Accept-Language": "en-US,en;q=0.9",
                }
                
                tqdm.write(f"Requesting: {url}")
                tqdm.write(f"Session cookies: {list(self.session.cookies.keys())}")
                
                response = self.session.get(url, headers=headers, timeout=30)
                
                tqdm.write(f"Response status: {response.status_code}")
                
                # Check for session expiration (redirect to login page)
                if self._is_session_expired(response):
                    tqdm.write("Session expired, re-authenticating...")
                    try:
                        self.session = self.authenticator.login()
                        # Retry the request with new session
                        response = self.session.get(url, headers=headers, timeout=30)
                        if self._is_session_expired(response):
                            raise SessionExpiredError(
                                "Re-authentication succeeded but session still appears expired"
                            )
                    except AuthenticationError as e:
                        raise SessionExpiredError(
                            f"Failed to re-authenticate: {str(e)}"
                        ) from e
                
                response.raise_for_status()
                return response
                
            except requests.RequestException as e:
                last_exception = e
                if attempt < retry_count - 1:
                    wait_time = 2 ** attempt  # Exponential backoff: 1, 2, 4 seconds
                    tqdm.write(f"Request failed, retrying in {wait_time}s... ({attempt + 1}/{retry_count})")
                    time.sleep(wait_time)
        
        # All retries exhausted
        if last_exception is not None:
            raise last_exception
        raise requests.RequestException("All retry attempts failed")
    
    def _is_session_expired(self, response: requests.Response) -> bool:
        """Check if the response indicates session expiration.
        
        Args:
            response: The HTTP response to check
            
        Returns:
            True if session appears expired, False otherwise
        """
        # If we got a 200 OK, the session is likely fine
        if response.status_code == 200:
            # Only check for login redirect if the URL actually changed to login
            if response.url.rstrip('/').endswith('/login'):
                return True
            return False
        
        # Check for redirect to login page
        if response.history:
            for r in response.history:
                location = r.headers.get("Location", "")
                if location.rstrip('/').endswith('/login'):
                    return True
        
        # Check if current URL is login page
        if response.url.rstrip('/').endswith('/login'):
            return True
        
        return False
