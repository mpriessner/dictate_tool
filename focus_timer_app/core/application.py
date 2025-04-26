#!/usr/bin/env python3
# application.py - Application detection and URL tracking

import subprocess

class ApplicationTracker:
    """Tracks active applications and browser URLs"""
    
    def __init__(self):
        self.supported_browsers = ["Google Chrome", "Safari", "Arc"]
    
    def get_active_application(self):
        """Get the name of the currently active application on macOS
        
        Returns:
            str: Name of the active application or "Unknown"
        """
        try:
            # Use AppleScript to get the name of the frontmost application
            cmd = ["osascript", "-e", 
                   """
                   tell application "System Events"
                       set frontApp to first application process whose frontmost is true
                       set frontAppName to name of frontApp
                       return frontAppName
                   end tell
                   """]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1)
            app_name = result.stdout.strip()
            return app_name if app_name else "Unknown"
        except (subprocess.SubprocessError, subprocess.TimeoutExpired):
            return "Unknown"
    
    def get_browser_url(self, browser_name):
        """Get the URL of the active tab in a browser
        
        Args:
            browser_name (str): Name of the browser
            
        Returns:
            str: Main domain of the URL or "n/a"
        """
        try:
            # Different AppleScript commands for different browsers
            if browser_name == "Google Chrome":
                script = """
                tell application "Google Chrome"
                    set currentURL to URL of active tab of front window
                    return currentURL
                end tell
                """
            elif browser_name == "Arc":
                script = """
                tell application "Arc"
                    set currentURL to URL of active tab of front window
                    return currentURL
                end tell
                """
            elif browser_name == "Safari":
                script = """
                tell application "Safari"
                    set currentURL to URL of current tab of front window
                    return currentURL
                end tell
                """
            else:
                return "n/a"  # Not a supported browser
                
            cmd = ["osascript", "-e", script]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=1)
            url = result.stdout.strip()
            
            # Extract just the main domain from the URL
            if url:
                return self.extract_main_domain(url)
            return "n/a"
        except (subprocess.SubprocessError, subprocess.TimeoutExpired):
            return "n/a"
    
    def extract_main_domain(self, url):
        """Extract just the main domain from a URL (up to the TLD)
        
        Args:
            url (str): Full URL
            
        Returns:
            str: Main domain
        """
        try:
            # Remove protocol (http://, https://, etc.)
            if "://" in url:
                url = url.split("://", 1)[1]
            
            # Remove path, query parameters, etc.
            if "/" in url:
                url = url.split("/", 1)[0]
            
            # Remove port if present
            if ":" in url:
                url = url.split(":", 1)[0]
            
            # Handle common subdomains
            parts = url.split(".")
            if len(parts) > 2:
                # Check if it's a known subdomain pattern like www.example.com
                if parts[0] == "www":
                    # Return example.com
                    return ".".join(parts[1:])
                
                # For other subdomains, try to identify the main domain + TLD
                # This is a simplified approach - for complex TLDs like co.uk, this would need refinement
                return ".".join(parts[-2:])
            
            return url
        except Exception:
            # If any parsing error occurs, return the original URL
            return url
