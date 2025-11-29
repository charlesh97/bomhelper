"""Configuration management for API keys and settings."""
import os
import re
from pathlib import Path


class Config:
    """Manages configuration settings and API keys."""
    
    def __init__(self, keys_file="keys.txt"):
        """Initialize configuration from keys file."""
        self.keys_file = Path(keys_file)
        self.mouser_api_key = None
        self.gemini_api_key = None
        self.load_keys()
    
    def load_keys(self):
        """Load API keys from keys.txt file."""
        if not self.keys_file.exists():
            # Try environment variables as fallback
            self.mouser_api_key = os.getenv("MOUSER_API_KEY")
            self.gemini_api_key = os.getenv("GEMINI_API_KEY")
            return
        
        try:
            with open(self.keys_file, 'r') as f:
                content = f.read()
                
            # Parse Mouser API key
            mouser_match = re.search(r'MouserAPIkey=([^\s\n]+)', content)
            if mouser_match:
                self.mouser_api_key = mouser_match.group(1).strip()
            else:
                self.mouser_api_key = os.getenv("MOUSER_API_KEY")
            
            # Parse Gemini API key
            gemini_match = re.search(r'GeminiKey=([^\s\n]+)', content)
            if gemini_match:
                self.gemini_api_key = gemini_match.group(1).strip()
            else:
                self.gemini_api_key = os.getenv("GEMINI_API_KEY")
                
        except Exception as e:
            print(f"Warning: Could not load keys from {self.keys_file}: {e}")
            # Fallback to environment variables
            self.mouser_api_key = os.getenv("MOUSER_API_KEY")
            self.gemini_api_key = os.getenv("GEMINI_API_KEY")
    
    def get_mouser_api_key(self):
        """Get Mouser API key."""
        return self.mouser_api_key
    
    def get_gemini_api_key(self):
        """Get Gemini API key."""
        return self.gemini_api_key
    
    def is_configured(self):
        """Check if both API keys are configured."""
        return self.mouser_api_key is not None and self.gemini_api_key is not None

