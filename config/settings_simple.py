# config/settings_simple.py
"""
Simplified configuration management for backward compatibility.
"""
import os
from typing import Optional, List
from pathlib import Path
from dotenv import load_dotenv

# Load .env file
load_dotenv()


class Settings:
    """Simplified application settings."""
    
    def __init__(self):
        # API Keys
        self.openai_api_key: Optional[str] = os.getenv("OPENAI_API_KEY")
        self.pplx_api_key: Optional[str] = os.getenv("PPLX_API_KEY")
        self.anthropic_api_key: Optional[str] = os.getenv("ANTHROPIC_API_KEY")
        self.google_api_key: Optional[str] = os.getenv("GOOGLE_API_KEY")
        self.xai_api_key: Optional[str] = os.getenv("XAI_API_KEY")
        self.mistral_api_key: Optional[str] = os.getenv("MISTRAL_API_KEY")
        
        # Local/LLaMA endpoints
        self.llama_api_base: Optional[str] = os.getenv("LLAMA_API_BASE")
        self.llama_api_model: Optional[str] = os.getenv("LLAMA_API_MODEL")
        self.llama_api_key: Optional[str] = os.getenv("LLAMA_API_KEY")
        self.local_openai_base: Optional[str] = os.getenv("LOCAL_OPENAI_BASE")
        self.local_openai_model: Optional[str] = os.getenv("LOCAL_OPENAI_MODEL")
        self.local_openai_key: Optional[str] = os.getenv("OLLAMA_API_KEY")
        
        # HTTP Configuration
        self.http_timeout: float = float(os.getenv("HTTP_TIMEOUT", "60"))
        self.max_tokens: int = int(os.getenv("MAX_TOKENS", "2000"))
        
        # Behavior Flags
        self.force_perplexity_web: bool = os.getenv("FORCE_PERPLEXITY_WEB", "0") == "1"
        self.prefer_perplexity: bool = os.getenv("PREFER_PERPLEXITY", "0") == "1"
        self.always_synthesize_web: bool = os.getenv("ALWAYS_SYNTHESIZE_WEB", "0") == "1"
        self.web_scoring_threshold: int = int(os.getenv("WEB_SCORING_THRESHOLD", "3"))
        self.pplx_model: str = os.getenv("PPLX_MODEL", "sonar")
        
        # Email Configuration
        self.email_from: str = os.getenv("EMAIL_FROM", "")
        self.email_whitelist: List[str] = self._parse_list(os.getenv("EMAIL_WHITELIST", ""))
        self.smtp_host: str = os.getenv("SMTP_HOST", "")
        self.smtp_port: int = int(os.getenv("SMTP_PORT", "587"))
        self.smtp_user: str = os.getenv("SMTP_USER", "")
        self.smtp_pass: str = os.getenv("SMTP_PASS", "")
        self.smtp_debug: bool = os.getenv("SMTP_DEBUG", "0") == "1"
        self.company_address: str = os.getenv("COMPANY_ADDRESS", "")
        
        # Security
        self.dispatch_secret: str = os.getenv("DISPATCH_SECRET", "dev-secret")
        self.allowed_numbers: List[str] = self._parse_list(os.getenv("ALLOWED_NUMBERS", ""))
        
        # File Paths
        self.data_dir: Path = Path(os.getenv("DATA_DIR", "."))
        self.cloud_out_dir: Path = Path(os.getenv("CLOUD_OUT_DIR", "~/Documents/JarvisCloud")).expanduser()
        self.sheets_dir: Path = Path(os.getenv("SHEETS_DIR", "~/Documents/JarvisSheets")).expanduser()
        self.codes_db: Path = Path(os.getenv("CODES_DB", "~/jarvis-demo/commands.json")).expanduser()
        
        # Gmail/Google Voice
        self.gv_client_json: str = os.getenv("GV_CLIENT_JSON", "client_secret.json")
        self.gv_token_json: str = os.getenv("GV_TOKEN_JSON", "token.json")
        self.gmail_query: str = os.getenv(
            "GMAIL_QUERY",
            'in:inbox is:unread newer_than:7d (from:@txt.voice.google.com OR (from:voice-noreply@google.com subject:"New message"))'
        )
        self.poll_seconds: int = int(os.getenv("POLL_SECONDS", "20"))
        
        # Debug/Development
        self.debug_log: bool = os.getenv("DEBUG_LOG", "0") == "1"
        self.greet_force: bool = os.getenv("GREET_FORCE", "0") == "1"
        self.greet_cooldown_seconds: int = int(os.getenv("GREET_COOLDOWN_SECONDS", "300"))
        self.node_name: str = os.getenv("NODE_NAME", "Echo-Nine")
        
        # SEC Configuration
        self.sec_app_name: str = os.getenv("SEC_APP_NAME", "Williams Cloud Control")
        self.sec_contact_email: str = os.getenv("SEC_CONTACT_EMAIL", "example@example.com")
        
        # Timezone
        self.app_timezone: str = os.getenv("APP_TIMEZONE", "America/Chicago")
        
        # Agent URLs
        self.mac_agent_url: str = os.getenv("MAC_AGENT_URL", "http://127.0.0.1:8787")
        self.win_agent_url: Optional[str] = os.getenv("WIN_AGENT_URL")
        self.public_base_url: str = os.getenv("PUBLIC_BASE_URL", "http://127.0.0.1:8000")
        
        # Ensure directories exist
        self.cloud_out_dir.mkdir(parents=True, exist_ok=True)
        self.sheets_dir.mkdir(parents=True, exist_ok=True)
        self.data_dir.mkdir(parents=True, exist_ok=True)
    
    def _parse_list(self, value: str) -> List[str]:
        """Parse comma-separated string into list."""
        if not value or not value.strip():
            return []
        return [item.strip() for item in value.split(",") if item.strip()]


# Global settings instance
_settings: Optional[Settings] = None


def get_settings() -> Settings:
    """Get the global settings instance, creating it if necessary."""
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings


def reload_settings() -> Settings:
    """Reload settings from environment variables."""
    global _settings
    _settings = Settings()
    return _settings
