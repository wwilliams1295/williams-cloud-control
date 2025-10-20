# config/settings.py
"""
Centralized configuration management with validation and type safety.
"""
import os
from typing import Optional, List, Dict, Any
from pydantic import Field, validator
from pydantic_settings import BaseSettings
from pathlib import Path


class Settings(BaseSettings):
    """Application settings with validation and environment variable support."""
    
    # API Keys
    openai_api_key: Optional[str] = Field(None, env="OPENAI_API_KEY")
    pplx_api_key: Optional[str] = Field(None, env="PPLX_API_KEY")
    anthropic_api_key: Optional[str] = Field(None, env="ANTHROPIC_API_KEY")
    google_api_key: Optional[str] = Field(None, env="GOOGLE_API_KEY")
    xai_api_key: Optional[str] = Field(None, env="XAI_API_KEY")
    mistral_api_key: Optional[str] = Field(None, env="MISTRAL_API_KEY")
    
    # Local/LLaMA endpoints
    llama_api_base: Optional[str] = Field(None, env="LLAMA_API_BASE")
    llama_api_model: Optional[str] = Field(None, env="LLAMA_API_MODEL")
    llama_api_key: Optional[str] = Field(None, env="LLAMA_API_KEY")
    local_openai_base: Optional[str] = Field(None, env="LOCAL_OPENAI_BASE")
    local_openai_model: Optional[str] = Field(None, env="LOCAL_OPENAI_MODEL")
    local_openai_key: Optional[str] = Field(None, env="OLLAMA_API_KEY")
    
    # HTTP Configuration
    http_timeout: float = Field(60.0, env="HTTP_TIMEOUT", ge=1.0, le=300.0)
    max_tokens: int = Field(2000, env="MAX_TOKENS", ge=100, le=8000)
    
    # Behavior Flags
    force_perplexity_web: bool = Field(False, env="FORCE_PERPLEXITY_WEB")
    prefer_perplexity: bool = Field(False, env="PREFER_PERPLEXITY")
    always_synthesize_web: bool = Field(False, env="ALWAYS_SYNTHESIZE_WEB")
    web_scoring_threshold: int = Field(3, env="WEB_SCORING_THRESHOLD", ge=1, le=10)
    pplx_model: str = Field("sonar", env="PPLX_MODEL")
    
    # Email Configuration
    email_from: str = Field("", env="EMAIL_FROM")
    email_whitelist: List[str] = Field(default_factory=list, env="EMAIL_WHITELIST")
    smtp_host: str = Field("", env="SMTP_HOST")
    smtp_port: int = Field(587, env="SMTP_PORT", ge=1, le=65535)
    smtp_user: str = Field("", env="SMTP_USER")
    smtp_pass: str = Field("", env="SMTP_PASS")
    smtp_debug: bool = Field(False, env="SMTP_DEBUG")
    company_address: str = Field("", env="COMPANY_ADDRESS")
    
    # Security
    dispatch_secret: str = Field("dev-secret", env="DISPATCH_SECRET")
    allowed_numbers: List[str] = Field(default_factory=list, env="ALLOWED_NUMBERS")
    
    # File Paths
    data_dir: Path = Field(Path("."), env="DATA_DIR")
    cloud_out_dir: Path = Field(Path("~/Documents/JarvisCloud").expanduser(), env="CLOUD_OUT_DIR")
    sheets_dir: Path = Field(Path("~/Documents/JarvisSheets").expanduser(), env="SHEETS_DIR")
    codes_db: Path = Field(Path("~/jarvis-demo/commands.json").expanduser(), env="CODES_DB")
    
    # Gmail/Google Voice
    gv_client_json: str = Field("client_secret.json", env="GV_CLIENT_JSON")
    gv_token_json: str = Field("token.json", env="GV_TOKEN_JSON")
    gmail_query: str = Field(
        'in:inbox is:unread newer_than:7d (from:@txt.voice.google.com OR (from:voice-noreply@google.com subject:"New message"))',
        env="GMAIL_QUERY"
    )
    poll_seconds: int = Field(20, env="POLL_SECONDS", ge=5, le=300)
    
    # Debug/Development
    debug_log: bool = Field(False, env="DEBUG_LOG")
    greet_force: bool = Field(False, env="GREET_FORCE")
    greet_cooldown_seconds: int = Field(300, env="GREET_COOLDOWN_SECONDS", ge=60, le=3600)
    node_name: str = Field("Echo-Nine", env="NODE_NAME")
    
    # SEC Configuration
    sec_app_name: str = Field("Williams Cloud Control", env="SEC_APP_NAME")
    sec_contact_email: str = Field("example@example.com", env="SEC_CONTACT_EMAIL")
    
    # Timezone
    app_timezone: str = Field("America/Chicago", env="APP_TIMEZONE")
    
    # Agent URLs
    mac_agent_url: str = Field("http://127.0.0.1:8787", env="MAC_AGENT_URL")
    win_agent_url: Optional[str] = Field(None, env="WIN_AGENT_URL")
    public_base_url: str = Field("http://127.0.0.1:8000", env="PUBLIC_BASE_URL")
    
    @validator('email_whitelist', pre=True)
    def parse_email_whitelist(cls, v):
        if isinstance(v, str):
            if not v.strip():
                return []
            return [email.strip().lower() for email in v.split(",") if email.strip()]
        if v is None:
            return []
        return v
    
    @validator('allowed_numbers', pre=True)
    def parse_allowed_numbers(cls, v):
        if isinstance(v, str):
            if not v.strip():
                return []
            return [num.strip() for num in v.split(",") if num.strip()]
        if v is None:
            return []
        return v
    
    @validator('data_dir', 'cloud_out_dir', 'sheets_dir', 'codes_db', pre=True)
    def ensure_paths_exist(cls, v):
        if isinstance(v, str):
            path = Path(v).expanduser()
        else:
            path = v
        path.mkdir(parents=True, exist_ok=True)
        return path
    
    class Config:
        env_file = ".env"
        case_sensitive = False
        validate_assignment = True


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
