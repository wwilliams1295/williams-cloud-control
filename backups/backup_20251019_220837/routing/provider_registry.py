# routing/provider_registry.py
"""
Provider registry for managing available LLM providers.
"""
from typing import Dict, List, Optional, Type
from dataclasses import dataclass
from config.settings_simple import get_settings

from providers.base import BaseAPIClient
from providers.openai import OpenAIClient
from providers.perplexity import PerplexityClient
from providers.anthropic import AnthropicClient
from providers.gemini import GeminiClient
from providers.grok import GrokClient
from providers.mistral import MistralClient


@dataclass
class ProviderInfo:
    """Information about a registered provider."""
    name: str
    client_class: Type[BaseAPIClient]
    default_model: str
    label: str
    client: Optional[BaseAPIClient] = None


class ProviderRegistry:
    """Registry for managing available LLM providers."""
    
    def __init__(self):
        self._providers: Dict[str, ProviderInfo] = {}
        self._initialize_providers()
    
    def _initialize_providers(self):
        """Initialize all available providers based on configuration."""
        settings = get_settings()
        
        # OpenAI
        if settings.openai_api_key:
            self.register(
                "openai",
                OpenAIClient,
                "gpt-4o-mini",
                "openai",
                api_key=settings.openai_api_key
            )
        
        # Perplexity
        if settings.pplx_api_key:
            self.register(
                "perplexity",
                PerplexityClient,
                settings.pplx_model,
                "perplexity",
                api_key=settings.pplx_api_key
            )
        
        # Anthropic
        if settings.anthropic_api_key:
            self.register(
                "anthropic",
                AnthropicClient,
                "claude-sonnet-4-5-20250929",
                "anthropic",
                api_key=settings.anthropic_api_key
            )
        
        # Google Gemini
        if settings.google_api_key:
            self.register(
                "gemini",
                GeminiClient,
                "gemini-2.5-flash",
                "google",
                api_key=settings.google_api_key
            )
        
        # xAI Grok
        if settings.xai_api_key:
            self.register(
                "grok",
                GrokClient,
                "grok-2-latest",
                "xai",
                api_key=settings.xai_api_key
            )
        
        # Mistral
        if settings.mistral_api_key:
            self.register(
                "mistral",
                MistralClient,
                "mistral-large-latest",
                "mistral",
                api_key=settings.mistral_api_key
            )
        
        # LLaMA (OpenAI-compatible)
        if settings.llama_api_base and settings.llama_api_model:
            from providers.openai_compatible import OpenAICompatibleClient
            self.register(
                "llama",
                OpenAICompatibleClient,
                settings.llama_api_model,
                "llama",
                api_key=settings.llama_api_key,
                base_url=settings.llama_api_base
            )
        
        # Local OpenAI-compatible
        if settings.local_openai_base and settings.local_openai_model:
            from providers.openai_compatible import OpenAICompatibleClient
            self.register(
                "local",
                OpenAICompatibleClient,
                settings.local_openai_model,
                "local",
                api_key=settings.local_openai_key,
                base_url=settings.local_openai_base
            )
    
    def register(
        self,
        name: str,
        client_class: Type[BaseAPIClient],
        default_model: str,
        label: str,
        **client_kwargs
    ) -> None:
        """Register a provider with the registry."""
        try:
            client = client_class(**client_kwargs)
            provider_info = ProviderInfo(
                name=name,
                client_class=client_class,
                default_model=default_model,
                label=label,
                client=client
            )
            self._providers[name] = provider_info
        except Exception as e:
            # Log error but don't fail registration
            import logging
            logger = logging.getLogger(__name__)
            logger.warning(f"Failed to register provider {name}: {e}")
    
    def get_provider(self, name: str) -> Optional[BaseAPIClient]:
        """Get a provider client by name."""
        provider_info = self._providers.get(name)
        return provider_info.client if provider_info else None
    
    def get_provider_info(self, name: str) -> Optional[ProviderInfo]:
        """Get provider information by name."""
        return self._providers.get(name)
    
    def list_providers(self) -> List[str]:
        """List all available provider names."""
        return list(self._providers.keys())
    
    def list_available_providers(self) -> List[str]:
        """List providers that are properly configured and available."""
        return [
            name for name, info in self._providers.items()
            if info.client and info.client.is_available()
        ]
    
    def get_provider_order(self, prompt: str, webby: bool = False) -> List[str]:
        """Get the preferred order of providers for a given prompt."""
        settings = get_settings()
        available = self.list_available_providers()
        
        if not available:
            return []
        
        # Check for forced provider
        forced = self._parse_force_provider(prompt)
        if forced and forced in available:
            return [forced]
        
        # Webby prompts prefer Perplexity
        if webby and "perplexity" in available:
            if settings.force_perplexity_web:
                return ["perplexity"]
            order = [
                "perplexity",
                "openai",
                "anthropic",
                "gemini",
                "mistral",
                "grok",
                "llama",
                "local",
            ]
        elif settings.prefer_perplexity and "perplexity" in available:
            order = [
                "perplexity",
                "openai",
                "anthropic",
                "gemini",
                "mistral",
                "grok",
                "llama",
                "local",
            ]
        elif any(k in prompt.lower() for k in ("ppt", "powerpoint", "slide")):
            # PowerPoint tasks prefer Anthropic
            order = [
                "anthropic",
                "openai",
                "gemini",
                "perplexity",
                "mistral",
                "grok",
                "llama",
                "local",
            ]
        elif any(k in prompt.lower() for k in ("code", "python", "error", "traceback", "stack trace")):
            # Code tasks prefer OpenAI
            order = [
                "openai",
                "mistral",
                "perplexity",
                "anthropic",
                "gemini",
                "llama",
                "local",
                "grok",
            ]
        else:
            # Default order
            order = [
                "openai",
                "anthropic",
                "perplexity",
                "gemini",
                "mistral",
                "grok",
                "llama",
                "local",
            ]
        
        # Filter to only available providers and limit to top 3
        return [name for name in order if name in available][:3]
    
    def _parse_force_provider(self, prompt: str) -> Optional[str]:
        """Parse forced provider from prompt."""
        import re
        m = re.search(
            r"\buse:\s*(openai|claude|anthropic|gemini|grok|mistral|perplexity|local|llama)\b",
            prompt,
            re.I,
        )
        if not m:
            return None
        tok = m.group(1).lower()
        return {"claude": "anthropic"}.get(tok, tok)


# Global registry instance
_registry: Optional[ProviderRegistry] = None


def get_provider_registry() -> ProviderRegistry:
    """Get the global provider registry instance."""
    global _registry
    if _registry is None:
        _registry = ProviderRegistry()
    return _registry
