# routing/router.py
"""
Main router for selecting and calling LLM providers.
"""
import asyncio
from typing import List, Optional, Dict, Any
from dataclasses import dataclass

from providers.base import ChatMessage
from .provider_registry import get_provider_registry
from .web_scorer import get_web_scorer
from config.settings_simple import get_settings
from core.errors import APIError, ErrorType


@dataclass
class RouterResult:
    """Result from router execution."""
    content: str
    provider: str
    model: str
    web_score: int
    metadata: Optional[Dict[str, Any]] = None


class Router:
    """Main router for LLM provider selection and execution."""
    
    def __init__(self):
        self.provider_registry = get_provider_registry()
        self.web_scorer = get_web_scorer()
        self.settings = get_settings()
    
    async def route(
        self,
        prompt: str,
        system: str = "Be precise and helpful.",
        provider: Optional[str] = None
    ) -> RouterResult:
        """
        Route a prompt to the appropriate provider.
        
        Args:
            prompt: User prompt
            system: System message
            provider: Force specific provider (optional)
            
        Returns:
            RouterResult with content and metadata
        """
        # Convert to chat messages
        messages = [
            ChatMessage(role="system", content=system),
            ChatMessage(role="user", content=prompt),
        ]
        
        # Determine if prompt is webby
        webby = self.web_scorer.looks_webby(prompt)
        web_score = self.web_scorer.get_web_score(prompt)
        
        # Get provider order
        if provider:
            providers = [provider] if provider in self.provider_registry.list_available_providers() else []
        else:
            providers = self.provider_registry.get_provider_order(prompt, webby)
        
        if not providers:
            raise APIError(
                provider="router",
                error_type=ErrorType.CONFIGURATION_ERROR,
                message="No providers configured or available"
            )
        
        # Try providers in order
        last_error = None
        for provider_name in providers:
            try:
                result = await self._call_provider(provider_name, messages, webby)
                return RouterResult(
                    content=result.content,
                    provider=provider_name,
                    model=result.model,
                    web_score=web_score,
                    metadata=result.metadata
                )
            except Exception as e:
                last_error = e
                # Log error but continue to next provider
                import logging
                logger = logging.getLogger(__name__)
                logger.warning(f"Provider {provider_name} failed: {e}")
                continue
        
        # If all providers failed, raise the last error
        if last_error:
            raise last_error
        
        raise APIError(
            provider="router",
            error_type=ErrorType.UNKNOWN_ERROR,
            message="All providers failed"
        )
    
    async def _call_provider(
        self,
        provider_name: str,
        messages: List[ChatMessage],
        webby: bool
    ):
        """Call a specific provider."""
        provider = self.provider_registry.get_provider(provider_name)
        if not provider:
            raise APIError(
                provider=provider_name,
                error_type=ErrorType.CONFIGURATION_ERROR,
                message=f"Provider {provider_name} not available"
            )
        
        # Special handling for Perplexity webby prompts
        if webby and provider_name == "perplexity":
            # Use web-optimized parameters
            result = await provider.chat(
                messages=messages,
                temperature=0.0,
                max_tokens=min(1800, self.settings.max_tokens)
            )
            
            # Check if we need synthesis
            if self._needs_synthesis(result.content):
                return await self._synthesize_result(result, messages, webby)
            
            return result
        
        # Standard call
        return await provider.chat(messages=messages)
    
    def _needs_synthesis(self, content: str) -> bool:
        """Check if content needs synthesis."""
        if not content:
            return False
        
        content_lower = content.lower()
        stale_indicators = [
            "as of ", "cannot browse", "no web access", "2023"
        ]
        
        return (
            self.settings.always_synthesize_web or
            any(indicator in content_lower for indicator in stale_indicators)
        )
    
    async def _synthesize_result(self, research_result, original_messages, webby):
        """Synthesize research result with GPT if available."""
        if "openai" not in self.provider_registry.list_available_providers():
            return research_result
        
        openai_provider = self.provider_registry.get_provider("openai")
        if not openai_provider:
            return research_result
        
        # Create synthesis prompt
        synthesis_messages = [
            ChatMessage(role="system", content=original_messages[0].content),
            ChatMessage(
                role="user",
                content=(
                    "Combine the live research below into a current answer.\n"
                    "• Start with a one-line answer using concrete numbers/dates.\n"
                    "• Follow with 4–7 crisp bullets.\n"
                    "• Keep inline citations (site or short URL) on specific claims.\n"
                    "• End with a one-line takeaway.\n\n"
                    f"Question: {original_messages[1].content}\n\nResearch:\n{research_result.content}"
                )
            )
        ]
        
        return await openai_provider.chat(messages=synthesis_messages)
    
    async def route_with_fallback(
        self,
        prompt: str,
        system: str = "Be precise and helpful.",
        provider: Optional[str] = None,
        timeout: float = 60.0
    ) -> RouterResult:
        """
        Route with timeout and fallback handling.
        
        Args:
            prompt: User prompt
            system: System message
            provider: Force specific provider (optional)
            timeout: Timeout in seconds
            
        Returns:
            RouterResult with content and metadata
        """
        try:
            return await asyncio.wait_for(
                self.route(prompt, system, provider),
                timeout=timeout
            )
        except asyncio.TimeoutError:
            raise APIError(
                provider="router",
                error_type=ErrorType.TIMEOUT_ERROR,
                message=f"Request timed out after {timeout} seconds"
            )


# Global router instance
_router: Optional[Router] = None


def get_router() -> Router:
    """Get the global router instance."""
    global _router
    if _router is None:
        _router = Router()
    return _router


# Convenience function for backward compatibility
async def superchat(prompt: str, system: str = "Be precise and helpful.") -> str:
    """Backward compatible function for the original superchat interface."""
    router = get_router()
    result = await router.route(prompt, system)
    return result.content
