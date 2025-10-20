# providers/perplexity.py
"""
Perplexity API client implementation.
"""
import re
from typing import Dict, Any, List, Optional
from .base import BaseAPIClient, APIResponse, ChatMessage


class PerplexityClient(BaseAPIClient):
    """Perplexity API client with reference cleaning."""
    
    def __init__(
        self,
        api_key: str,
        model: str = "sonar",
        timeout: float = 60.0,
        max_tokens: int = 1800,
        temperature: float = 0.0
    ):
        super().__init__(
            api_key=api_key,
            base_url="https://api.perplexity.ai",
            model=model,
            timeout=timeout,
            max_tokens=max_tokens,
            temperature=temperature
        )
    
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for Perplexity API requests."""
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
    
    def _build_payload(
        self,
        messages: List[ChatMessage],
        **kwargs
    ) -> Dict[str, Any]:
        """Build the request payload for Perplexity API."""
        # Convert ChatMessage objects to Perplexity format
        pplx_messages = [
            {"role": msg.role, "content": msg.content}
            for msg in messages
        ]
        
        # Light token clamp for speed if it's clearly webby by phrasing
        max_tokens = kwargs.get("max_tokens", self.max_tokens)
        try:
            user_text = " ".join(
                msg.content for msg in messages if msg.role == "user"
            ).lower()
            if any(
                k in user_text
                for k in (
                    "latest", "current", "today", "now", "trend", "news",
                    "rate", "price", "yield"
                )
            ):
                max_tokens = min(max_tokens, 400)
        except Exception:
            pass
        
        payload = {
            "model": kwargs.get("model", self.model),
            "messages": pplx_messages,
            "temperature": kwargs.get("temperature", self.temperature),
            "max_tokens": max_tokens,
        }
        
        # Add any additional parameters
        for key, value in kwargs.items():
            if key not in ["model", "messages", "temperature", "max_tokens"]:
                payload[key] = value
        
        return payload
    
    def _parse_response(self, response: Dict[str, Any]) -> APIResponse:
        """Parse Perplexity API response and clean references."""
        try:
            content = response["choices"][0]["message"]["content"]
            # Clean Perplexity-style references
            cleaned_content = self._clean_perplexity_refs(content)
            
            usage = response.get("usage")
            model = response.get("model", self.model)
            
            return APIResponse(
                content=cleaned_content,
                model=model,
                usage=usage,
                metadata={
                    "finish_reason": response["choices"][0].get("finish_reason"),
                    "response_id": response.get("id"),
                }
            )
        except (KeyError, IndexError) as e:
            raise ValueError(f"Invalid Perplexity response format: {e}")
    
    def _clean_perplexity_refs(self, text: str) -> str:
        """Remove Perplexity-style bracketed refs like [1], [12], (1), [a]."""
        if not text:
            return text
        
        # Remove various reference patterns
        text = re.sub(r"\[\s*\d+\s*\]", "", text)  # [1], [12]
        text = re.sub(r"\(\s*\d+\s*\)", "", text)  # (1)
        text = re.sub(r"\[\s*[a-zA-Z]\s*\]", "", text)  # [a]
        text = re.sub(r"\s{2,}", " ", text)  # extra spaces
        text = re.sub(r"\s+([,.;:!?])", r"\1", text)  # trim before punctuation
        
        return text.strip()
    
    def _get_endpoint(self) -> str:
        """Get the Perplexity chat endpoint."""
        return "/chat/completions"
