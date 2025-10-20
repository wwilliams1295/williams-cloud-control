# providers/anthropic.py
"""
Anthropic Claude API client implementation.
"""
from typing import Dict, Any, List, Optional
from .base import BaseAPIClient, APIResponse, ChatMessage


class AnthropicClient(BaseAPIClient):
    """Anthropic Claude API client."""
    
    def __init__(
        self,
        api_key: str,
        model: str = "claude-sonnet-4-5-20250929",
        timeout: float = 60.0,
        max_tokens: int = 2000,
        temperature: float = 0.2
    ):
        super().__init__(
            api_key=api_key,
            base_url="https://api.anthropic.com",
            model=model,
            timeout=timeout,
            max_tokens=max_tokens,
            temperature=temperature
        )
    
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for Anthropic API requests."""
        return {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        }
    
    def _build_payload(
        self,
        messages: List[ChatMessage],
        **kwargs
    ) -> Dict[str, Any]:
        """Build the request payload for Anthropic API."""
        # Separate system message from other messages
        system_messages = [msg for msg in messages if msg.role == "system"]
        other_messages = [msg for msg in messages if msg.role != "system"]
        
        system_text = (
            "\n".join(msg.content for msg in system_messages)
            if system_messages else "You are helpful."
        )
        
        # Convert to Anthropic format
        turns = [
            {"role": msg.role, "content": [{"type": "text", "text": msg.content}]}
            for msg in other_messages
            if msg.role in ("user", "assistant")
        ]
        
        payload = {
            "model": kwargs.get("model", self.model),
            "system": system_text,
            "messages": turns,
            "temperature": kwargs.get("temperature", self.temperature),
            "max_tokens": kwargs.get("max_tokens", self.max_tokens),
        }
        
        # Add any additional parameters
        for key, value in kwargs.items():
            if key not in ["model", "system", "messages", "temperature", "max_tokens"]:
                payload[key] = value
        
        return payload
    
    def _parse_response(self, response: Dict[str, Any]) -> APIResponse:
        """Parse Anthropic API response."""
        try:
            content_parts = response.get("content", [])
            content = "".join(
                part.get("text", "")
                for part in content_parts
                if part.get("type") == "text"
            ) or str(response)
            
            model = response.get("model", self.model)
            usage = response.get("usage")
            
            return APIResponse(
                content=content,
                model=model,
                usage=usage,
                metadata={
                    "stop_reason": response.get("stop_reason"),
                    "response_id": response.get("id"),
                }
            )
        except (KeyError, IndexError) as e:
            raise ValueError(f"Invalid Anthropic response format: {e}")
    
    def _get_endpoint(self) -> str:
        """Get the Anthropic messages endpoint."""
        return "/v1/messages"
