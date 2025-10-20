# providers/mistral.py
"""
Mistral API client implementation.
"""
from typing import Dict, Any, List, Optional
from .base import BaseAPIClient, APIResponse, ChatMessage


class MistralClient(BaseAPIClient):
    """Mistral API client."""
    
    def __init__(
        self,
        api_key: str,
        model: str = "mistral-large-latest",
        timeout: float = 60.0,
        max_tokens: int = 2000,
        temperature: float = 0.4
    ):
        super().__init__(
            api_key=api_key,
            base_url="https://api.mistral.ai",
            model=model,
            timeout=timeout,
            max_tokens=max_tokens,
            temperature=temperature
        )
    
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for Mistral API requests."""
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
    
    def _build_payload(
        self,
        messages: List[ChatMessage],
        **kwargs
    ) -> Dict[str, Any]:
        """Build the request payload for Mistral API."""
        # Convert ChatMessage objects to Mistral format
        mistral_messages = [
            {"role": msg.role, "content": msg.content}
            for msg in messages
        ]
        
        payload = {
            "model": kwargs.get("model", self.model),
            "messages": mistral_messages,
            "temperature": kwargs.get("temperature", self.temperature),
            "max_tokens": kwargs.get("max_tokens", self.max_tokens),
        }
        
        # Add any additional parameters
        for key, value in kwargs.items():
            if key not in ["model", "messages", "temperature", "max_tokens"]:
                payload[key] = value
        
        return payload
    
    def _parse_response(self, response: Dict[str, Any]) -> APIResponse:
        """Parse Mistral API response."""
        try:
            content = response["choices"][0]["message"]["content"]
            usage = response.get("usage")
            model = response.get("model", self.model)
            
            return APIResponse(
                content=content,
                model=model,
                usage=usage,
                metadata={
                    "finish_reason": response["choices"][0].get("finish_reason"),
                    "response_id": response.get("id"),
                }
            )
        except (KeyError, IndexError) as e:
            raise ValueError(f"Invalid Mistral response format: {e}")
    
    def _get_endpoint(self) -> str:
        """Get the Mistral chat endpoint."""
        return "/v1/chat/completions"
