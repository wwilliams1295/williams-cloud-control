# providers/openai_compatible.py
"""
OpenAI-compatible API client for local and other compatible endpoints.
"""
from typing import Dict, Any, List, Optional
from .base import BaseAPIClient, APIResponse, ChatMessage


class OpenAICompatibleClient(BaseAPIClient):
    """OpenAI-compatible API client for local and other endpoints."""
    
    def __init__(
        self,
        api_key: Optional[str],
        base_url: str,
        model: str,
        timeout: float = 60.0,
        max_tokens: int = 2000,
        temperature: float = 0.4
    ):
        super().__init__(
            api_key=api_key or "",
            base_url=base_url,
            model=model,
            timeout=timeout,
            max_tokens=max_tokens,
            temperature=temperature
        )
    
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for OpenAI-compatible API requests."""
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        return headers
    
    def _build_payload(
        self,
        messages: List[ChatMessage],
        **kwargs
    ) -> Dict[str, Any]:
        """Build the request payload for OpenAI-compatible API."""
        # Convert ChatMessage objects to OpenAI format
        openai_messages = [
            {"role": msg.role, "content": msg.content}
            for msg in messages
        ]
        
        payload = {
            "model": kwargs.get("model", self.model),
            "messages": openai_messages,
            "temperature": kwargs.get("temperature", self.temperature),
            "max_tokens": kwargs.get("max_tokens", self.max_tokens),
        }
        
        # Add any additional parameters
        for key, value in kwargs.items():
            if key not in ["model", "messages", "temperature", "max_tokens"]:
                payload[key] = value
        
        return payload
    
    def _parse_response(self, response: Dict[str, Any]) -> APIResponse:
        """Parse OpenAI-compatible API response."""
        try:
            # Try OpenAI-style response first
            if "choices" in response and response["choices"]:
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
            # Try text-style response
            elif "text" in response:
                return APIResponse(
                    content=response["text"],
                    model=self.model,
                    usage=None,
                    metadata={"response_type": "text"}
                )
            else:
                # Fallback to string representation
                return APIResponse(
                    content=str(response),
                    model=self.model,
                    usage=None,
                    metadata={"response_type": "fallback"}
                )
        except (KeyError, IndexError) as e:
            raise ValueError(f"Invalid OpenAI-compatible response format: {e}")
    
    def _get_endpoint(self) -> str:
        """Get the OpenAI-compatible chat endpoint."""
        return "/v1/chat/completions"
