# providers/gemini.py
"""
Google Gemini API client implementation.
"""
from typing import Dict, Any, List, Optional
from .base import BaseAPIClient, APIResponse, ChatMessage


class GeminiClient(BaseAPIClient):
    """Google Gemini API client."""
    
    def __init__(
        self,
        api_key: str,
        model: str = "gemini-2.5-flash",
        timeout: float = 60.0,
        max_tokens: int = 2000,
        temperature: float = 0.0
    ):
        super().__init__(
            api_key=api_key,
            base_url="https://generativelanguage.googleapis.com",
            model=model,
            timeout=timeout,
            max_tokens=max_tokens,
            temperature=temperature
        )
    
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for Gemini API requests."""
        return {
            "Content-Type": "application/json"
        }
    
    def _build_payload(
        self,
        messages: List[ChatMessage],
        **kwargs
    ) -> Dict[str, Any]:
        """Build the request payload for Gemini API."""
        # Convert OpenAI-style messages to Gemini contents
        contents = []
        for msg in messages:
            role = (
                "user" if msg.role == "user"
                else ("model" if msg.role == "assistant" else "user")
            )
            contents.append({"role": role, "parts": [{"text": msg.content}]})
        
        payload = {
            "contents": contents,
            "generationConfig": {
                "temperature": kwargs.get("temperature", self.temperature),
                "maxOutputTokens": kwargs.get("max_tokens", self.max_tokens),
            }
        }
        
        # Add any additional parameters
        for key, value in kwargs.items():
            if key not in ["contents", "generationConfig", "temperature", "max_tokens"]:
                payload[key] = value
        
        return payload
    
    def _parse_response(self, response: Dict[str, Any]) -> APIResponse:
        """Parse Gemini API response."""
        try:
            content = response["candidates"][0]["content"]["parts"][0]["text"]
            model = self.model  # Gemini doesn't return model in response
            
            return APIResponse(
                content=content,
                model=model,
                usage=None,  # Gemini doesn't provide usage info in this format
                metadata={
                    "finish_reason": response["candidates"][0].get("finishReason"),
                    "candidate_count": len(response.get("candidates", [])),
                }
            )
        except (KeyError, IndexError) as e:
            raise ValueError(f"Invalid Gemini response format: {e}")
    
    def _get_endpoint(self) -> str:
        """Get the Gemini generate content endpoint."""
        return f"/v1beta/models/{self.model}:generateContent"
    
    async def _make_request(
        self,
        endpoint: str,
        payload: Dict[str, Any],
        context: Optional[Dict[str, Any]] = None
    ) -> APIResponse:
        """Override to add API key to URL for Gemini."""
        # Gemini uses API key in URL, not headers
        url = f"{self.base_url}{endpoint}?key={self.api_key}"
        
        async def _request():
            import httpx
            async with httpx.AsyncClient(timeout=self.timeout) as client:
                response = await client.post(
                    url,
                    headers=self._get_headers(),
                    json=payload
                )
                response.raise_for_status()
                return self._parse_response(response.json())
        
        from core.errors import retry_operation
        return await retry_operation(
            _request,
            operation_name=f"{self.__class__.__name__}_request",
            context=context
        )
