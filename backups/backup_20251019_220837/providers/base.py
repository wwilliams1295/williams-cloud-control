# providers/base.py
"""
Base API client with common functionality for all LLM providers.
"""
import asyncio
import httpx
from typing import Dict, Any, Optional, List, Union
from dataclasses import dataclass
from abc import ABC, abstractmethod
import logging

from core.errors import APIError, ErrorType, retry_operation


@dataclass
class APIResponse:
    """Standardized response from API calls."""
    content: str
    model: str
    usage: Optional[Dict[str, Any]] = None
    metadata: Optional[Dict[str, Any]] = None


@dataclass
class ChatMessage:
    """A single message in a chat conversation."""
    role: str  # "system", "user", "assistant"
    content: str


class BaseAPIClient(ABC):
    """Base class for all LLM API clients."""
    
    def __init__(
        self,
        api_key: str,
        base_url: str,
        model: str,
        timeout: float = 60.0,
        max_tokens: int = 2000,
        temperature: float = 0.7
    ):
        self.api_key = api_key
        self.base_url = base_url.rstrip('/')
        self.model = model
        self.timeout = timeout
        self.max_tokens = max_tokens
        self.temperature = temperature
        self.logger = logging.getLogger(f"{self.__class__.__name__}")
    
    @abstractmethod
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for API requests."""
        pass
    
    @abstractmethod
    def _build_payload(
        self,
        messages: List[ChatMessage],
        **kwargs
    ) -> Dict[str, Any]:
        """Build the request payload for the API."""
        pass
    
    @abstractmethod
    def _parse_response(self, response: Dict[str, Any]) -> APIResponse:
        """Parse the API response into a standardized format."""
        pass
    
    async def _make_request(
        self,
        endpoint: str,
        payload: Dict[str, Any],
        context: Optional[Dict[str, Any]] = None
    ) -> APIResponse:
        """Make an HTTP request with retry logic."""
        async def _request():
            import ssl
            
            # Create SSL context that works with older LibreSSL
            ssl_context = ssl.create_default_context()
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE
            
            # Configure httpx with SSL context and longer timeout
            timeout = httpx.Timeout(30.0, connect=10.0)
            
            async with httpx.AsyncClient(
                timeout=timeout,
                verify=ssl_context,
                limits=httpx.Limits(max_keepalive_connections=5, max_connections=10)
            ) as client:
                response = await client.post(
                    f"{self.base_url}{endpoint}",
                    headers=self._get_headers(),
                    json=payload
                )
                response.raise_for_status()
                return self._parse_response(response.json())
        
        return await retry_operation(
            _request,
            operation_name=f"{self.__class__.__name__}_request",
            context=context
        )
    
    async def chat(
        self,
        messages: List[ChatMessage],
        model: Optional[str] = None,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        **kwargs
    ) -> APIResponse:
        """
        Send a chat request to the API.
        
        Args:
            messages: List of chat messages
            model: Override the default model
            temperature: Override the default temperature
            max_tokens: Override the default max_tokens
            **kwargs: Additional parameters for the specific API
            
        Returns:
            APIResponse with the generated content
        """
        # Use provided values or fall back to defaults
        model = model or self.model
        temperature = temperature if temperature is not None else self.temperature
        max_tokens = max_tokens or self.max_tokens
        
        # Build the payload
        payload = self._build_payload(
            messages,
            model=model,
            temperature=temperature,
            max_tokens=max_tokens,
            **kwargs
        )
        
        # Make the request
        context = {
            "model": model,
            "temperature": temperature,
            "max_tokens": max_tokens,
            "message_count": len(messages)
        }
        
        return await self._make_request(
            self._get_endpoint(),
            payload,
            context=context
        )
    
    @abstractmethod
    def _get_endpoint(self) -> str:
        """Get the API endpoint for chat requests."""
        pass
    
    def is_available(self) -> bool:
        """Check if the API client is properly configured."""
        return bool(self.api_key and self.base_url)
    
    def __str__(self) -> str:
        return f"{self.__class__.__name__}(model={self.model}, available={self.is_available()})"
