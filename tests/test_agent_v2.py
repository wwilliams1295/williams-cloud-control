# tests/test_agent_v2.py
"""
Comprehensive test suite for the refactored agent system.
"""
import pytest
import asyncio
from unittest.mock import patch, AsyncMock, MagicMock
from typing import Dict, Any

from config.settings import Settings
from core.errors import APIError, ErrorType
from providers.base import APIResponse, ChatMessage
from routing.router import Router, RouterResult
from routing.provider_registry import ProviderRegistry
from routing.web_scorer import WebScorer


@pytest.fixture
def mock_settings():
    """Mock settings for testing."""
    settings = Settings(
        openai_api_key="test-openai-key",
        pplx_api_key="test-pplx-key",
        anthropic_api_key="test-anthropic-key",
        google_api_key="test-google-key",
        xai_api_key="test-xai-key",
        mistral_api_key="test-mistral-key",
        web_scoring_threshold=3,
        max_tokens=2000,
        http_timeout=60.0
    )
    return settings


@pytest.fixture
def mock_api_response():
    """Mock API response."""
    return APIResponse(
        content="Test response",
        model="gpt-4o-mini",
        usage={"prompt_tokens": 10, "completion_tokens": 20},
        metadata={"finish_reason": "stop"}
    )


@pytest.fixture
def mock_chat_messages():
    """Mock chat messages."""
    return [
        ChatMessage(role="system", content="You are helpful."),
        ChatMessage(role="user", content="Test prompt")
    ]


class TestWebScorer:
    """Test web scoring functionality."""
    
    def test_web_scorer_initialization(self):
        """Test web scorer initializes correctly."""
        scorer = WebScorer()
        assert scorer is not None
        assert hasattr(scorer, 'web_terms')
        assert hasattr(scorer, 'date_patterns')
    
    def test_score_webbiness(self):
        """Test web scoring logic."""
        scorer = WebScorer()
        
        # High web score
        webby_text = "What's the latest news about Apple stock today?"
        score = scorer.score_webbiness(webby_text)
        assert score > 3
        
        # Low web score
        general_text = "Explain how photosynthesis works"
        score = scorer.score_webbiness(general_text)
        assert score < 3
    
    def test_looks_webby(self):
        """Test webby detection."""
        scorer = WebScorer()
        
        assert scorer.looks_webby("What's the latest news about Apple stock today?")
        assert not scorer.looks_webby("Explain how photosynthesis works")
    
    def test_date_patterns(self):
        """Test date pattern recognition."""
        scorer = WebScorer()
        
        test_cases = [
            "Jan 15, 2024",
            "2024-01-15", 
            "Q1 2024",
            "YTD performance"
        ]
        
        for case in test_cases:
            score = scorer.score_webbiness(case)
            assert score > 0


class TestProviderRegistry:
    """Test provider registry functionality."""
    
    @patch('config.settings.get_settings')
    def test_provider_registry_initialization(self, mock_get_settings, mock_settings):
        """Test provider registry initializes correctly."""
        mock_get_settings.return_value = mock_settings
        
        registry = ProviderRegistry()
        assert registry is not None
        assert hasattr(registry, '_providers')
    
    def test_register_provider(self):
        """Test provider registration."""
        registry = ProviderRegistry()
        
        # Mock provider class
        class MockProvider:
            def __init__(self, **kwargs):
                self.api_key = kwargs.get('api_key')
            
            def is_available(self):
                return bool(self.api_key)
        
        registry.register(
            "test_provider",
            MockProvider,
            "test-model",
            "test",
            api_key="test-key"
        )
        
        assert "test_provider" in registry.list_providers()
        assert registry.get_provider("test_provider") is not None
    
    def test_get_provider_order_webby(self):
        """Test provider ordering for webby prompts."""
        registry = ProviderRegistry()
        
        # Mock available providers
        registry._providers = {
            "perplexity": MagicMock(),
            "openai": MagicMock(),
            "anthropic": MagicMock()
        }
        
        order = registry.get_provider_order("What's the latest news?", webby=True)
        assert "perplexity" in order
        assert order[0] == "perplexity"
    
    def test_parse_force_provider(self):
        """Test forced provider parsing."""
        registry = ProviderRegistry()
        
        assert registry._parse_force_provider("use: openai") == "openai"
        assert registry._parse_force_provider("use: claude") == "anthropic"
        assert registry._parse_force_provider("no provider here") is None


class TestRouter:
    """Test router functionality."""
    
    @patch('routing.router.get_provider_registry')
    @patch('routing.router.get_web_scorer')
    def test_router_initialization(self, mock_get_web_scorer, mock_get_provider_registry):
        """Test router initializes correctly."""
        mock_get_web_scorer.return_value = MagicMock()
        mock_get_provider_registry.return_value = MagicMock()
        
        router = Router()
        assert router is not None
        assert hasattr(router, 'provider_registry')
        assert hasattr(router, 'web_scorer')
    
    @patch('routing.router.get_provider_registry')
    @patch('routing.router.get_web_scorer')
    async def test_route_success(self, mock_get_web_scorer, mock_get_provider_registry, mock_api_response):
        """Test successful routing."""
        # Setup mocks
        mock_web_scorer = MagicMock()
        mock_web_scorer.looks_webby.return_value = False
        mock_web_scorer.get_web_score.return_value = 1
        mock_get_web_scorer.return_value = mock_web_scorer
        
        mock_registry = MagicMock()
        mock_registry.list_available_providers.return_value = ["openai"]
        mock_registry.get_provider_order.return_value = ["openai"]
        mock_provider = AsyncMock()
        mock_provider.chat.return_value = mock_api_response
        mock_registry.get_provider.return_value = mock_provider
        mock_get_provider_registry.return_value = mock_registry
        
        router = Router()
        result = await router.route("Test prompt")
        
        assert isinstance(result, RouterResult)
        assert result.content == "Test response"
        assert result.provider == "openai"
    
    @patch('routing.router.get_provider_registry')
    @patch('routing.router.get_web_scorer')
    async def test_route_no_providers(self, mock_get_web_scorer, mock_get_provider_registry):
        """Test routing with no available providers."""
        # Setup mocks
        mock_web_scorer = MagicMock()
        mock_get_web_scorer.return_value = mock_web_scorer
        
        mock_registry = MagicMock()
        mock_registry.list_available_providers.return_value = []
        mock_registry.get_provider_order.return_value = []
        mock_get_provider_registry.return_value = mock_registry
        
        router = Router()
        
        with pytest.raises(APIError) as exc_info:
            await router.route("Test prompt")
        
        assert exc_info.value.error_type == ErrorType.CONFIGURATION_ERROR
    
    @patch('routing.router.get_provider_registry')
    @patch('routing.router.get_web_scorer')
    async def test_route_with_fallback(self, mock_get_web_scorer, mock_get_provider_registry, mock_api_response):
        """Test routing with timeout fallback."""
        # Setup mocks
        mock_web_scorer = MagicMock()
        mock_web_scorer.looks_webby.return_value = False
        mock_web_scorer.get_web_score.return_value = 1
        mock_get_web_scorer.return_value = mock_web_scorer
        
        mock_registry = MagicMock()
        mock_registry.list_available_providers.return_value = ["openai"]
        mock_registry.get_provider_order.return_value = ["openai"]
        mock_provider = AsyncMock()
        mock_provider.chat.return_value = mock_api_response
        mock_registry.get_provider.return_value = mock_provider
        mock_get_provider_registry.return_value = mock_registry
        
        router = Router()
        result = await router.route_with_fallback("Test prompt", timeout=1.0)
        
        assert isinstance(result, RouterResult)
        assert result.content == "Test response"


class TestAgentV2:
    """Test the main agent_v2 module."""
    
    @patch('agent_v2.get_router')
    async def test_chat_with_provider_success(self, mock_get_router):
        """Test successful chat with provider."""
        mock_router = AsyncMock()
        mock_router.route.return_value = RouterResult(
            content="Test response",
            provider="openai",
            model="gpt-4o-mini",
            web_score=1
        )
        mock_get_router.return_value = mock_router
        
        from agent_v2 import chat_with_provider
        
        result = await chat_with_provider("Test prompt")
        assert result == "Test response"
    
    @patch('agent_v2.get_router')
    async def test_chat_with_provider_error(self, mock_get_router):
        """Test chat with provider error handling."""
        mock_router = AsyncMock()
        mock_router.route.side_effect = APIError(
            provider="test",
            error_type=ErrorType.API_ERROR,
            message="Test error"
        )
        mock_get_router.return_value = mock_router
        
        from agent_v2 import chat_with_provider
        
        with pytest.raises(APIError):
            await chat_with_provider("Test prompt")
    
    @patch('agent_v2.get_settings')
    @patch('agent_v2.OpenAIClient')
    async def test_openai_chat_backward_compatibility(self, mock_openai_client, mock_get_settings, mock_settings):
        """Test OpenAI chat backward compatibility."""
        mock_get_settings.return_value = mock_settings
        
        mock_client = AsyncMock()
        mock_client.chat.return_value = APIResponse(
            content="Test response",
            model="gpt-4o-mini",
            usage=None
        )
        mock_openai_client.return_value = mock_client
        
        from agent_v2 import openai_chat
        
        messages = [
            {"role": "user", "content": "Test prompt"}
        ]
        
        result = await openai_chat(messages)
        assert result == "Test response"
    
    @patch('agent_v2.get_settings')
    async def test_openai_chat_no_key(self, mock_get_settings, mock_settings):
        """Test OpenAI chat with no API key."""
        mock_settings.openai_api_key = None
        mock_get_settings.return_value = mock_settings
        
        from agent_v2 import openai_chat
        
        messages = [{"role": "user", "content": "Test prompt"}]
        result = await openai_chat(messages)
        assert result == "(OpenAI key missing.)"


class TestErrorHandling:
    """Test error handling functionality."""
    
    def test_api_error_creation(self):
        """Test API error creation."""
        error = APIError(
            provider="test",
            error_type=ErrorType.API_ERROR,
            message="Test error",
            status_code=500
        )
        
        assert error.provider == "test"
        assert error.error_type == ErrorType.API_ERROR
        assert error.message == "Test error"
        assert error.status_code == 500
        assert str(error) == "test api_error: Test error"
    
    def test_error_types(self):
        """Test error type enum."""
        assert ErrorType.API_ERROR.value == "api_error"
        assert ErrorType.NETWORK_ERROR.value == "network_error"
        assert ErrorType.AUTHENTICATION_ERROR.value == "authentication_error"


if __name__ == "__main__":
    pytest.main([__file__])
