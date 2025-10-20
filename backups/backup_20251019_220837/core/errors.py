# core/errors.py
"""
Structured error handling for the application.
"""
import asyncio
import logging
import time
from typing import Optional, Dict, Any, Union, Callable, TypeVar, Generic
from dataclasses import dataclass
from enum import Enum
import traceback


class ErrorType(Enum):
    """Types of errors that can occur in the application."""
    API_ERROR = "api_error"
    NETWORK_ERROR = "network_error"
    AUTHENTICATION_ERROR = "authentication_error"
    VALIDATION_ERROR = "validation_error"
    CONFIGURATION_ERROR = "configuration_error"
    RATE_LIMIT_ERROR = "rate_limit_error"
    TIMEOUT_ERROR = "timeout_error"
    UNKNOWN_ERROR = "unknown_error"


@dataclass
class APIError(Exception):
    """Structured API error with context and retry information."""
    provider: str
    error_type: ErrorType
    message: str
    status_code: Optional[int] = None
    retry_after: Optional[int] = None
    original_error: Optional[Exception] = None
    context: Optional[Dict[str, Any]] = None
    
    def __str__(self) -> str:
        return f"{self.provider} {self.error_type.value}: {self.message}"


@dataclass
class RetryConfig:
    """Configuration for retry logic."""
    max_retries: int = 3
    base_delay: float = 1.0
    max_delay: float = 60.0
    exponential_base: float = 2.0
    jitter: bool = True
    
    def get_delay(self, attempt: int) -> float:
        """Calculate delay for the given attempt number."""
        delay = self.base_delay * (self.exponential_base ** attempt)
        delay = min(delay, self.max_delay)
        
        if self.jitter:
            # Add random jitter to prevent thundering herd
            import random
            delay *= (0.5 + random.random() * 0.5)
        
        return delay


class ErrorHandler:
    """Centralized error handling and logging."""
    
    def __init__(self, logger_name: str = "jarvis"):
        self.logger = logging.getLogger(logger_name)
        self._setup_logging()
    
    def _setup_logging(self):
        """Setup logging configuration."""
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)
    
    def log_error(self, error: APIError, context: Optional[Dict[str, Any]] = None):
        """Log an error with context."""
        log_context = {
            "provider": error.provider,
            "error_type": error.error_type.value,
            "status_code": error.status_code,
            "retry_after": error.retry_after,
        }
        if context:
            log_context.update(context)
        
        self.logger.error(
            f"API Error: {error.message}",
            extra=log_context,
            exc_info=error.original_error
        )
    
    def should_retry(self, error: APIError, attempt: int, max_retries: int) -> bool:
        """Determine if an error should be retried."""
        if attempt >= max_retries:
            return False
        
        # Don't retry authentication errors
        if error.error_type == ErrorType.AUTHENTICATION_ERROR:
            return False
        
        # Don't retry validation errors
        if error.error_type == ErrorType.VALIDATION_ERROR:
            return False
        
        # Retry rate limit errors with longer delay
        if error.error_type == ErrorType.RATE_LIMIT_ERROR:
            return True
        
        # Retry network and timeout errors
        if error.error_type in [ErrorType.NETWORK_ERROR, ErrorType.TIMEOUT_ERROR]:
            return True
        
        # Retry API errors (server errors)
        if error.error_type == ErrorType.API_ERROR and error.status_code and error.status_code >= 500:
            return True
        
        return False


T = TypeVar('T')


class RetryableOperation(Generic[T]):
    """Wrapper for operations that can be retried."""
    
    def __init__(self, error_handler: ErrorHandler, retry_config: RetryConfig):
        self.error_handler = error_handler
        self.retry_config = retry_config
    
    async def execute(
        self,
        operation: Callable[[], T],
        operation_name: str = "operation",
        context: Optional[Dict[str, Any]] = None
    ) -> T:
        """Execute an operation with retry logic."""
        last_error = None
        
        for attempt in range(self.retry_config.max_retries + 1):
            try:
                if asyncio.iscoroutinefunction(operation):
                    return await operation()
                else:
                    return operation()
            except Exception as e:
                error = self._wrap_exception(e, operation_name, context)
                last_error = error
                
                self.error_handler.log_error(error, context)
                
                if not self.error_handler.should_retry(error, attempt, self.retry_config.max_retries):
                    break
                
                if attempt < self.retry_config.max_retries:
                    delay = self.retry_config.get_delay(attempt)
                    if error.retry_after:
                        delay = max(delay, error.retry_after)
                    
                    self.logger.info(f"Retrying {operation_name} in {delay:.2f}s (attempt {attempt + 1})")
                    await asyncio.sleep(delay)
        
        raise last_error
    
    def _wrap_exception(self, exc: Exception, operation_name: str, context: Optional[Dict[str, Any]]) -> APIError:
        """Wrap a generic exception into an APIError."""
        error_type = ErrorType.UNKNOWN_ERROR
        status_code = None
        retry_after = None
        
        # Determine error type based on exception
        if isinstance(exc, (ConnectionError, TimeoutError)):
            error_type = ErrorType.NETWORK_ERROR
        elif "timeout" in str(exc).lower():
            error_type = ErrorType.TIMEOUT_ERROR
        elif "rate limit" in str(exc).lower():
            error_type = ErrorType.RATE_LIMIT_ERROR
            retry_after = 60  # Default retry after 1 minute
        elif "unauthorized" in str(exc).lower() or "forbidden" in str(exc).lower():
            error_type = ErrorType.AUTHENTICATION_ERROR
        
        return APIError(
            provider=operation_name,
            error_type=error_type,
            message=str(exc),
            original_error=exc,
            context=context
        )


# Global error handler instance
error_handler = ErrorHandler()
default_retry_config = RetryConfig()


def create_retryable_operation(
    retry_config: Optional[RetryConfig] = None
) -> RetryableOperation:
    """Create a retryable operation with the given retry configuration."""
    return RetryableOperation(error_handler, retry_config or default_retry_config)


# Convenience function for simple retry operations
async def retry_operation(
    operation: Callable[[], T],
    operation_name: str = "operation",
    retry_config: Optional[RetryConfig] = None,
    context: Optional[Dict[str, Any]] = None
) -> T:
    """Execute an operation with retry logic."""
    retryable = create_retryable_operation(retry_config)
    return await retryable.execute(operation, operation_name, context)
