#!/usr/bin/env python3
"""
Migration script to transition from agent.py to agent_v2.py

This script helps migrate existing code to use the new modular architecture.
"""

import os
import shutil
from pathlib import Path
from datetime import datetime


def backup_original():
    """Create a backup of the original agent.py"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"agent_backup_{timestamp}.py"
    
    if os.path.exists("agent.py"):
        shutil.copy2("agent.py", backup_path)
        print(f"✅ Backed up original agent.py to {backup_path}")
        return backup_path
    else:
        print("⚠️  No agent.py found to backup")
        return None


def create_migration_guide():
    """Create a migration guide for users"""
    guide_content = """# Migration Guide: agent.py → agent_v2.py

## Overview
The new agent_v2.py provides the same functionality as the original agent.py but with:
- Better error handling
- Modular architecture
- Centralized configuration
- Improved logging
- Comprehensive testing

## Key Changes

### 1. Configuration
- All configuration is now centralized in `config/settings.py`
- Environment variables are validated and typed
- No more scattered configuration throughout the code

### 2. Provider Management
- Providers are now modular classes in `providers/`
- Each provider has its own implementation
- Better error handling and retry logic

### 3. Error Handling
- Structured error types in `core/errors.py`
- Proper logging and error context
- Retry logic with exponential backoff

### 4. Routing
- Smart provider selection in `routing/`
- Web content scoring for better provider choice
- Fallback mechanisms

## Migration Steps

### Step 1: Update Imports
Replace:
```python
from agent import superchat, openai_chat, perplexity_chat
```

With:
```python
from agent_v2 import superchat, openai_chat, perplexity_chat
```

### Step 2: Update Configuration
Move your environment variables to `.env` file if not already there:
```bash
OPENAI_API_KEY=your_key
PPLX_API_KEY=your_key
# ... etc
```

### Step 3: Update Error Handling
The new system provides better error information:
```python
try:
    result = await superchat("Hello")
except APIError as e:
    print(f"Provider {e.provider} failed: {e.message}")
    print(f"Error type: {e.error_type}")
```

### Step 4: Test Your Integration
Run the test suite to ensure everything works:
```bash
python -m pytest tests/test_agent_v2.py -v
```

## Backward Compatibility
All original functions are still available:
- `superchat(prompt, system)`
- `openai_chat(messages, model, temperature, max_tokens)`
- `perplexity_chat(messages, model, temperature, max_tokens)`
- `anthropic_chat(messages, model, temperature, max_tokens)`
- `gemini_chat(messages, model, temperature, max_tokens)`
- `grok_chat(messages, model, temperature, max_tokens)`
- `mistral_chat(messages, model, temperature, max_tokens)`

## New Features

### 1. Better Provider Selection
```python
from agent_v2 import chat_with_provider

# Auto-select best provider
result = await chat_with_provider("What's the latest news?")

# Force specific provider
result = await chat_with_provider("Hello", provider="openai")
```

### 2. Structured Results
```python
from routing.router import get_router

router = get_router()
result = await router.route("Hello")

print(f"Content: {result.content}")
print(f"Provider: {result.provider}")
print(f"Model: {result.model}")
print(f"Web Score: {result.web_score}")
```

### 3. Configuration Management
```python
from config.settings import get_settings

settings = get_settings()
print(f"Max tokens: {settings.max_tokens}")
print(f"Available providers: {settings.openai_api_key is not None}")
```

## Troubleshooting

### Common Issues

1. **Import Errors**: Make sure all new modules are in your Python path
2. **Configuration Errors**: Check that your `.env` file is properly formatted
3. **Provider Errors**: Verify API keys are correct and providers are available

### Getting Help
- Check the test suite for usage examples
- Review the error logs for detailed error information
- Use the structured error types to handle specific error cases

## Performance Improvements
- Connection pooling for HTTP requests
- Better retry logic with exponential backoff
- Reduced code duplication
- More efficient provider selection

## Security Improvements
- Input validation and sanitization
- Better API key management
- Structured error handling prevents information leakage
"""
    
    with open("MIGRATION_GUIDE.md", "w") as f:
        f.write(guide_content)
    
    print("✅ Created MIGRATION_GUIDE.md")


def create_requirements_v2():
    """Create updated requirements.txt with new dependencies"""
    requirements_v2 = """# Updated requirements for agent_v2.py
# Core dependencies
pydantic>=2.0.0
httpx>=0.27.0
asyncio-mqtt>=0.16.0

# Existing dependencies (keep these)
fastapi>=0.111,<1
uvicorn[standard]>=0.30,<1
requests>=2.32,<3
python-multipart>=0.0.9
python-dotenv>=1.0

# LLM/Search/Parsing
openai>=1.51
beautifulsoup4>=4.12
duckduckgo-search>=6.3
lxml>=5.3
feedparser>=6.0

# Data / Office docs
pandas>=2.2
openpyxl>=3.1
python-pptx>=0.6.23
Pillow>=10.4
reportlab>=4.2

# SEC / Finance
sec-edgar-downloader>=5.0

# Google / Microsoft / Cloud connectors
msal>=1.28
google-api-python-client>=2.142
google-auth>=2.34
google-auth-oauthlib>=1.2

# DevOps / self-improver / tests / lint
gitpython>=3.1
pyyaml>=6.0
pytest>=8.3
pytest-asyncio>=0.21.0
ruff>=0.6
black>=24.8
mypy>=1.11
bandit>=1.7

# Windows-only (safe to ignore on macOS/Linux)
pywin32; platform_system == "Windows"
"""
    
    with open("requirements_v2.txt", "w") as f:
        f.write(requirements_v2)
    
    print("✅ Created requirements_v2.txt")


def main():
    """Main migration function"""
    print("🚀 Starting migration to agent_v2.py...")
    
    # Create backup
    backup_path = backup_original()
    
    # Create migration guide
    create_migration_guide()
    
    # Create updated requirements
    create_requirements_v2()
    
    print("\n📋 Migration Summary:")
    print("1. ✅ Backed up original agent.py")
    print("2. ✅ Created MIGRATION_GUIDE.md")
    print("3. ✅ Created requirements_v2.txt")
    
    print("\n🔧 Next Steps:")
    print("1. Review MIGRATION_GUIDE.md")
    print("2. Update your imports to use agent_v2")
    print("3. Install new requirements: pip install -r requirements_v2.txt")
    print("4. Run tests: python -m pytest tests/test_agent_v2.py -v")
    print("5. Test your integration with the new system")
    
    if backup_path:
        print(f"\n💾 Original agent.py backed up to: {backup_path}")
    
    print("\n✨ Migration complete! Check MIGRATION_GUIDE.md for detailed instructions.")


if __name__ == "__main__":
    main()
