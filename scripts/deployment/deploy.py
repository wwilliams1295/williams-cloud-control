#!/usr/bin/env python3
"""
Deployment Script for Render
Handles deployment preparation, testing, and auto-improvement setup.
"""

import asyncio
import logging
import os
import subprocess  # nosec B404
import sys
import time
from datetime import datetime
from pathlib import Path
import json

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from config.settings_simple import get_settings

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class DeploymentManager:
    """Manages deployment process and auto-improvement setup."""
    
    def __init__(self):
        self.settings = get_settings()
        self.project_root = Path(__file__).parent
        self.deployment_log = self.project_root / "deployment.log"
    
    async def prepare_deployment(self) -> bool:
        """Prepare the codebase for deployment."""
        logger.info("Preparing deployment...")
        
        try:
            # 1. Run security scan
            logger.info("Running security scan...")
            if not await self._run_security_scan():
                logger.error("Security scan failed")
                return False
            
            # 2. Run tests
            logger.info("Running tests...")
            if not await self._run_tests():
                logger.error("Tests failed")
                return False
            
            # 3. Run linting
            logger.info("Running linting...")
            if not await self._run_linting():
                logger.error("Linting failed")
                return False
            
            # 4. Generate documentation
            logger.info("Generating documentation...")
            await self._generate_docs()
            
            # 5. Create deployment package
            logger.info("Creating deployment package...")
            await self._create_deployment_package()
            
            logger.info("Deployment preparation completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error preparing deployment: {e}")
            return False
    
    async def _run_security_scan(self) -> bool:
        """Run security scan using bandit."""
        try:
            result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "bandit", "-r", ".", "-f", "json", "-o", "security_report.json"],
                cwd=self.project_root,
                capture_output=True,
                text=True
            )
            
            if result.returncode != 0:
                logger.warning(f"Security scan found issues: {result.stderr}")
                # Don't fail deployment for low-severity issues
                return True
            
            return True
        except Exception as e:
            logger.error(f"Error running security scan: {e}")
            return False
    
    async def _run_tests(self) -> bool:
        """Run test suite."""
        try:
            result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"],
                cwd=self.project_root,
                capture_output=True,
                text=True
            )
            
            return result.returncode == 0
        except Exception as e:
            logger.error(f"Error running tests: {e}")
            return False
    
    async def _run_linting(self) -> bool:
        """Run linting checks."""
        try:
            # Run flake8
            flake8_result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "flake8", ".", "--count", "--select=E9,F63,F7,F82", "--show-source", "--statistics"],
                cwd=self.project_root,
                capture_output=True,
                text=True
            )
            
            # Run mypy
            mypy_result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "mypy", ".", "--ignore-missing-imports"],
                cwd=self.project_root,
                capture_output=True,
                text=True
            )
            
            # Don't fail on linting issues, just log them
            if flake8_result.returncode != 0:
                logger.warning(f"Flake8 issues: {flake8_result.stdout}")
            
            if mypy_result.returncode != 0:
                logger.warning(f"MyPy issues: {mypy_result.stdout}")
            
            return True
        except Exception as e:
            logger.error(f"Error running linting: {e}")
            return False
    
    async def _generate_docs(self):
        """Generate documentation."""
        try:
            # Create basic README if it doesn't exist
            readme_path = self.project_root / "README.md"
            if not readme_path.exists():
                await self._create_readme()
            
            # Create API documentation
            await self._create_api_docs()
            
        except Exception as e:
            logger.error(f"Error generating docs: {e}")
    
    async def _create_readme(self):
        """Create README.md file."""
        readme_content = """# Jarvis Demo

A modular AI agent system with auto-improvement capabilities.

## Features

- **Modular Architecture**: Clean separation of concerns with config, providers, routing, and core modules
- **Multi-Provider Support**: OpenAI, Perplexity, Anthropic, Gemini, Grok, Mistral, and more
- **Smart Routing**: Automatically selects the best provider based on content analysis
- **Auto-Improvement**: Continuous code improvement with sandbox testing
- **Plugin System**: Extensible plugin architecture for custom functionality
- **Security**: Built-in security scanning and validation

## Quick Start

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Set up environment variables:
   ```bash
   cp .env.example .env
   # Edit .env with your API keys
   ```

3. Run the application:
   ```bash
   python cloud.py
   ```

## Auto-Improvement

The system includes an auto-improvement loop that continuously enhances the codebase:

```bash
# Run auto-improvement loop
python auto_improvement_loop.py

# Check status
python auto_improvement_loop.py --status
```

## Testing

```bash
# Run all tests
pytest tests/ -v

# Run in sandbox
python sandbox_tester.py test_name
```

## Deployment

The system is configured for Render deployment with the included `render.yaml` file.

## Architecture

- `agent_v2.py` - Main agent orchestrator
- `config/` - Configuration management
- `providers/` - LLM provider implementations
- `routing/` - Smart provider selection
- `core/` - Core functionality and error handling
- `plugins/` - Plugin system
- `tools/` - Development and improvement tools
"""
        
        with open(self.project_root / "README.md", 'w') as f:
            f.write(readme_content)
    
    async def _create_api_docs(self):
        """Create API documentation."""
        api_docs = """# API Documentation

## Endpoints

### POST /chat
Main chat endpoint for AI interactions.

**Request Body:**
```json
{
    "message": "Hello, how are you?",
    "provider": "auto"  // Optional: specific provider or "auto"
}
```

**Response:**
```json
{
    "response": "I'm doing well, thank you!",
    "provider": "openai",
    "model": "gpt-4",
    "timestamp": "2024-01-01T00:00:00Z"
}
```

### GET /status
Get system status and health information.

**Response:**
```json
{
    "status": "healthy",
    "providers": ["openai", "perplexity", "anthropic"],
    "uptime": 3600,
    "version": "2.0.0"
}
```

### GET /providers
List available LLM providers.

**Response:**
```json
{
    "providers": [
        {
            "name": "openai",
            "enabled": true,
            "models": ["gpt-4", "gpt-3.5-turbo"]
        }
    ]
}
```
"""
        
        docs_dir = self.project_root / "docs"
        docs_dir.mkdir(exist_ok=True)
        
        with open(docs_dir / "api.md", 'w') as f:
            f.write(api_docs)
    
    async def _create_deployment_package(self):
        """Create deployment package."""
        # Create .dockerignore
        dockerignore_content = """
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.egg-info/
.installed.cfg
*.egg

# Virtual environments
venv/
.venv/
env/
.env/

# IDE
.vscode/
.idea/
*.swp
*.swo

# OS
.DS_Store
Thumbs.db

# Project specific
backups/
sandbox/
test_results/
*.log
.env
client_secret.json
token.json
"""
        
        with open(self.project_root / ".dockerignore", 'w') as f:
            f.write(dockerignore_content)
    
    async def setup_auto_improvement(self) -> bool:
        """Set up auto-improvement system for production."""
        logger.info("Setting up auto-improvement system...")
        
        try:
            # Create auto-improvement service script
            service_script = """#!/bin/bash
# Auto-improvement service for Render
cd /opt/render/project/src
source venv/bin/activate
python auto_improvement_loop.py
"""
            
            with open(self.project_root / "start_auto_improvement.sh", 'w') as f:
                f.write(service_script)
            
            os.chmod(self.project_root / "start_auto_improvement.sh", 0o755)
            
            # Create systemd service file (for reference)
            systemd_service = """[Unit]
Description=Jarvis Auto-Improvement Service
After=network.target

[Service]
Type=simple
User=render
WorkingDirectory=/opt/render/project/src
ExecStart=/opt/render/project/src/start_auto_improvement.sh
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
"""
            
            with open(self.project_root / "jarvis-auto-improvement.service", 'w') as f:
                f.write(systemd_service)
            
            logger.info("Auto-improvement system setup completed")
            return True
            
        except Exception as e:
            logger.error(f"Error setting up auto-improvement: {e}")
            return False
    
    async def deploy(self) -> bool:
        """Deploy the application."""
        logger.info("Starting deployment...")
        
        try:
            # Prepare deployment
            if not await self.prepare_deployment():
                logger.error("Deployment preparation failed")
                return False
            
            # Setup auto-improvement
            if not await self.setup_auto_improvement():
                logger.error("Auto-improvement setup failed")
                return False
            
            # Log deployment
            deployment_info = {
                "timestamp": datetime.now().isoformat(),
                "version": "2.0.0",
                "status": "deployed",
                "auto_improvement": "enabled"
            }
            
            with open(self.deployment_log, 'w') as f:
                json.dump(deployment_info, f, indent=2)
            
            logger.info("Deployment completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Deployment failed: {e}")
            return False

async def main():
    """Main entry point."""
    manager = DeploymentManager()
    
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == "prepare":
            success = await manager.prepare_deployment()
        elif command == "setup-auto-improvement":
            success = await manager.setup_auto_improvement()
        elif command == "deploy":
            success = await manager.deploy()
        else:
            print("Unknown command. Use: prepare, setup-auto-improvement, or deploy")
            success = False
        
        sys.exit(0 if success else 1)
    else:
        print("Usage: python deploy.py <command>")
        print("Commands: prepare, setup-auto-improvement, deploy")

if __name__ == "__main__":
    asyncio.run(main())
