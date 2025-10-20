# Jarvis AI Assistant

A comprehensive AI assistant system with SMS/Email integration, file management, and auto-improvement capabilities.

## Features

- **Multi-LLM Support**: OpenAI, Perplexity, Anthropic, Gemini, Grok, Mistral
- **SMS/Email Integration**: Google Voice via Gmail API
- **File Management**: Create and manage PPTX, PDF, Excel files
- **Auto-Improvement**: Self-evolving codebase with AI-powered enhancements
- **Plugin System**: Extensible architecture for new functionalities
- **Cloud Storage**: AWS S3 integration for persistent file storage

## Quick Start

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Set Environment Variables**:
   ```bash
   # Required API Keys
   export OPENAI_API_KEY="your-key"
   export PPLX_API_KEY="your-key"
   
   # Optional API Keys
   export ANTHROPIC_API_KEY="your-key"
   export GOOGLE_API_KEY="your-key"
   export XAI_API_KEY="your-key"
   export MISTRAL_API_KEY="your-key"
   
   # Email Configuration
   export SMTP_HOST="smtp.gmail.com"
   export SMTP_PORT="587"
   export SMTP_USER="your-email@gmail.com"
   export SMTP_PASS="your-app-password"
   export FROM_EMAIL="your-email@gmail.com"
   
   # Cloud Storage (AWS S3)
   export STORAGE_TYPE="s3"
   export STORAGE_BUCKET="your-bucket"
   export STORAGE_REGION="us-east-2"
   export AWS_ACCESS_KEY_ID="your-key"
   export AWS_SECRET_ACCESS_KEY="your-secret"
   ```

3. **Run the Application**:
   ```bash
   python cloud.py
   ```

## Deployment

The application is configured for deployment on Render.com. See `render.yaml` for configuration details.

## Auto-Improvement System

The system includes an auto-improvement loop that continuously enhances the codebase:

```bash
# Run auto-improvement once
python scripts/auto_improvement_loop.py --once

# Run continuous improvement
python scripts/auto_improvement_loop.py
```

## Project Structure

- `agent.py` - Core AI routing and LLM integration
- `cloud.py` - Main FastAPI application with SMS/Email handling
- `plugins/` - Plugin system for extensible functionality
- `scripts/` - Auto-improvement and deployment scripts
- `tools/` - Development and maintenance tools
- `backups/` - Automated backups of previous versions

## Configuration

The system uses environment variables for configuration. See the environment variables section above for required settings.

## Security

- Phone number allowlist for SMS access
- Email allowlist for email access
- Secure API key management
- Input validation and sanitization

## License

Private project - All rights reserved.
