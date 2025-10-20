# Jarvis AI Assistant

A sophisticated AI assistant with Google Voice SMS integration, file management, calendar invites, and auto-improvement capabilities.

## 🏗️ Project Structure

```
jarvis-demo/
├── main.py                     # Main entry point
├── requirements.txt            # Python dependencies
├── README.md                   # This file
│
├── src/                        # Source code
│   ├── core/                   # Core functionality
│   │   ├── agent.py           # AI agent and LLM routing
│   │   ├── functions.py       # Utility functions
│   │   ├── mailer.py          # Email functionality
│   │   └── onedrive_device_login.py
│   │
│   ├── api/                    # API layer
│   │   └── cloud.py           # FastAPI application
│   │
│   ├── memory/                 # Memory system
│   │   └── memory_system.py   # Conversation memory
│   │
│   ├── plugins/                # Plugin system
│   │   ├── file_edit.py       # File editing plugin
│   │   ├── send_pdf.py        # PDF generation plugin
│   │   ├── calendar_invite_plugin.py
│   │   ├── edgar_pull.py      # SEC data plugin
│   │   └── ...
│   │
│   ├── ai/                     # AI/ML scripts
│   │   ├── auto_improvement_loop.py
│   │   ├── advanced_auto_improvement.py
│   │   ├── creative_ai_evolution.py
│   │   └── ...
│   │
│   └── tools/                  # Development tools
│       ├── auto_improver.py
│       ├── mypy_autofix.py
│       └── ...
│
├── config/                     # Configuration files
│   ├── deployment/            # Deployment configs
│   │   ├── render.yaml        # Render deployment
│   │   └── runtime.txt        # Python runtime
│   │
│   └── security/              # Security configs
│       ├── bandit.yaml        # Security linting
│       ├── mypy.ini          # Type checking
│       ├── client_secret.json # Gmail credentials
│       └── token.json         # Gmail token
│
├── tests/                      # Test suite
│   ├── unit/                  # Unit tests
│   ├── integration/           # Integration tests
│   └── e2e/                   # End-to-end tests
│
├── scripts/                    # Utility scripts
│   ├── deployment/            # Deployment scripts
│   ├── maintenance/           # Maintenance scripts
│   └── testing/               # Testing scripts
│
├── docs/                      # Documentation
│   ├── api/                   # API documentation
│   ├── deployment/            # Deployment guides
│   └── IMPROVEMENT_SUMMARY.md
│
├── data/                      # Data storage
│   ├── cache/                 # Cache files
│   ├── exports/               # Export files
│   ├── backups/               # Backup files
│   ├── storage/               # General storage
│   └── conversations.db       # Conversation database
│
└── logs/                      # Log files
    └── *.log
```

## 🚀 Quick Start

### Installation

1. Clone the repository:
```bash
git clone https://github.com/wwilliams1295/williams-cloud-control.git
cd jarvis-demo
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up environment variables:
```bash
cp .env.example .env
# Edit .env with your API keys
```

### Running the Application

```bash
# Start the API server
python main.py

# Or with specific options
python main.py --start-api        # Start API server
python main.py --memory-stats     # Show memory statistics
python main.py --test             # Run tests
python main.py --help             # Show help
```

## 🔧 Features

### Core Capabilities
- **Multi-LLM Routing**: OpenAI, Perplexity, Anthropic, Gemini, Grok, Mistral
- **Google Voice SMS**: Integration via Gmail API
- **File Management**: Create and edit PPTX, PDF, Excel files
- **Calendar Integration**: Generate and send calendar invites
- **SEC Data**: Pull Edgar filings and financial data
- **Auto-Improvement**: Self-evolving code and capabilities

### Memory System
- **Conversation Memory**: Persistent storage of all conversations
- **Context Awareness**: AI remembers previous interactions
- **Search Capabilities**: Find past conversations by content
- **User-Specific Memory**: Separate memory for each user

### Plugin System
- **File Editing**: Create and modify various file types
- **PDF Generation**: Generate and send PDF documents
- **Calendar Invites**: Create and send calendar invitations
- **SEC Data Pulling**: Retrieve company filing data
- **System Monitoring**: Track performance and metrics

## 📱 Usage

### SMS Commands
Send SMS to your Google Voice number:

- `what plugins exist` - List available plugins
- `what files are on server` - Show server files
- `run auto improvement` - Start improvement system
- `system status` - Check system health
- `memory stats` - View conversation statistics
- `search memory [query]` - Find past conversations
- `show conversation history` - View recent chat history

### API Endpoints

- `GET /` - Health check
- `GET /debug/gmail_status` - Gmail integration status
- `GET /debug/gmail_messages` - Recent Gmail messages
- `POST /debug/test_gmail_send` - Test Gmail sending

## 🔒 Security

- **API Key Management**: Secure environment variable storage
- **User Authentication**: Phone number and email allowlists
- **Data Encryption**: Secure conversation storage
- **Access Control**: Role-based permissions

## 🚀 Deployment

### Render Deployment

1. Connect your GitHub repository to Render
2. Set environment variables in Render dashboard
3. Deploy using the `config/deployment/render.yaml` configuration

### Environment Variables

Required:
- `OPENAI_API_KEY` - OpenAI API key
- `PPLX_API_KEY` - Perplexity API key
- `GMAIL_CLIENT_SECRET_JSON` - Gmail OAuth credentials
- `GMAIL_TOKEN_JSON` - Gmail access token

Optional:
- `ANTHROPIC_API_KEY` - Anthropic API key
- `GOOGLE_API_KEY` - Google API key
- `XAI_API_KEY` - Grok API key
- `MISTRAL_API_KEY` - Mistral API key
- `AWS_ACCESS_KEY_ID` - AWS S3 access
- `AWS_SECRET_ACCESS_KEY` - AWS S3 secret

## 🧪 Testing

```bash
# Run all tests
python main.py --test

# Run specific test suites
python -m pytest tests/unit/
python -m pytest tests/integration/
python -m pytest tests/e2e/
```

## 📊 Monitoring

- **Memory Statistics**: Track conversation counts and user activity
- **System Health**: Monitor API status and integrations
- **Performance Metrics**: Track response times and usage
- **Error Logging**: Comprehensive error tracking and reporting

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For support and questions:
- Create an issue on GitHub
- Check the documentation in `docs/`
- Review the logs in `logs/`

---

**Jarvis AI Assistant** - Your intelligent cloud control system 🤖✨