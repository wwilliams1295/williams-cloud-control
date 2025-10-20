# Jarvis AI Assistant

An intelligent AI assistant with modular architecture, cloud storage, and auto-improvement capabilities.

## 🏗️ Project Structure

```
jarvis-demo/
├── config/                 # Configuration files
│   ├── __init__.py
│   ├── settings.py
│   ├── settings_simple.py
│   ├── .env.example
│   ├── bandit.yaml
│   ├── mypy.ini
│   └── .ruff.toml
├── core/                   # Core functionality
│   ├── __init__.py
│   ├── capabilities.py
│   ├── execute.py
│   ├── errors.py
│   └── planner.py
├── providers/              # LLM API providers
│   ├── __init__.py
│   ├── base.py
│   ├── openai.py
│   ├── perplexity.py
│   ├── anthropic.py
│   ├── gemini.py
│   ├── grok.py
│   ├── mistral.py
│   └── openai_compatible.py
├── routing/                # Request routing logic
│   ├── __init__.py
│   ├── provider_registry.py
│   ├── router.py
│   └── web_scorer.py
├── plugins/                # Extensible plugin system
│   ├── __init__.py
│   ├── plugin_manager.py
│   ├── loader.py
│   ├── calendar_plugin.py
│   ├── system_monitor_plugin.py
│   └── sends_calendar_invite.py
├── tools/                  # Development tools
│   ├── auto_improver.py
│   ├── mypy_autofix.py
│   ├── self_loop.py
│   ├── safe-apply-patch.sh
│   └── policy.yaml
├── scripts/                # Utility scripts
│   ├── deploy.py
│   ├── migrate_to_v2.py
│   ├── auto_improvement_loop.py
│   └── sandbox_tester.py
├── tests/                  # Test suites
│   ├── unit/
│   ├── integration/
│   └── test_plugins.py
├── storage/                # Local storage and patches
│   └── patch.diff
├── docs/                   # Documentation
│   └── IMPROVEMENT_SUMMARY.md
├── data/                   # Data files
├── agent.py               # Legacy agent (deprecated)
├── agent_v2.py            # New modular agent
├── cloud.py               # FastAPI web service
├── cloud_storage.py       # Cloud storage abstraction
├── file_manager.py        # File operations
├── remote_commands.py     # SMS/email command processing
├── ai_plugin_integration.py # AI plugin integration
├── functions.py           # Utility functions
├── mailer.py              # Email functionality
├── onedrive_device_login.py # OneDrive integration
├── requirements.txt       # Python dependencies
├── runtime.txt            # Python version
├── render.yaml            # Render deployment config
└── README.md              # This file
```

## 🚀 Features

- **Modular Architecture**: Clean separation of concerns
- **Multiple LLM Providers**: OpenAI, Perplexity, Anthropic, Gemini, Grok, Mistral
- **Cloud Storage**: AWS S3, Google Cloud Storage, Azure Blob
- **Plugin System**: Extensible functionality
- **Auto-Improvement**: Self-evolving codebase
- **File Management**: PPT, Excel, PDF creation and editing
- **Remote Control**: SMS/email command interface
- **SEC Filings**: Automated 10-K, 10-Q retrieval
- **Email Integration**: Gmail/Google Voice support

## 🛠️ Setup

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure environment**:
   ```bash
   cp config/.env.example .env
   # Edit .env with your API keys
   ```

3. **Run the application**:
   ```bash
   python cloud.py
   ```

## 📱 Usage

### SMS/Email Commands
- `help` - Show available commands
- `status` - Check system status
- `create ppt "Title" with bullets` - Create PowerPoint
- `pull 10-k AAPL` - Get Apple's 10-K filing
- `add plugin "description"` - Add new plugin

### File Operations
- Create and edit PowerPoint presentations
- Generate Excel spreadsheets
- Download SEC filings as PDFs
- Send files via email with download links

## 🌐 Deployment

Deploy to Render using the included `render.yaml` configuration.

## 🔧 Development

- **Linting**: `ruff check .`
- **Type checking**: `mypy .`
- **Security**: `bandit -r .`
- **Testing**: `pytest tests/`

## 📄 License

MIT License
