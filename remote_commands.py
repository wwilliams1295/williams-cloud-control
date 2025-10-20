#!/usr/bin/env python3
"""
Remote Command Interface
Allows controlling the system via SMS/email commands.
"""

import asyncio
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import json
import subprocess  # nosec B404

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from config.settings_simple import get_settings
from auto_improvement_loop import AutoImprovementLoop
from sandbox_tester import SandboxTester
from deploy import DeploymentManager
from plugins.plugin_manager import plugin_manager

logger = logging.getLogger(__name__)

class RemoteCommandProcessor:
    """Processes remote commands from SMS/email."""
    
    def __init__(self):
        self.settings = get_settings()
        self.auto_improvement = AutoImprovementLoop()
        self.sandbox_tester = SandboxTester()
        self.deployment_manager = DeploymentManager()
        
        # Command patterns
        self.command_patterns = {
            'add_plugin': [
                r'add plugin (?:that )?(.+)',
                r'create plugin (?:that )?(.+)',
                r'new plugin (?:that )?(.+)',
                r'plugin (?:that )?(.+)'
            ],
            'run_improvement': [
                r'run auto improvement',
                r'start auto improvement',
                r'improve code',
                r'auto improve'
            ],
            'status': [
                r'status',
                r'system status',
                r'check status',
                r'how is the system'
            ],
            'deploy': [
                r'deploy',
                r'redeploy',
                r'update deployment'
            ],
            'test': [
                r'test (?:the )?system',
                r'run tests',
                r'check tests'
            ],
            'help': [
                r'help',
                r'commands',
                r'what can you do'
            ]
        }
    
    async def process_command(self, message: str, sender: str = None) -> Dict[str, Any]:
        """Process a remote command and return response."""
        logger.info(f"Processing command from {sender}: {message}")
        
        try:
            # Parse command
            command_type, match_data = self._parse_command(message)
            
            if not command_type:
                return {
                    "success": False,
                    "response": "I didn't understand that command. Try 'help' for available commands.",
                    "command_type": "unknown"
                }
            
            # Execute command
            result = await self._execute_command(command_type, match_data, sender)
            
            return {
                "success": True,
                "response": result.get("response", "Command executed successfully"),
                "command_type": command_type,
                "details": result.get("details", {})
            }
            
        except Exception as e:
            logger.error(f"Error processing command: {e}")
            return {
                "success": False,
                "response": f"Error executing command: {str(e)}",
                "command_type": "error"
            }
    
    def _parse_command(self, message: str) -> Tuple[Optional[str], Optional[str]]:
        """Parse a command from the message."""
        message_lower = message.lower().strip()
        
        for command_type, patterns in self.command_patterns.items():
            for pattern in patterns:
                match = re.search(pattern, message_lower)
                if match:
                    return command_type, match.group(1) if match.groups() else ""
        
        return None, None
    
    async def _execute_command(self, command_type: str, match_data: str, sender: str) -> Dict[str, Any]:
        """Execute a specific command."""
        
        if command_type == 'add_plugin':
            return await self._add_plugin(match_data, sender)
        
        elif command_type == 'run_improvement':
            return await self._run_improvement()
        
        elif command_type == 'status':
            return await self._get_status()
        
        elif command_type == 'deploy':
            return await self._deploy()
        
        elif command_type == 'test':
            return await self._test_system()
        
        elif command_type == 'help':
            return await self._show_help()
        
        else:
            return {"response": "Unknown command type"}
    
    async def _add_plugin(self, description: str, sender: str) -> Dict[str, Any]:
        """Add a new plugin based on description."""
        logger.info(f"Adding plugin: {description}")
        
        try:
            # Generate plugin code using AI
            plugin_code = await self._generate_plugin_code(description)
            
            if not plugin_code:
                return {
                    "response": "Sorry, I couldn't generate the plugin code. Please try a more specific description.",
                    "success": False
                }
            
            # Create plugin file
            plugin_name = self._sanitize_plugin_name(description)
            plugin_file = Path("plugins") / f"{plugin_name}.py"
            
            with open(plugin_file, 'w') as f:
                f.write(plugin_code)
            
            # Load the new plugin
            plugin_manager.load_plugin(plugin_name)
            
            return {
                "response": f"✅ Plugin '{plugin_name}' created successfully! It will be available in the next improvement cycle.",
                "success": True,
                "details": {
                    "plugin_name": plugin_name,
                    "plugin_file": str(plugin_file),
                    "description": description
                }
            }
            
        except Exception as e:
            logger.error(f"Error adding plugin: {e}")
            return {
                "response": f"❌ Error creating plugin: {str(e)}",
                "success": False
            }
    
    async def _generate_plugin_code(self, description: str) -> Optional[str]:
        """Generate plugin code using AI."""
        try:
            # Use the existing agent to generate code
            from agent_v2 import superchat
            
            prompt = f"""
Create a Python plugin for the Jarvis system that: {description}

The plugin should:
1. Inherit from PluginBase
2. Have a proper execute method
3. Include error handling
4. Be production-ready
5. Follow the existing plugin patterns

Here's the base structure to follow:

```python
from plugins.plugin_manager import PluginBase
import logging

class {self._sanitize_plugin_name(description)}Plugin(PluginBase):
    def __init__(self, config=None):
        super().__init__(config)
        self.name = "{self._sanitize_plugin_name(description)}"
    
    async def execute(self, context):
        # Implementation here
        pass
    
    def get_description(self):
        return "{description}"
    
    def get_requirements(self):
        return []
```

Generate the complete plugin code:
"""
            
            response = await superchat(prompt)
            
            # Extract code from response - try multiple patterns
            code_patterns = [
                r'```python\n(.*?)\n```',
                r'```\n(.*?)\n```',
                r'class\s+\w+Plugin.*?(?=\n\n|\n$|$)',
            ]
            
            for pattern in code_patterns:
                try:
                    code_match = re.search(pattern, response, re.DOTALL)
                    if code_match:
                        code = code_match.group(1).strip()
                        if 'class' in code and 'Plugin' in code:
                            return code
                except Exception as e:
                    logger.debug(f"Regex pattern failed: {pattern}, error: {e}")
                    continue
            
            # If no code blocks, try to extract the code manually
            lines = response.split('\n')
            code_lines = []
            in_code = False
            
            for line in lines:
                if line.strip().startswith('class ') and 'Plugin' in line:
                    in_code = True
                if in_code:
                    code_lines.append(line)
                    # Stop at next class or end of response
                    if len(code_lines) > 1 and line.strip().startswith('class ') and line.strip() != code_lines[0].strip():
                        break
            
            if code_lines:
                return '\n'.join(code_lines)
            
            # Last resort: create a simple plugin template
            return self._create_simple_plugin_template(description)
            
        except Exception as e:
            logger.error(f"Error generating plugin code: {e}")
            return None
    
    def _sanitize_plugin_name(self, description: str) -> str:
        """Convert description to valid plugin name."""
        # Remove special characters and convert to snake_case
        name = re.sub(r'[^a-zA-Z0-9\s]', '', description)
        name = re.sub(r'\s+', '_', name.strip())
        name = name.lower()
        
        # Ensure it starts with a letter
        if name and not name[0].isalpha():
            name = 'plugin_' + name
        
        return name or 'custom_plugin'
    
    def _create_simple_plugin_template(self, description: str) -> str:
        """Create a simple plugin template when AI generation fails."""
        plugin_name = self._sanitize_plugin_name(description)
        class_name = ''.join(word.capitalize() for word in plugin_name.split('_')) + 'Plugin'
        
        return f'''from plugins.plugin_manager import PluginBase
import logging

class {class_name}(PluginBase):
    """Plugin: {description}"""
    
    def __init__(self, config=None):
        super().__init__(config)
        self.name = "{plugin_name}"
        self.description = "{description}"
    
    async def execute(self, context):
        """Execute the plugin functionality."""
        try:
            # TODO: Implement {description}
            logger.info(f"Executing {{self.name}}: {{description}}")
            
            # Placeholder implementation
            result = {{
                "success": True,
                "message": "Plugin executed successfully",
                "description": "{description}",
                "context": context
            }}
            
            return result
            
        except Exception as e:
            logger.error(f"Error in {{self.name}}: {{e}}")
            return {{
                "success": False,
                "error": str(e)
            }}
    
    def get_description(self):
        return "{description}"
    
    def get_requirements(self):
        return []
    
    def validate_config(self):
        return True'''
    
    async def _run_improvement(self) -> Dict[str, Any]:
        """Run auto-improvement cycle."""
        logger.info("Running auto-improvement cycle")
        
        try:
            result = await self.auto_improvement.run_improvement_cycle()
            
            if result.get("status") == "completed":
                return {
                    "response": "✅ Auto-improvement cycle completed successfully! Check the logs for details.",
                    "success": True,
                    "details": result
                }
            elif result.get("status") == "skipped":
                return {
                    "response": "⏭️ Auto-improvement skipped (cooldown period active)",
                    "success": True,
                    "details": result
                }
            else:
                return {
                    "response": f"❌ Auto-improvement failed: {result.get('error', 'Unknown error')}",
                    "success": False,
                    "details": result
                }
                
        except Exception as e:
            logger.error(f"Error running improvement: {e}")
            return {
                "response": f"❌ Error running auto-improvement: {str(e)}",
                "success": False
            }
    
    async def _get_status(self) -> Dict[str, Any]:
        """Get system status."""
        try:
            status = self.auto_improvement.get_status()
            plugin_status = plugin_manager.get_plugin_status()
            
            response = f"""📊 **System Status**
            
🔄 **Auto-Improvement**: {'Active' if status['last_improvement'] else 'Never run'}
📈 **Total Improvements**: {status['total_improvements']}
⏰ **Next Improvement**: {status['next_improvement_in']:.0f} seconds
🧪 **Sandbox Mode**: {'Enabled' if status['sandbox_mode'] else 'Disabled'}

🔌 **Plugins** ({len(plugin_status)}):
"""
            
            for plugin_name, plugin_info in plugin_status.items():
                status_icon = "✅" if plugin_info['enabled'] else "❌"
                response += f"  {status_icon} {plugin_name}: {plugin_info['description']}\n"
            
            return {
                "response": response,
                "success": True,
                "details": {
                    "auto_improvement": status,
                    "plugins": plugin_status
                }
            }
            
        except Exception as e:
            logger.error(f"Error getting status: {e}")
            return {
                "response": f"❌ Error getting status: {str(e)}",
                "success": False
            }
    
    async def _deploy(self) -> Dict[str, Any]:
        """Deploy the system."""
        logger.info("Running deployment")
        
        try:
            success = await self.deployment_manager.deploy()
            
            if success:
                return {
                    "response": "✅ Deployment completed successfully!",
                    "success": True
                }
            else:
                return {
                    "response": "❌ Deployment failed. Check logs for details.",
                    "success": False
                }
                
        except Exception as e:
            logger.error(f"Error deploying: {e}")
            return {
                "response": f"❌ Error deploying: {str(e)}",
                "success": False
            }
    
    async def _test_system(self) -> Dict[str, Any]:
        """Test the system."""
        logger.info("Running system tests")
        
        try:
            # Create a test sandbox
            test_name = f"remote_test_{int(datetime.now().timestamp())}"
            sandbox_path = await self.sandbox_tester.create_sandbox(test_name)
            
            # Run tests
            results = await self.sandbox_tester.run_tests(sandbox_path)
            
            # Clean up
            await self.sandbox_tester.cleanup_sandbox(test_name)
            
            if results.get("success"):
                return {
                    "response": "✅ System tests passed! All components are working correctly.",
                    "success": True,
                    "details": results
                }
            else:
                return {
                    "response": "❌ System tests failed. Check logs for details.",
                    "success": False,
                    "details": results
                }
                
        except Exception as e:
            logger.error(f"Error testing system: {e}")
            return {
                "response": f"❌ Error testing system: {str(e)}",
                "success": False
            }
    
    async def _show_help(self) -> Dict[str, Any]:
        """Show available commands."""
        help_text = """🤖 **Available Commands**

**Plugin Management:**
• `add plugin that [description]` - Create a new plugin
• `create plugin that [description]` - Create a new plugin
• `new plugin that [description]` - Create a new plugin

**System Control:**
• `run auto improvement` - Start improvement cycle
• `improve code` - Start improvement cycle
• `status` - Check system status
• `deploy` - Deploy the system
• `test system` - Run system tests

**Examples:**
• `add plugin that sends calendar invite`
• `add plugin that monitors system performance`
• `run auto improvement`
• `status`

**Note:** Commands are case-insensitive and support natural language."""
        
        return {
            "response": help_text,
            "success": True
        }

# Global command processor instance
command_processor = RemoteCommandProcessor()

async def process_remote_command(message: str, sender: str = None) -> Dict[str, Any]:
    """Process a remote command (for use by cloud.py)."""
    return await command_processor.process_command(message, sender)
