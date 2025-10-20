#!/usr/bin/env python3
"""
AI Plugin Integration
Shows how AI can discover, understand, and use plugins.
"""

import asyncio
import logging
from typing import Dict, Any, List
from pathlib import Path
import sys

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from plugins.plugin_manager import plugin_manager
from agent_v2 import superchat

logger = logging.getLogger(__name__)

class AIPluginIntegration:
    """Integrates AI with the plugin system."""
    
    def __init__(self):
        self.plugin_manager = plugin_manager
        self.load_all_plugins()
    
    def load_all_plugins(self):
        """Load all available plugins."""
        self.plugin_manager.load_all_plugins()
        logger.info(f"Loaded {len(self.plugin_manager.plugins)} plugins")
    
    async def get_plugin_capabilities(self) -> str:
        """Get a description of all available plugins for the AI."""
        capabilities = []
        
        for plugin_name, plugin in self.plugin_manager.plugins.items():
            if plugin.enabled:
                capabilities.append(f"""
**{plugin_name}**:
- Description: {plugin.get_description()}
- Requirements: {', '.join(plugin.get_requirements()) or 'None'}
- Status: {'Enabled' if plugin.enabled else 'Disabled'}
""")
        
        return "\n".join(capabilities) if capabilities else "No plugins available"
    
    async def ai_can_use_plugins(self, user_request: str) -> bool:
        """Check if the user request can be handled by available plugins."""
        capabilities = await self.get_plugin_capabilities()
        
        prompt = f"""
You are an AI assistant with access to the following plugins:

{capabilities}

User request: "{user_request}"

Can any of these plugins help fulfill this request? Answer with:
- YES: [plugin_name] - [brief explanation]
- NO: [reason why no plugin can help]

Be specific about which plugin and how it would help.
"""
        
        response = await superchat(prompt)
        return response
    
    async def execute_plugin_for_ai(self, plugin_name: str, context: Dict[str, Any] = None) -> Dict[str, Any]:
        """Execute a specific plugin for the AI."""
        if plugin_name not in self.plugin_manager.plugins:
            return {"success": False, "error": f"Plugin '{plugin_name}' not found"}
        
        plugin = self.plugin_manager.plugins[plugin_name]
        if not plugin.enabled:
            return {"success": False, "error": f"Plugin '{plugin_name}' is disabled"}
        
        try:
            result = await plugin.execute(context or {})
            return result
        except Exception as e:
            logger.error(f"Error executing plugin {plugin_name}: {e}")
            return {"success": False, "error": str(e)}
    
    async def ai_with_plugin_access(self, user_request: str) -> str:
        """AI response with plugin capabilities."""
        # First, check if plugins can help
        plugin_analysis = await self.ai_can_use_plugins(user_request)
        
        # Get plugin capabilities for context
        capabilities = await self.get_plugin_capabilities()
        
        # Create enhanced prompt with plugin context
        enhanced_prompt = f"""
You are an AI assistant with access to these plugins:

{capabilities}

User request: "{user_request}"

Plugin analysis: {plugin_analysis}

Instructions:
1. If a plugin can help, mention it and explain how
2. If no plugin is needed, provide a regular AI response
3. Always be helpful and specific about plugin capabilities
4. If you suggest using a plugin, explain what it would do

Respond naturally and helpfully:
"""
        
        response = await superchat(enhanced_prompt)
        return response

# Global instance
ai_plugin_integration = AIPluginIntegration()

async def demo_plugin_integration():
    """Demonstrate how AI can use plugins."""
    print("🤖 AI Plugin Integration Demo")
    print("=" * 50)
    
    # Test 1: AI discovers plugins
    print("\n1. AI discovering available plugins:")
    capabilities = await ai_plugin_integration.get_plugin_capabilities()
    print(capabilities)
    
    # Test 2: AI analyzes if plugins can help
    print("\n2. AI analyzing if plugins can help with 'send a calendar invite':")
    analysis = await ai_plugin_integration.ai_can_use_plugins("send a calendar invite")
    print(analysis)
    
    # Test 3: AI with plugin context
    print("\n3. AI response with plugin context:")
    response = await ai_plugin_integration.ai_with_plugin_access("I need to send a calendar invite for a meeting tomorrow")
    print(response)
    
    # Test 4: Execute plugin directly
    print("\n4. Executing plugin directly:")
    result = await ai_plugin_integration.execute_plugin_for_ai("sends_calendar_invite", {
        "title": "Team Meeting",
        "start_time": "2024-01-20T10:00:00",
        "attendees": ["team@example.com"]
    })
    print(f"Plugin result: {result}")

if __name__ == "__main__":
    asyncio.run(demo_plugin_integration())
