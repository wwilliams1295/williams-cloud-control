#!/usr/bin/env python3
"""
Code Implementation Engine
Actually implements the AI-generated improvements by modifying the codebase.
"""

import asyncio
import logging
import os
import sys
import json
import re
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from config.settings_simple import get_settings
from agent_v2 import superchat

logger = logging.getLogger(__name__)

class CodeImplementer:
    """Actually implements AI-generated improvements in the codebase."""
    
    def __init__(self):
        self.settings = get_settings()
        self.project_root = Path(__file__).parent.parent
        self.backup_dir = self.project_root / "backups" / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
    async def implement_improvement(self, idea: Dict[str, Any]) -> Dict[str, Any]:
        """Implement a specific improvement idea in the codebase."""
        logger.info(f"🔧 Implementing: {idea.get('title', 'Unknown')}")
        
        try:
            # Create backup
            await self._create_backup()
            
            # Generate specific code changes
            code_changes = await self._generate_code_changes(idea)
            
            # Apply the changes
            result = await self._apply_code_changes(code_changes, idea)
            
            return {
                "success": True,
                "idea": idea,
                "changes_applied": result,
                "backup_location": str(self.backup_dir)
            }
            
        except Exception as e:
            logger.error(f"❌ Error implementing improvement: {e}")
            return {
                "success": False,
                "error": str(e),
                "idea": idea
            }
    
    async def _create_backup(self):
        """Create backup of current state."""
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        
        # Copy important files
        important_files = [
            "agent.py", "agent_v2.py", "cloud.py", "file_manager.py",
            "remote_commands.py", "ai_plugin_integration.py", "cloud_storage.py",
            "core/", "providers/", "routing/", "plugins/"
        ]
        
        for file_path in important_files:
            src = self.project_root / file_path
            if src.exists():
                if src.is_file():
                    shutil.copy2(src, self.backup_dir / file_path)
                elif src.is_dir():
                    shutil.copytree(src, self.backup_dir / file_path, dirs_exist_ok=True)
        
        logger.info(f"📁 Backup created: {self.backup_dir}")
    
    async def _generate_code_changes(self, idea: Dict[str, Any]) -> Dict[str, Any]:
        """Generate specific code changes for an idea."""
        
        # Use string formatting instead of f-string to avoid brace issues
        prompt = """
        You are an expert Python developer. Generate specific, implementable code changes for this Jarvis AI system improvement:
        
        Title: {title}
        Description: {description}
        Impact: {impact}
        Improvement Type: {improvement_type}
        
        Current Jarvis codebase structure:
        - agent.py, agent_v2.py, cloud.py (main application files)
        - core/ (core functionality: capabilities.py, execute.py, errors.py, planner.py)
        - providers/ (LLM API clients: openai.py, perplexity.py, anthropic.py, etc.)
        - routing/ (request routing: router.py, provider_registry.py, web_scorer.py)
        - plugins/ (extensible plugins: plugin_manager.py, calendar_plugin.py, etc.)
        - config/ (configuration: settings.py, settings_simple.py)
        - scripts/ (utility scripts)
        
        Generate SPECIFIC code changes that will ACTUALLY IMPROVE the Jarvis system:
        1. Add new functions, classes, or modules that enhance functionality
        2. Improve existing code with better error handling, performance, or features
        3. Add new API endpoints, commands, or user-facing features
        4. Enhance the plugin system, AI capabilities, or user experience
        5. Add monitoring, logging, or debugging capabilities
        6. Improve security, validation, or data handling
        
        Focus on PRACTICAL improvements that users will notice and benefit from.
        
        Return as JSON with:
        {{
            "files_to_modify": [
                {{
                    "path": "file_path.py",
                    "changes": [
                        {{
                            "type": "add_function|modify_function|add_class|add_import|add_endpoint",
                            "location": "after_line_X|before_line_X|end_of_file|beginning_of_file",
                            "code": "actual Python code here with proper imports and error handling"
                        }}
                    ]
                }}
            ],
            "files_to_create": [
                {{
                    "path": "new_file.py",
                    "content": "complete file content here with proper structure"
                }}
            ],
            "description": "What this implementation does for the Jarvis system"
        }}
        
        Provide REAL, WORKING CODE that enhances the Jarvis AI assistant.
        """.format(
            title=idea.get('title', 'Unknown'),
            description=idea.get('description', 'No description'),
            impact=idea.get('impact', 'Unknown'),
            improvement_type=idea.get('improvement_type', 'Unknown')
        )
        
        try:
            response = await superchat(prompt)
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                return json.loads(json_match.group(0))
            else:
                return self._generate_fallback_changes(idea)
        except Exception as e:
            logger.error(f"Error generating code changes: {e}")
            return self._generate_fallback_changes(idea)
    
    def _generate_fallback_changes(self, idea: Dict[str, Any]) -> Dict[str, Any]:
        """Generate fallback code changes if AI fails."""
        return {
            "files_to_modify": [
                {
                    "path": "agent_v2.py",
                    "changes": [
                        {
                            "type": "add_function",
                            "location": "end_of_file",
                            "code": f"""
# Auto-improvement: {idea.get('title', 'Unknown')}
async def {idea.get('title', 'unknown').lower().replace(' ', '_').replace('-', '_')}():
    \"\"\"{idea.get('description', 'Auto-generated improvement')}\"\"\"
    logger.info("Executing auto-improvement: {idea.get('title', 'Unknown')}")
    return {{"success": True, "message": "Improvement executed"}}
"""
                        }
                    ]
                }
            ],
            "files_to_create": [],
            "description": f"Added function for {idea.get('title', 'Unknown')}"
        }
    
    async def _apply_code_changes(self, code_changes: Dict[str, Any], idea: Dict[str, Any]) -> Dict[str, Any]:
        """Apply the generated code changes to the codebase."""
        result = {
            "files_modified": [],
            "files_created": [],
            "errors": []
        }
        
        # Create new files
        for file_spec in code_changes.get("files_to_create", []):
            try:
                file_path = self.project_root / file_spec["path"]
                file_path.parent.mkdir(parents=True, exist_ok=True)
                
                with open(file_path, "w") as f:
                    f.write(file_spec["content"])
                
                result["files_created"].append(str(file_path))
                logger.info(f"✅ Created file: {file_path}")
            except Exception as e:
                error_msg = f"Error creating {file_spec['path']}: {e}"
                result["errors"].append(error_msg)
                logger.error(error_msg)
        
        # Modify existing files
        for file_spec in code_changes.get("files_to_modify", []):
            try:
                file_path = self.project_root / file_spec["path"]
                if not file_path.exists():
                    logger.warning(f"File not found: {file_path}")
                    continue
                
                # Read current content
                with open(file_path, "r") as f:
                    lines = f.readlines()
                
                # Apply changes
                new_lines = lines.copy()
                for change in file_spec.get("changes", []):
                    new_lines = self._apply_single_change(new_lines, change)
                
                # Write back
                with open(file_path, "w") as f:
                    f.writelines(new_lines)
                
                result["files_modified"].append(str(file_path))
                logger.info(f"✅ Modified file: {file_path}")
                
            except Exception as e:
                error_msg = f"Error modifying {file_spec['path']}: {e}"
                result["errors"].append(error_msg)
                logger.error(error_msg)
        
        return result
    
    def _apply_single_change(self, lines: List[str], change: Dict[str, Any]) -> List[str]:
        """Apply a single change to a file."""
        change_type = change.get("type", "add_function")
        location = change.get("location", "end_of_file")
        code = change.get("code", "")
        
        if change_type == "add_function":
            if location == "end_of_file":
                lines.append(f"\n{code}\n")
            elif location.startswith("after_line_"):
                line_num = int(location.split("_")[-1]) - 1
                lines.insert(line_num + 1, f"\n{code}\n")
            elif location == "beginning_of_file":
                lines.insert(0, f"{code}\n")
        
        elif change_type == "add_import":
            if location == "beginning_of_file":
                lines.insert(0, f"{code}\n")
            else:
                lines.append(f"{code}\n")
        
        elif change_type == "add_class":
            if location == "end_of_file":
                lines.append(f"\n{code}\n")
            else:
                line_num = int(location.split("_")[-1]) - 1
                lines.insert(line_num + 1, f"\n{code}\n")
        
        return lines

async def main():
    """Test the code implementer."""
    implementer = CodeImplementer()
    
    # Test with a simple idea
    test_idea = {
        "title": "Enhanced Logging System",
        "description": "Add better logging capabilities to track AI improvements",
        "impact": "medium",
        "improvement_type": "logging"
    }
    
    result = await implementer.implement_improvement(test_idea)
    print(json.dumps(result, indent=2))

if __name__ == "__main__":
    asyncio.run(main())
