"""
Plugin Management System
Handles loading, managing, and executing plugins for the auto-improvement system.
"""

import importlib
import inspect
import logging
import os
import sys
from pathlib import Path
from typing import Dict, List, Any, Optional, Type, Callable
from abc import ABC, abstractmethod
import json

logger = logging.getLogger(__name__)

class PluginBase(ABC):
    """Base class for all plugins."""
    
    def __init__(self, config: Dict[str, Any] = None):
        self.config = config or {}
        self.name = self.__class__.__name__
        self.enabled = True
    
    @abstractmethod
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Execute the plugin with given context."""
        pass
    
    @abstractmethod
    def get_description(self) -> str:
        """Get plugin description."""
        pass
    
    def validate_config(self) -> bool:
        """Validate plugin configuration."""
        return True
    
    def get_requirements(self) -> List[str]:
        """Get list of required dependencies."""
        return []

class PluginManager:
    """Manages plugin loading and execution."""
    
    def __init__(self, plugins_dir: str = "plugins"):
        self.plugins_dir = Path(plugins_dir)
        self.plugins: Dict[str, PluginBase] = {}
        self.plugin_configs = {}
        self.load_plugin_configs()
    
    def load_plugin_configs(self):
        """Load plugin configurations from config file."""
        config_file = self.plugins_dir / "plugin_configs.json"
        if config_file.exists():
            try:
                with open(config_file, 'r') as f:
                    self.plugin_configs = json.load(f)
            except Exception as e:
                logger.error(f"Error loading plugin configs: {e}")
                self.plugin_configs = {}
        else:
            # Create default config
            self.plugin_configs = {
                "code_analyzer": {"enabled": True, "priority": 1},
                "test_generator": {"enabled": True, "priority": 2},
                "documentation_generator": {"enabled": True, "priority": 3},
                "security_scanner": {"enabled": True, "priority": 4},
                "performance_optimizer": {"enabled": True, "priority": 5}
            }
            self.save_plugin_configs()
    
    def save_plugin_configs(self):
        """Save plugin configurations to file."""
        config_file = self.plugins_dir / "plugin_configs.json"
        try:
            with open(config_file, 'w') as f:
                json.dump(self.plugin_configs, f, indent=2)
        except Exception as e:
            logger.error(f"Error saving plugin configs: {e}")
    
    def discover_plugins(self) -> List[str]:
        """Discover available plugins in the plugins directory."""
        plugins = []
        for file_path in self.plugins_dir.glob("*.py"):
            if file_path.name != "__init__.py" and not file_path.name.startswith("_"):
                plugins.append(file_path.stem)
        return plugins
    
    def load_plugin(self, plugin_name: str) -> Optional[PluginBase]:
        """Load a specific plugin."""
        try:
            # Add plugins directory to path
            if str(self.plugins_dir) not in sys.path:
                sys.path.insert(0, str(self.plugins_dir))
            
            # Import the plugin module
            module = importlib.import_module(plugin_name)
            
            # Find plugin classes
            plugin_classes = []
            for name, obj in inspect.getmembers(module):
                if (inspect.isclass(obj) and 
                    issubclass(obj, PluginBase) and 
                    obj != PluginBase):
                    plugin_classes.append(obj)
            
            if not plugin_classes:
                logger.warning(f"No plugin classes found in {plugin_name}")
                return None
            
            # Use the first plugin class found
            plugin_class = plugin_classes[0]
            
            # Get config for this plugin
            config = self.plugin_configs.get(plugin_name, {})
            
            # Create plugin instance
            plugin = plugin_class(config)
            
            # Validate plugin
            if not plugin.validate_config():
                logger.error(f"Plugin {plugin_name} failed validation")
                return None
            
            self.plugins[plugin_name] = plugin
            logger.info(f"Loaded plugin: {plugin_name}")
            return plugin
            
        except Exception as e:
            logger.error(f"Error loading plugin {plugin_name}: {e}")
            return None
    
    def load_all_plugins(self):
        """Load all available plugins."""
        plugin_names = self.discover_plugins()
        for plugin_name in plugin_names:
            self.load_plugin(plugin_name)
    
    async def execute_plugins(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Execute all enabled plugins in priority order."""
        results = {}
        
        # Sort plugins by priority
        sorted_plugins = sorted(
            self.plugins.items(),
            key=lambda x: self.plugin_configs.get(x[0], {}).get("priority", 999)
        )
        
        for plugin_name, plugin in sorted_plugins:
            if not plugin.enabled:
                continue
            
            try:
                logger.info(f"Executing plugin: {plugin_name}")
                result = await plugin.execute(context)
                results[plugin_name] = result
                logger.info(f"Plugin {plugin_name} completed successfully")
                
            except Exception as e:
                logger.error(f"Error executing plugin {plugin_name}: {e}")
                results[plugin_name] = {"error": str(e), "success": False}
        
        return results
    
    def get_plugin_status(self) -> Dict[str, Any]:
        """Get status of all plugins."""
        status = {}
        for plugin_name, plugin in self.plugins.items():
            status[plugin_name] = {
                "enabled": plugin.enabled,
                "description": plugin.get_description(),
                "config": plugin.config,
                "requirements": plugin.get_requirements()
            }
        return status
    
    def enable_plugin(self, plugin_name: str):
        """Enable a plugin."""
        if plugin_name in self.plugins:
            self.plugins[plugin_name].enabled = True
            self.plugin_configs[plugin_name]["enabled"] = True
            self.save_plugin_configs()
    
    def disable_plugin(self, plugin_name: str):
        """Disable a plugin."""
        if plugin_name in self.plugins:
            self.plugins[plugin_name].enabled = False
            self.plugin_configs[plugin_name]["enabled"] = False
            self.save_plugin_configs()

# Example plugins
class CodeAnalyzerPlugin(PluginBase):
    """Analyzes code for potential improvements."""
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze code and return suggestions."""
        # This would analyze the codebase and return improvement suggestions
        return {
            "suggestions": [
                "Consider adding type hints to function parameters",
                "This function could be refactored for better readability",
                "Consider using async/await for I/O operations"
            ],
            "success": True
        }
    
    def get_description(self) -> str:
        return "Analyzes code for potential improvements and refactoring opportunities"
    
    def get_requirements(self) -> List[str]:
        return ["ast", "pylint"]

class TestGeneratorPlugin(PluginBase):
    """Generates tests for existing code."""
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Generate tests for the codebase."""
        # This would generate tests for the codebase
        return {
            "tests_generated": 5,
            "coverage_improvement": "15%",
            "success": True
        }
    
    def get_description(self) -> str:
        return "Generates comprehensive tests for existing code to improve coverage"
    
    def get_requirements(self) -> List[str]:
        return ["pytest", "coverage"]

class DocumentationGeneratorPlugin(PluginBase):
    """Generates documentation for the codebase."""
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Generate documentation."""
        # This would generate documentation
        return {
            "docs_generated": 3,
            "files_updated": 2,
            "success": True
        }
    
    def get_description(self) -> str:
        return "Generates and updates documentation for the codebase"
    
    def get_requirements(self) -> List[str]:
        return ["sphinx", "mkdocs"]

class SecurityScannerPlugin(PluginBase):
    """Scans code for security vulnerabilities."""
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Scan for security issues."""
        # This would scan for security issues
        return {
            "vulnerabilities_found": 2,
            "high_severity": 0,
            "medium_severity": 1,
            "low_severity": 1,
            "success": True
        }
    
    def get_description(self) -> str:
        return "Scans code for security vulnerabilities and provides remediation suggestions"
    
    def get_requirements(self) -> List[str]:
        return ["bandit", "safety"]

class PerformanceOptimizerPlugin(PluginBase):
    """Optimizes code for better performance."""
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Optimize code performance."""
        # This would optimize code performance
        return {
            "optimizations_applied": 3,
            "performance_improvement": "25%",
            "success": True
        }
    
    def get_description(self) -> str:
        return "Analyzes and optimizes code for better performance"
    
    def get_requirements(self) -> List[str]:
        return ["cProfile", "memory_profiler"]

# Initialize plugin manager
plugin_manager = PluginManager()
