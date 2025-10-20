#!/usr/bin/env python3
"""
Dynamic Plugin System
====================

A next-level plugin system that enables:
- Hot-swapping plugins without restart
- Dynamic plugin discovery and loading
- Plugin versioning and rollback
- Real-time plugin performance monitoring
- AI-generated plugin creation
"""

import asyncio
import importlib
import json
import logging
import os
import sys
import time
import threading
import types
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable
import hashlib
import zipfile
import tempfile

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.agent import superchat
from memory.memory_system import get_memory

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/dynamic_plugin_system.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class PluginMetadata:
    """Plugin metadata and versioning information."""
    
    def __init__(self, plugin_id: str, version: str, author: str, description: str):
        self.plugin_id = plugin_id
        self.version = version
        self.author = author
        self.description = description
        self.created_at = datetime.now().isoformat()
        self.last_modified = datetime.now().isoformat()
        self.dependencies = []
        self.compatibility = "1.0.0"
        self.performance_metrics = {
            'load_time': 0.0,
            'execution_time': 0.0,
            'memory_usage': 0.0,
            'success_rate': 0.0
        }
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'plugin_id': self.plugin_id,
            'version': self.version,
            'author': self.author,
            'description': self.description,
            'created_at': self.created_at,
            'last_modified': self.last_modified,
            'dependencies': self.dependencies,
            'compatibility': self.compatibility,
            'performance_metrics': self.performance_metrics
        }

class PluginInstance:
    """Runtime plugin instance."""
    
    def __init__(self, metadata: PluginMetadata, module: types.ModuleType):
        self.metadata = metadata
        self.module = module
        self.is_loaded = False
        self.is_active = False
        self.execution_count = 0
        self.last_execution = None
        self.error_count = 0
    
    def execute(self, *args, **kwargs) -> Any:
        """Execute the plugin."""
        if not self.is_loaded:
            raise RuntimeError(f"Plugin {self.metadata.plugin_id} is not loaded")
        
        start_time = time.time()
        
        try:
            # Execute the plugin's main function
            if hasattr(self.module, 'execute'):
                result = self.module.execute(*args, **kwargs)
            elif hasattr(self.module, 'main'):
                result = self.module.main(*args, **kwargs)
            else:
                raise AttributeError("Plugin must have 'execute' or 'main' function")
            
            # Update metrics
            execution_time = time.time() - start_time
            self.metadata.performance_metrics['execution_time'] = (
                (self.metadata.performance_metrics['execution_time'] * self.execution_count + execution_time) /
                (self.execution_count + 1)
            )
            
            self.execution_count += 1
            self.last_execution = datetime.now().isoformat()
            
            return result
            
        except Exception as e:
            self.error_count += 1
            logger.error(f"Plugin {self.metadata.plugin_id} execution failed: {e}")
            raise
    
    def get_status(self) -> Dict[str, Any]:
        """Get plugin status."""
        return {
            'plugin_id': self.metadata.plugin_id,
            'version': self.metadata.version,
            'is_loaded': self.is_loaded,
            'is_active': self.is_active,
            'execution_count': self.execution_count,
            'last_execution': self.last_execution,
            'error_count': self.error_count,
            'performance_metrics': self.metadata.performance_metrics
        }

class DynamicPluginManager:
    """Dynamic plugin management system."""
    
    def __init__(self):
        self.plugins: Dict[str, PluginInstance] = {}
        self.plugin_directory = Path("src/plugins")
        self.temp_directory = Path("data/cache/plugin_cache")
        self.temp_directory.mkdir(parents=True, exist_ok=True)
        self.plugin_registry = {}
        self.load_plugin_registry()
    
    def load_plugin_registry(self):
        """Load plugin registry from file."""
        registry_file = self.plugin_directory / "plugin_registry.json"
        if registry_file.exists():
            with open(registry_file, 'r') as f:
                self.plugin_registry = json.load(f)
        else:
            self.plugin_registry = {}
    
    def save_plugin_registry(self):
        """Save plugin registry to file."""
        registry_file = self.plugin_directory / "plugin_registry.json"
        with open(registry_file, 'w') as f:
            json.dump(self.plugin_registry, f, indent=2)
    
    def discover_plugins(self) -> List[Path]:
        """Discover available plugins in the plugin directory."""
        plugins = []
        for plugin_file in self.plugin_directory.glob("*.py"):
            if plugin_file.name != "__init__.py":
                plugins.append(plugin_file)
        return plugins
    
    def load_plugin(self, plugin_path: Path) -> Optional[PluginInstance]:
        """Load a plugin from file."""
        try:
            # Generate plugin ID from filename
            plugin_id = plugin_path.stem
            
            # Load module
            spec = importlib.util.spec_from_file_location(plugin_id, plugin_path)
            if spec is None:
                logger.error(f"Could not load spec for {plugin_path}")
                return None
            
            module = importlib.util.module_from_spec(spec)
            sys.modules[plugin_id] = module
            spec.loader.exec_module(module)
            
            # Extract metadata
            metadata = self._extract_metadata(module, plugin_id)
            
            # Create plugin instance
            plugin_instance = PluginInstance(metadata, module)
            plugin_instance.is_loaded = True
            
            # Register plugin
            self.plugins[plugin_id] = plugin_instance
            self.plugin_registry[plugin_id] = metadata.to_dict()
            
            logger.info(f"Loaded plugin: {plugin_id} v{metadata.version}")
            return plugin_instance
            
        except Exception as e:
            logger.error(f"Failed to load plugin {plugin_path}: {e}")
            return None
    
    def _extract_metadata(self, module: types.ModuleType, plugin_id: str) -> PluginMetadata:
        """Extract metadata from plugin module."""
        # Try to get metadata from module attributes
        version = getattr(module, '__version__', '1.0.0')
        author = getattr(module, '__author__', 'Unknown')
        description = getattr(module, '__doc__', f'Plugin {plugin_id}')
        
        # Clean up description
        if description:
            description = description.strip().split('\n')[0]
        
        return PluginMetadata(plugin_id, version, author, description)
    
    def unload_plugin(self, plugin_id: str) -> bool:
        """Unload a plugin."""
        try:
            if plugin_id in self.plugins:
                plugin = self.plugins[plugin_id]
                plugin.is_loaded = False
                plugin.is_active = False
                
                # Remove from sys.modules
                if plugin_id in sys.modules:
                    del sys.modules[plugin_id]
                
                # Remove from registry
                if plugin_id in self.plugin_registry:
                    del self.plugin_registry[plugin_id]
                
                logger.info(f"Unloaded plugin: {plugin_id}")
                return True
            else:
                logger.warning(f"Plugin {plugin_id} not found")
                return False
                
        except Exception as e:
            logger.error(f"Failed to unload plugin {plugin_id}: {e}")
            return False
    
    def reload_plugin(self, plugin_id: str) -> bool:
        """Reload a plugin (hot-swap)."""
        try:
            if plugin_id in self.plugins:
                # Unload first
                self.unload_plugin(plugin_id)
                
                # Find plugin file
                plugin_file = self.plugin_directory / f"{plugin_id}.py"
                if plugin_file.exists():
                    # Load again
                    plugin_instance = self.load_plugin(plugin_file)
                    if plugin_instance:
                        plugin_instance.is_active = True
                        logger.info(f"Reloaded plugin: {plugin_id}")
                        return True
                
            return False
            
        except Exception as e:
            logger.error(f"Failed to reload plugin {plugin_id}: {e}")
            return False
    
    def execute_plugin(self, plugin_id: str, *args, **kwargs) -> Any:
        """Execute a plugin."""
        if plugin_id not in self.plugins:
            raise ValueError(f"Plugin {plugin_id} not found")
        
        plugin = self.plugins[plugin_id]
        if not plugin.is_loaded:
            raise RuntimeError(f"Plugin {plugin_id} is not loaded")
        
        return plugin.execute(*args, **kwargs)
    
    def get_plugin_status(self, plugin_id: str) -> Optional[Dict[str, Any]]:
        """Get plugin status."""
        if plugin_id in self.plugins:
            return self.plugins[plugin_id].get_status()
        return None
    
    def list_plugins(self) -> List[Dict[str, Any]]:
        """List all plugins."""
        return [plugin.get_status() for plugin in self.plugins.values()]
    
    def auto_discover_and_load(self):
        """Automatically discover and load all plugins."""
        discovered_plugins = self.discover_plugins()
        loaded_count = 0
        
        for plugin_path in discovered_plugins:
            plugin_id = plugin_path.stem
            if plugin_id not in self.plugins:
                plugin_instance = self.load_plugin(plugin_path)
                if plugin_instance:
                    plugin_instance.is_active = True
                    loaded_count += 1
        
        logger.info(f"Auto-discovered and loaded {loaded_count} plugins")
        return loaded_count

class AIPluginGenerator:
    """AI-powered plugin generation system."""
    
    def __init__(self, plugin_manager: DynamicPluginManager):
        self.plugin_manager = plugin_manager
        self.memory = get_memory()
    
    async def generate_plugin(self, description: str, plugin_type: str = "utility") -> Optional[str]:
        """Generate a plugin based on description."""
        try:
            # Create plugin generation prompt
            prompt = f"""
            Create a Python plugin for the Jarvis AI Assistant system based on this description:
            
            Description: {description}
            Plugin Type: {plugin_type}
            
            Requirements:
            1. Must have a main 'execute' function that takes *args and **kwargs
            2. Must include proper metadata (__version__, __author__, __doc__)
            3. Must be compatible with the existing plugin system
            4. Must include error handling
            5. Must be well-documented
            
            Generate the complete Python code for this plugin.
            """
            
            # Generate plugin code using AI
            plugin_code = await superchat(prompt, "You are an expert Python developer creating plugins.")
            
            # Extract code from response (remove markdown formatting if present)
            if "```python" in plugin_code:
                plugin_code = plugin_code.split("```python")[1].split("```")[0]
            elif "```" in plugin_code:
                plugin_code = plugin_code.split("```")[1].split("```")[0]
            
            # Clean up the code
            plugin_code = plugin_code.strip()
            
            # Generate plugin filename
            plugin_name = self._generate_plugin_name(description)
            plugin_file = self.plugin_manager.plugin_directory / f"{plugin_name}.py"
            
            # Write plugin file
            with open(plugin_file, 'w') as f:
                f.write(plugin_code)
            
            # Load the new plugin
            plugin_instance = self.plugin_manager.load_plugin(plugin_file)
            if plugin_instance:
                plugin_instance.is_active = True
                logger.info(f"Generated and loaded plugin: {plugin_name}")
                return plugin_name
            
            return None
            
        except Exception as e:
            logger.error(f"Failed to generate plugin: {e}")
            return None
    
    def _generate_plugin_name(self, description: str) -> str:
        """Generate a plugin name from description."""
        # Simple name generation - in production, would use more sophisticated logic
        words = description.lower().split()
        name = "_".join(words[:3])  # Take first 3 words
        name = "".join(c for c in name if c.isalnum() or c == "_")
        return f"ai_generated_{name}"
    
    async def improve_plugin(self, plugin_id: str, improvement_description: str) -> bool:
        """Improve an existing plugin based on description."""
        try:
            if plugin_id not in self.plugin_manager.plugins:
                logger.error(f"Plugin {plugin_id} not found")
                return False
            
            plugin = self.plugin_manager.plugins[plugin_id]
            
            # Read current plugin code
            plugin_file = self.plugin_manager.plugin_directory / f"{plugin_id}.py"
            with open(plugin_file, 'r') as f:
                current_code = f.read()
            
            # Create improvement prompt
            prompt = f"""
            Improve this existing plugin based on the description:
            
            Current Plugin Code:
            {current_code}
            
            Improvement Description: {improvement_description}
            
            Requirements:
            1. Keep the same interface (execute function)
            2. Maintain backward compatibility
            3. Add the requested improvements
            4. Preserve existing functionality
            5. Add proper error handling
            
            Generate the improved Python code.
            """
            
            # Generate improved code
            improved_code = await superchat(prompt, "You are an expert Python developer improving existing code.")
            
            # Extract code from response
            if "```python" in improved_code:
                improved_code = improved_code.split("```python")[1].split("```")[0]
            elif "```" in improved_code:
                improved_code = improved_code.split("```")[1].split("```")[0]
            
            improved_code = improved_code.strip()
            
            # Create backup
            backup_file = plugin_file.with_suffix('.py.backup')
            with open(backup_file, 'w') as f:
                f.write(current_code)
            
            # Write improved code
            with open(plugin_file, 'w') as f:
                f.write(improved_code)
            
            # Reload plugin
            success = self.plugin_manager.reload_plugin(plugin_id)
            
            if success:
                logger.info(f"Improved plugin: {plugin_id}")
                return True
            else:
                # Rollback on failure
                with open(plugin_file, 'w') as f:
                    f.write(current_code)
                logger.error(f"Failed to reload improved plugin {plugin_id}, rolled back")
                return False
                
        except Exception as e:
            logger.error(f"Failed to improve plugin {plugin_id}: {e}")
            return False

class PluginPerformanceMonitor:
    """Monitor plugin performance and health."""
    
    def __init__(self, plugin_manager: DynamicPluginManager):
        self.plugin_manager = plugin_manager
        self.monitoring_active = False
        self.performance_data = {}
    
    def start_monitoring(self):
        """Start performance monitoring."""
        self.monitoring_active = True
        monitoring_thread = threading.Thread(target=self._monitoring_loop, daemon=True)
        monitoring_thread.start()
        logger.info("Started plugin performance monitoring")
    
    def stop_monitoring(self):
        """Stop performance monitoring."""
        self.monitoring_active = False
        logger.info("Stopped plugin performance monitoring")
    
    def _monitoring_loop(self):
        """Main monitoring loop."""
        while self.monitoring_active:
            try:
                for plugin_id, plugin in self.plugin_manager.plugins.items():
                    if plugin.is_loaded and plugin.is_active:
                        # Collect performance data
                        status = plugin.get_status()
                        self.performance_data[plugin_id] = {
                            'timestamp': datetime.now().isoformat(),
                            'execution_count': status['execution_count'],
                            'error_count': status['error_count'],
                            'success_rate': self._calculate_success_rate(plugin),
                            'average_execution_time': status['performance_metrics']['execution_time']
                        }
                
                time.sleep(60)  # Monitor every minute
                
            except Exception as e:
                logger.error(f"Error in monitoring loop: {e}")
                time.sleep(30)
    
    def _calculate_success_rate(self, plugin: PluginInstance) -> float:
        """Calculate plugin success rate."""
        if plugin.execution_count == 0:
            return 1.0
        
        success_count = plugin.execution_count - plugin.error_count
        return success_count / plugin.execution_count
    
    def get_performance_report(self) -> Dict[str, Any]:
        """Get performance report for all plugins."""
        return {
            'timestamp': datetime.now().isoformat(),
            'monitoring_active': self.monitoring_active,
            'plugins': self.performance_data
        }

# Main execution
async def main():
    """Main execution function."""
    # Initialize plugin manager
    plugin_manager = DynamicPluginManager()
    
    # Auto-discover and load plugins
    plugin_manager.auto_discover_and_load()
    
    # Initialize AI plugin generator
    ai_generator = AIPluginGenerator(plugin_manager)
    
    # Initialize performance monitor
    performance_monitor = PluginPerformanceMonitor(plugin_manager)
    performance_monitor.start_monitoring()
    
    # Example: Generate a new plugin
    plugin_name = await ai_generator.generate_plugin(
        "A plugin that provides weather information",
        "utility"
    )
    
    if plugin_name:
        print(f"Generated plugin: {plugin_name}")
        
        # Execute the plugin
        try:
            result = plugin_manager.execute_plugin(plugin_name, "New York")
            print(f"Plugin result: {result}")
        except Exception as e:
            print(f"Plugin execution failed: {e}")
    
    # List all plugins
    plugins = plugin_manager.list_plugins()
    print(f"Loaded plugins: {len(plugins)}")
    for plugin in plugins:
        print(f"- {plugin['plugin_id']} v{plugin['version']}")

if __name__ == "__main__":
    asyncio.run(main())
