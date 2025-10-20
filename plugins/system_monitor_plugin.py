"""
System Monitor Plugin
Monitors system performance and health.
"""

import logging
import psutil
import time
from datetime import datetime
from typing import Dict, Any, List
from plugins.plugin_manager import PluginBase

logger = logging.getLogger(__name__)

class SystemMonitorPlugin(PluginBase):
    """Plugin for monitoring system performance."""
    
    def __init__(self, config: Dict[str, Any] = None):
        super().__init__(config)
        self.name = "system_monitor"
        self.description = "Monitors system performance and health"
        self.thresholds = {
            "cpu_percent": 80.0,
            "memory_percent": 85.0,
            "disk_percent": 90.0
        }
    
    async def execute(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """Execute system monitoring."""
        try:
            # Collect system metrics
            metrics = self._collect_metrics()
            
            # Check for alerts
            alerts = self._check_alerts(metrics)
            
            # Generate report
            report = self._generate_report(metrics, alerts)
            
            return {
                "success": True,
                "message": "System monitoring completed",
                "details": {
                    "metrics": metrics,
                    "alerts": alerts,
                    "report": report,
                    "timestamp": datetime.now().isoformat()
                }
            }
            
        except Exception as e:
            logger.error(f"Error in system monitor plugin: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def _collect_metrics(self) -> Dict[str, Any]:
        """Collect system performance metrics."""
        try:
            # CPU metrics
            cpu_percent = psutil.cpu_percent(interval=1)
            cpu_count = psutil.cpu_count()
            cpu_freq = psutil.cpu_freq()
            
            # Memory metrics
            memory = psutil.virtual_memory()
            swap = psutil.swap_memory()
            
            # Disk metrics
            disk = psutil.disk_usage('/')
            
            # Network metrics
            network = psutil.net_io_counters()
            
            # Process metrics
            processes = []
            for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'memory_percent']):
                try:
                    processes.append(proc.info)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            # Sort by CPU usage
            processes.sort(key=lambda x: x.get('cpu_percent', 0), reverse=True)
            top_processes = processes[:5]
            
            return {
                "cpu": {
                    "percent": cpu_percent,
                    "count": cpu_count,
                    "frequency": cpu_freq.current if cpu_freq else None
                },
                "memory": {
                    "total": memory.total,
                    "available": memory.available,
                    "percent": memory.percent,
                    "used": memory.used,
                    "free": memory.free
                },
                "swap": {
                    "total": swap.total,
                    "used": swap.used,
                    "free": swap.free,
                    "percent": swap.percent
                },
                "disk": {
                    "total": disk.total,
                    "used": disk.used,
                    "free": disk.free,
                    "percent": (disk.used / disk.total) * 100
                },
                "network": {
                    "bytes_sent": network.bytes_sent,
                    "bytes_recv": network.bytes_recv,
                    "packets_sent": network.packets_sent,
                    "packets_recv": network.packets_recv
                },
                "top_processes": top_processes
            }
            
        except Exception as e:
            logger.error(f"Error collecting metrics: {e}")
            return {}
    
    def _check_alerts(self, metrics: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Check for system alerts based on thresholds."""
        alerts = []
        
        try:
            # CPU alert
            if metrics.get("cpu", {}).get("percent", 0) > self.thresholds["cpu_percent"]:
                alerts.append({
                    "type": "cpu_high",
                    "severity": "warning",
                    "message": f"CPU usage is {metrics['cpu']['percent']:.1f}% (threshold: {self.thresholds['cpu_percent']}%)"
                })
            
            # Memory alert
            if metrics.get("memory", {}).get("percent", 0) > self.thresholds["memory_percent"]:
                alerts.append({
                    "type": "memory_high",
                    "severity": "warning",
                    "message": f"Memory usage is {metrics['memory']['percent']:.1f}% (threshold: {self.thresholds['memory_percent']}%)"
                })
            
            # Disk alert
            if metrics.get("disk", {}).get("percent", 0) > self.thresholds["disk_percent"]:
                alerts.append({
                    "type": "disk_high",
                    "severity": "critical",
                    "message": f"Disk usage is {metrics['disk']['percent']:.1f}% (threshold: {self.thresholds['disk_percent']}%)"
                })
            
        except Exception as e:
            logger.error(f"Error checking alerts: {e}")
        
        return alerts
    
    def _generate_report(self, metrics: Dict[str, Any], alerts: List[Dict[str, Any]]) -> str:
        """Generate a human-readable system report."""
        try:
            report = f"""📊 **System Performance Report**
            
🖥️ **CPU**: {metrics.get('cpu', {}).get('percent', 0):.1f}% usage
🧠 **Memory**: {metrics.get('memory', {}).get('percent', 0):.1f}% used ({metrics.get('memory', {}).get('used', 0) // (1024**3):.1f}GB / {metrics.get('memory', {}).get('total', 0) // (1024**3):.1f}GB)
💾 **Disk**: {metrics.get('disk', {}).get('percent', 0):.1f}% used ({metrics.get('disk', {}).get('used', 0) // (1024**3):.1f}GB / {metrics.get('disk', {}).get('total', 0) // (1024**3):.1f}GB)
🌐 **Network**: {metrics.get('network', {}).get('bytes_sent', 0) // (1024**2):.1f}MB sent, {metrics.get('network', {}).get('bytes_recv', 0) // (1024**2):.1f}MB received

"""
            
            if alerts:
                report += "⚠️ **Alerts:**\n"
                for alert in alerts:
                    severity_icon = "🔴" if alert["severity"] == "critical" else "🟡"
                    report += f"  {severity_icon} {alert['message']}\n"
            else:
                report += "✅ **No alerts** - System running normally\n"
            
            # Top processes
            top_processes = metrics.get("top_processes", [])
            if top_processes:
                report += "\n🔥 **Top Processes:**\n"
                for proc in top_processes[:3]:
                    name = proc.get('name', 'Unknown')
                    cpu = proc.get('cpu_percent', 0)
                    memory = proc.get('memory_percent', 0)
                    report += f"  • {name}: {cpu:.1f}% CPU, {memory:.1f}% Memory\n"
            
            return report
            
        except Exception as e:
            logger.error(f"Error generating report: {e}")
            return f"Error generating report: {e}"
    
    def get_description(self) -> str:
        return "Monitors system performance and generates health reports"
    
    def get_requirements(self) -> List[str]:
        return ["psutil"]
    
    def validate_config(self) -> bool:
        """Validate plugin configuration."""
        try:
            import psutil
            return True
        except ImportError:
            return False
