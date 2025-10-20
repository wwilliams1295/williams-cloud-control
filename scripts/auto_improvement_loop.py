#!/usr/bin/env python3
"""
Auto-Improvement Loop System
Continuously improves the codebase by analyzing, generating improvements, and testing them.
"""

import asyncio
import logging
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional
import subprocess  # nosec B404
import json

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from config.settings_simple import get_settings
# Import tools dynamically to avoid import errors
# from tools.auto_improver import main as auto_improve
# from tools.self_loop import main as self_loop

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('auto_improvement.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class AutoImprovementLoop:
    """Manages the continuous improvement of the codebase."""
    
    def __init__(self):
        self.settings = get_settings()
        self.improvement_interval = int(os.getenv('IMPROVE_INTERVAL_MIN', '60')) * 60  # Convert to seconds
        self.last_improvement = None
        self.improvement_history = []
        self.max_improvements_per_day = 10
        self.sandbox_mode = os.getenv('SANDBOX_MODE', 'false').lower() == 'true'
        
    async def run_improvement_cycle(self) -> Dict[str, Any]:
        """Run a single improvement cycle."""
        logger.info("Starting improvement cycle...")
        
        try:
            # Check if we should run improvements
            if not self._should_run_improvement():
                logger.info("Skipping improvement cycle - cooldown period active")
                return {"status": "skipped", "reason": "cooldown"}
            
            # Try to run advanced improvement first
            logger.info("Attempting advanced improvement...")
            advanced_result = await self._run_advanced_improvement()
            
            if advanced_result.get("success", False):
                logger.info("Advanced improvement successful")
                result = {
                    "status": "completed",
                    "timestamp": datetime.now().isoformat(),
                    "improvement_type": "advanced",
                    "result": advanced_result,
                    "total_improvements": len(self.improvement_history)
                }
            else:
                # Fallback to basic improvement
                logger.info("Falling back to basic improvement...")
                improvement_result = await self._run_auto_improver()
                self_loop_result = await self._run_self_loop()
                
                result = {
                    "status": "completed",
                    "timestamp": datetime.now().isoformat(),
                    "improvement_type": "basic",
                    "improvement_result": improvement_result,
                    "self_loop_result": self_loop_result,
                    "total_improvements": len(self.improvement_history)
                }
            
            self.improvement_history.append(result)
            self.last_improvement = datetime.now()
            
            logger.info(f"Improvement cycle completed: {result}")
            return result
            
        except Exception as e:
            logger.error(f"Error in improvement cycle: {e}")
            return {"status": "error", "error": str(e)}
    
    def _should_run_improvement(self) -> bool:
        """Check if we should run an improvement cycle."""
        if self.last_improvement is None:
            return True
        
        time_since_last = datetime.now() - self.last_improvement
        if time_since_last.total_seconds() < self.improvement_interval:
            return False
        
        # Check daily limit
        today = datetime.now().date()
        today_improvements = [
            imp for imp in self.improvement_history
            if datetime.fromisoformat(imp['timestamp']).date() == today
        ]
        
        return len(today_improvements) < self.max_improvements_per_day
    
    async def _run_auto_improver(self) -> Dict[str, Any]:
        """Run the auto-improver tool."""
        try:
            # Run auto-improver in a subprocess
            result = subprocess.run(  # nosec B603
                [sys.executable, "tools/auto_improver.py", "--auto"],
                capture_output=True,
                text=True,
                cwd=Path(__file__).parent
            )
            
            return {
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "success": result.returncode == 0
            }
        except Exception as e:
            logger.error(f"Error running auto-improver: {e}")
            return {"success": False, "error": str(e)}
    
    async def _run_self_loop(self) -> Dict[str, Any]:
        """Run the self-loop tool."""
        try:
            # Run self-loop in a subprocess
            result = subprocess.run(  # nosec B603
                [sys.executable, "tools/self_loop.py"],
                capture_output=True,
                text=True,
                cwd=Path(__file__).parent
            )
            
            return {
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "success": result.returncode == 0
            }
        except Exception as e:
            logger.error(f"Error running self-loop: {e}")
            return {"success": False, "error": str(e)}
    
    async def _run_advanced_improvement(self) -> Dict[str, Any]:
        """Run advanced improvement system."""
        try:
            # Try to run the advanced improvement system
            result = subprocess.run(  # nosec B603
                [sys.executable, "scripts/advanced_auto_improvement.py", "--once"],
                capture_output=True,
                text=True,
                cwd=Path(__file__).parent.parent
            )
            
            return {
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "success": result.returncode == 0
            }
        except Exception as e:
            logger.error(f"Error running advanced improvement: {e}")
            return {"success": False, "error": str(e)}
    
    async def run_continuous_loop(self):
        """Run the continuous improvement loop."""
        logger.info("Starting continuous improvement loop...")
        logger.info(f"Improvement interval: {self.improvement_interval} seconds")
        logger.info(f"Sandbox mode: {self.sandbox_mode}")
        
        while True:
            try:
                result = await self.run_improvement_cycle()
                logger.info(f"Cycle result: {result['status']}")
                
                # Wait for next cycle
                await asyncio.sleep(self.improvement_interval)
                
            except KeyboardInterrupt:
                logger.info("Stopping improvement loop...")
                break
            except Exception as e:
                logger.error(f"Unexpected error in improvement loop: {e}")
                await asyncio.sleep(60)  # Wait 1 minute before retrying
    
    def get_status(self) -> Dict[str, Any]:
        """Get current status of the improvement system."""
        return {
            "last_improvement": self.last_improvement.isoformat() if self.last_improvement else None,
            "improvement_interval": self.improvement_interval,
            "total_improvements": len(self.improvement_history),
            "sandbox_mode": self.sandbox_mode,
            "next_improvement_in": max(0, self.improvement_interval - (
                (datetime.now() - self.last_improvement).total_seconds()
                if self.last_improvement else 0
            ))
        }

async def main():
    """Main entry point."""
    loop = AutoImprovementLoop()
    
    if len(sys.argv) > 1 and sys.argv[1] == "--status":
        # Just show status
        status = loop.get_status()
        print(json.dumps(status, indent=2))
        return
    
    # Run continuous loop
    await loop.run_continuous_loop()

if __name__ == "__main__":
    asyncio.run(main())
