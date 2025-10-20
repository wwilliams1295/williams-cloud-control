#!/usr/bin/env python3
"""
Sandbox Testing System
Provides a safe environment for testing code improvements before applying them.
"""

import asyncio
import logging
import os
import shutil
import subprocess  # nosec B404
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional
try:
    import docker
except ImportError:
    docker = None
import json

logger = logging.getLogger(__name__)

class SandboxTester:
    """Manages sandbox testing of code improvements."""
    
    def __init__(self, base_dir: str = None):
        self.base_dir = Path(base_dir) if base_dir else Path(__file__).parent
        self.sandbox_dir = self.base_dir / "sandbox"
        self.test_results_dir = self.base_dir / "test_results"
        self.docker_client = None
        self._init_docker()
    
    def _init_docker(self):
        """Initialize Docker client if available."""
        if docker is None:
            logger.warning("Docker module not available")
            self.docker_client = None
            return
            
        try:
            self.docker_client = docker.from_env()
            logger.info("Docker client initialized")
        except Exception as e:
            logger.warning(f"Docker not available: {e}")
            self.docker_client = None
    
    async def create_sandbox(self, test_name: str) -> Path:
        """Create a sandbox environment for testing."""
        sandbox_path = self.sandbox_dir / test_name
        sandbox_path.mkdir(parents=True, exist_ok=True)
        
        # Copy project files to sandbox
        await self._copy_project_files(sandbox_path)
        
        # Create sandbox-specific files
        await self._create_sandbox_config(sandbox_path)
        
        logger.info(f"Created sandbox: {sandbox_path}")
        return sandbox_path
    
    async def _copy_project_files(self, sandbox_path: Path):
        """Copy project files to sandbox, excluding sensitive files."""
        exclude_dirs = {'.git', '__pycache__', '.pytest_cache', 'venv', '.venv', 'backups', 'sandbox'}
        exclude_files = {'.env', 'client_secret.json', 'token.json', '*.log'}
        
        for item in self.base_dir.iterdir():
            if item.name in exclude_dirs:
                continue
            if any(item.name.endswith(ext) for ext in ['.log', '.pyc', '.pyo']):
                continue
            
            if item.is_file():
                shutil.copy2(item, sandbox_path / item.name)
            elif item.is_dir():
                shutil.copytree(item, sandbox_path / item.name, ignore=shutil.ignore_patterns(*exclude_files))
    
    async def _create_sandbox_config(self, sandbox_path: Path):
        """Create sandbox-specific configuration."""
        # Create test environment file
        env_content = """
# Sandbox test environment
DEBUG_LOG=1
NODE_NAME=sandbox-test
APP_TIMEZONE=UTC
SANDBOX_MODE=true
IMPROVE_INTERVAL_MIN=1
MAX_IMPROVEMENTS_PER_DAY=100
"""
        with open(sandbox_path / ".env.test", 'w') as f:
            f.write(env_content)
        
        # Create test requirements
        test_requirements = """
# Test requirements for sandbox
pytest>=7.0.0
pytest-asyncio>=0.21.0
pytest-cov>=4.0.0
bandit>=1.7.0
safety>=2.0.0
mypy>=1.0.0
black>=23.0.0
isort>=5.12.0
flake8>=6.0.0
"""
        with open(sandbox_path / "requirements.test.txt", 'w') as f:
            f.write(test_requirements)
    
    async def run_tests(self, sandbox_path: Path, test_type: str = "all") -> Dict[str, Any]:
        """Run tests in the sandbox environment."""
        results = {
            "test_type": test_type,
            "timestamp": datetime.now().isoformat(),
            "sandbox_path": str(sandbox_path),
            "tests": {}
        }
        
        try:
            # Run unit tests
            if test_type in ["all", "unit"]:
                unit_results = await self._run_unit_tests(sandbox_path)
                results["tests"]["unit"] = unit_results
            
            # Run integration tests
            if test_type in ["all", "integration"]:
                integration_results = await self._run_integration_tests(sandbox_path)
                results["tests"]["integration"] = integration_results
            
            # Run security tests
            if test_type in ["all", "security"]:
                security_results = await self._run_security_tests(sandbox_path)
                results["tests"]["security"] = security_results
            
            # Run performance tests
            if test_type in ["all", "performance"]:
                performance_results = await self._run_performance_tests(sandbox_path)
                results["tests"]["performance"] = performance_results
            
            # Run linting
            if test_type in ["all", "lint"]:
                lint_results = await self._run_linting(sandbox_path)
                results["tests"]["lint"] = lint_results
            
            results["success"] = all(
                test.get("success", False) 
                for test in results["tests"].values()
            )
            
        except Exception as e:
            logger.error(f"Error running tests: {e}")
            results["error"] = str(e)
            results["success"] = False
        
        return results
    
    async def _run_unit_tests(self, sandbox_path: Path) -> Dict[str, Any]:
        """Run unit tests."""
        try:
            result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "pytest", "tests/", "-v", "--tb=short"],
                cwd=sandbox_path,
                capture_output=True,
                text=True,
                timeout=300
            )
            
            return {
                "success": result.returncode == 0,
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr
            }
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "Test timeout"}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _run_integration_tests(self, sandbox_path: Path) -> Dict[str, Any]:
        """Run integration tests."""
        try:
            # Test the main application
            result = subprocess.run(  # nosec B603
                [sys.executable, "-c", "import sys; sys.path.insert(0, '.'); from cloud import app; print('App loads successfully')"],
                cwd=sandbox_path,
                capture_output=True,
                text=True,
                timeout=60
            )
            
            return {
                "success": result.returncode == 0,
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr
            }
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "Integration test timeout"}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _run_security_tests(self, sandbox_path: Path) -> Dict[str, Any]:
        """Run security tests."""
        try:
            # Run bandit security scanner
            result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "bandit", "-r", ".", "-f", "json"],
                cwd=sandbox_path,
                capture_output=True,
                text=True,
                timeout=120
            )
            
            # Parse bandit results
            bandit_results = {}
            if result.stdout:
                try:
                    bandit_results = json.loads(result.stdout)
                except json.JSONDecodeError:
                    bandit_results = {"raw_output": result.stdout}
            
            return {
                "success": result.returncode == 0,
                "returncode": result.returncode,
                "bandit_results": bandit_results,
                "stderr": result.stderr
            }
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "Security test timeout"}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _run_performance_tests(self, sandbox_path: Path) -> Dict[str, Any]:
        """Run performance tests."""
        try:
            # Simple performance test - time to import and initialize
            start_time = time.time()
            
            result = subprocess.run(  # nosec B603
                [sys.executable, "-c", "import sys; sys.path.insert(0, '.'); from agent_v2 import superchat; print('Import successful')"],
                cwd=sandbox_path,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            end_time = time.time()
            import_time = end_time - start_time
            
            return {
                "success": result.returncode == 0,
                "import_time": import_time,
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr
            }
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "Performance test timeout"}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def _run_linting(self, sandbox_path: Path) -> Dict[str, Any]:
        """Run linting tests."""
        try:
            # Run flake8
            flake8_result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "flake8", ".", "--count", "--select=E9,F63,F7,F82", "--show-source", "--statistics"],
                cwd=sandbox_path,
                capture_output=True,
                text=True,
                timeout=120
            )
            
            # Run mypy
            mypy_result = subprocess.run(  # nosec B603
                [sys.executable, "-m", "mypy", ".", "--ignore-missing-imports"],
                cwd=sandbox_path,
                capture_output=True,
                text=True,
                timeout=120
            )
            
            return {
                "success": flake8_result.returncode == 0 and mypy_result.returncode == 0,
                "flake8": {
                    "returncode": flake8_result.returncode,
                    "stdout": flake8_result.stdout,
                    "stderr": flake8_result.stderr
                },
                "mypy": {
                    "returncode": mypy_result.returncode,
                    "stdout": mypy_result.stdout,
                    "stderr": mypy_result.stderr
                }
            }
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "Linting test timeout"}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    async def test_improvement(self, improvement_patch: str, test_name: str = None) -> Dict[str, Any]:
        """Test a specific improvement in the sandbox."""
        if not test_name:
            test_name = f"improvement_{int(time.time())}"
        
        sandbox_path = await self.create_sandbox(test_name)
        
        try:
            # Apply the improvement patch
            patch_file = sandbox_path / "improvement.patch"
            with open(patch_file, 'w') as f:
                f.write(improvement_patch)
            
            # Apply patch
            apply_result = subprocess.run(  # nosec B603
                ["git", "apply", str(patch_file)],
                cwd=sandbox_path,
                capture_output=True,
                text=True
            )
            
            if apply_result.returncode != 0:
                return {
                    "success": False,
                    "error": "Failed to apply patch",
                    "patch_error": apply_result.stderr
                }
            
            # Run tests
            test_results = await self.run_tests(sandbox_path)
            
            # Clean up
            shutil.rmtree(sandbox_path, ignore_errors=True)
            
            return test_results
            
        except Exception as e:
            logger.error(f"Error testing improvement: {e}")
            return {"success": False, "error": str(e)}
    
    async def cleanup_sandbox(self, test_name: str):
        """Clean up a specific sandbox."""
        sandbox_path = self.sandbox_dir / test_name
        if sandbox_path.exists():
            shutil.rmtree(sandbox_path, ignore_errors=True)
            logger.info(f"Cleaned up sandbox: {test_name}")

async def main():
    """Main entry point for testing."""
    tester = SandboxTester()
    
    if len(sys.argv) > 1:
        test_name = sys.argv[1]
        sandbox_path = await tester.create_sandbox(test_name)
        results = await tester.run_tests(sandbox_path)
        print(json.dumps(results, indent=2))
    else:
        print("Usage: python sandbox_tester.py <test_name>")

if __name__ == "__main__":
    asyncio.run(main())
