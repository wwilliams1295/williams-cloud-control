#!/usr/bin/env python3
"""
Jarvis AI Assistant - Main Entry Point
=====================================

This is the main entry point for the Jarvis AI Assistant application.
It provides a clean interface to the reorganized codebase.

Usage:
    python main.py                    # Start the API server
    python main.py --test            # Run tests
    python main.py --memory-stats    # Show memory statistics
    python main.py --help            # Show help
"""

import sys
import os
import argparse
from pathlib import Path

# Add src to Python path
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

def main():
    parser = argparse.ArgumentParser(description="Jarvis AI Assistant")
    parser.add_argument("--test", action="store_true", help="Run tests")
    parser.add_argument("--memory-stats", action="store_true", help="Show memory statistics")
    parser.add_argument("--start-api", action="store_true", help="Start the API server")
    
    args = parser.parse_args()
    
    if args.test:
        print("Running tests...")
        # Import and run tests
        try:
            from tests import run_tests
            run_tests()
        except ImportError:
            print("Tests not available")
    
    elif args.memory_stats:
        print("Memory Statistics:")
        try:
            from memory.memory_system import get_memory
            memory = get_memory()
            stats = memory.get_system_stats()
            print(f"Total conversations: {stats.get('total_conversations', 0)}")
            print(f"Unique users: {stats.get('unique_users', 0)}")
            print(f"Cache size: {stats.get('cache_size', 0)}")
        except ImportError as e:
            print(f"Memory system not available: {e}")
    
    elif args.start_api:
        print("Starting API server...")
        try:
            from api.cloud import app
            import uvicorn
            uvicorn.run(app, host="0.0.0.0", port=8000)
        except ImportError as e:
            print(f"API server not available: {e}")
    
    else:
        # Default: start API server
        print("Starting Jarvis AI Assistant...")
        try:
            from api.cloud import app
            import uvicorn
            uvicorn.run(app, host="0.0.0.0", port=8000)
        except ImportError as e:
            print(f"Error starting application: {e}")
            print("Make sure all dependencies are installed: pip install -r requirements.txt")

if __name__ == "__main__":
    main()
