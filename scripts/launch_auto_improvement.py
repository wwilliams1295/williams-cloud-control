#!/usr/bin/env python3
"""
Auto-Improvement Launcher
Easy launcher for the auto-improvement system with different modes.
"""

import asyncio
import sys
import argparse
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

async def main():
    parser = argparse.ArgumentParser(description="Launch Auto-Improvement System")
    parser.add_argument("--mode", choices=["basic", "advanced", "creative", "master"], 
                       default="master", help="Improvement mode to use")
    parser.add_argument("--interval", type=int, default=20, 
                       help="Interval in minutes between improvements")
    parser.add_argument("--once", action="store_true", 
                       help="Run once instead of continuously")
    parser.add_argument("--status", action="store_true", 
                       help="Show status and exit")
    
    args = parser.parse_args()
    
    if args.mode == "basic":
        from scripts.auto_improvement_loop import main as basic_main
        if args.status:
            sys.argv = ["auto_improvement_loop.py", "--status"]
        elif args.once:
            sys.argv = ["auto_improvement_loop.py"]
        else:
            sys.argv = ["auto_improvement_loop.py"]
        await basic_main()
    
    elif args.mode == "advanced":
        from scripts.advanced_auto_improvement import main as advanced_main
        if args.status:
            sys.argv = ["advanced_auto_improvement.py", "--status"]
        elif args.once:
            sys.argv = ["advanced_auto_improvement.py", "--once"]
        else:
            sys.argv = ["advanced_auto_improvement.py", f"--interval={args.interval}"]
        await advanced_main()
    
    elif args.mode == "creative":
        from scripts.creative_ai_evolution import main as creative_main
        if args.status:
            sys.argv = ["creative_ai_evolution.py", "--status"]
        elif args.once:
            sys.argv = ["creative_ai_evolution.py", "--once"]
        else:
            sys.argv = ["creative_ai_evolution.py", f"--interval={args.interval}"]
        await creative_main()
    
    elif args.mode == "master":
        from scripts.master_auto_improvement import main as master_main
        if args.status:
            sys.argv = ["master_auto_improvement.py", "--status"]
        elif args.once:
            sys.argv = ["master_auto_improvement.py", "--once"]
        else:
            sys.argv = ["master_auto_improvement.py", f"--interval={args.interval}"]
        await master_main()

if __name__ == "__main__":
    asyncio.run(main())
