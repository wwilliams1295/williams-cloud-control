#!/usr/bin/env python3
"""
Test Auto-Improvement System
Easy to start/stop testing of the auto-improvement system.
"""

import asyncio
import signal
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from scripts.master_auto_improvement import MasterAutoImprovementSystem

class TestAutoImprovement:
    def __init__(self):
        self.system = MasterAutoImprovementSystem()
        self.running = False
        
    def signal_handler(self, signum, frame):
        print("\n🛑 Stopping auto-improvement system...")
        self.running = False
        
    async def run_test(self, interval_minutes=2):
        """Run the auto-improvement system with easy stopping."""
        print("🚀 Starting Auto-Improvement Test")
        print(f"⏱️  Running every {interval_minutes} minutes")
        print("🛑 Press Ctrl+C to stop at any time")
        print("=" * 50)
        
        # Set up signal handler for graceful stopping
        signal.signal(signal.SIGINT, self.signal_handler)
        self.running = True
        
        cycle_count = 0
        
        while self.running:
            try:
                cycle_count += 1
                print(f"\n🔄 CYCLE #{cycle_count} - {asyncio.get_event_loop().time():.0f}")
                print("-" * 30)
                
                # Run one improvement cycle
                result = await self.system.run_master_improvement_cycle()
                
                # Show what happened
                if result.get("success", False):
                    print("✅ Cycle completed successfully!")
                else:
                    print("⚠️  Cycle had issues")
                
                # Show the selected idea
                selected_idea = result.get("improvement_result", {}).get("output", {}).get("selected_idea", {})
                if selected_idea:
                    print(f"💡 Selected Idea: {selected_idea.get('title', 'Unknown')}")
                    print(f"🎯 Impact: {selected_idea.get('impact', 'Unknown')}")
                    print(f"⭐ Wow Factor: {selected_idea.get('wow_factor', 'Unknown')}/10")
                
                # Show creativity level
                creativity_level = result.get("creativity_level", 1.0)
                print(f"🧠 Creativity Level: {creativity_level}")
                
                # Show evolution stage
                evolution_stage = result.get("evolution_stage", "unknown")
                print(f"🌱 Evolution Stage: {evolution_stage}")
                
                if not self.running:
                    break
                    
                print(f"⏳ Waiting {interval_minutes} minutes for next cycle...")
                print("   (Press Ctrl+C to stop)")
                
                # Wait for next cycle (with ability to stop)
                for i in range(interval_minutes * 60):
                    if not self.running:
                        break
                    await asyncio.sleep(1)
                    
            except KeyboardInterrupt:
                print("\n🛑 Stopping...")
                break
            except Exception as e:
                print(f"❌ Error in cycle: {e}")
                await asyncio.sleep(10)  # Wait 10 seconds before retrying
        
        print("\n🏁 Auto-improvement test stopped")
        print(f"📊 Total cycles completed: {cycle_count}")
        
        # Show final status
        status = self.system.get_master_status()
        print(f"🎯 Final creativity level: {status['creativity_level']}")
        print(f"🌱 Final evolution stage: {status['evolution_stage']}")
        print(f"🚀 Total breakthroughs: {status['breakthrough_count']}")

async def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Test Auto-Improvement System")
    parser.add_argument("--interval", type=int, default=2, 
                       help="Interval in minutes between cycles (default: 2)")
    parser.add_argument("--once", action="store_true", 
                       help="Run just one cycle and exit")
    
    args = parser.parse_args()
    
    tester = TestAutoImprovement()
    
    if args.once:
        print("🔄 Running single cycle...")
        result = await tester.system.run_master_improvement_cycle()
        print("✅ Single cycle completed!")
        
        # Show results
        selected_idea = result.get("improvement_result", {}).get("output", {}).get("selected_idea", {})
        if selected_idea:
            print(f"💡 Generated Idea: {selected_idea.get('title', 'Unknown')}")
            print(f"🎯 Impact: {selected_idea.get('impact', 'Unknown')}")
            print(f"⭐ Wow Factor: {selected_idea.get('wow_factor', 'Unknown')}/10")
    else:
        await tester.run_test(args.interval)

if __name__ == "__main__":
    asyncio.run(main())
