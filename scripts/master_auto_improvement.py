#!/usr/bin/env python3
"""
Master Auto-Improvement System
Orchestrates multiple improvement systems to create a self-evolving AI that continuously
improves beyond what was originally imagined.
"""

import asyncio
import logging
import os
import sys
import time
import json
import random
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional
import subprocess

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Config system removed - using environment variables directly
from agent import superchat

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('master_auto_improvement.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class MasterAutoImprovementSystem:
    """Master system that orchestrates all improvement mechanisms."""
    
    def __init__(self):
        self.improvement_systems = {
            "basic": "scripts/auto_improvement_loop.py",
            "advanced": "scripts/advanced_auto_improvement.py", 
            "creative": "scripts/creative_ai_evolution.py"
        }
        self.current_system = "basic"
        self.improvement_history = []
        self.evolution_stage = "exploration"  # exploration, growth, mastery, transcendence
        self.breakthrough_count = 0
        self.creativity_level = 1.0
        
    async def run_master_improvement_cycle(self) -> Dict[str, Any]:
        """Run a master improvement cycle that combines all systems."""
        logger.info("🚀 Starting MASTER improvement cycle...")
        
        try:
            # 1. Analyze current state
            current_state = await self._analyze_current_state()
            
            # 2. Determine which system to use
            system_choice = await self._choose_improvement_system(current_state)
            
            # 3. Run the chosen system
            improvement_result = await self._run_improvement_system(system_choice)
            
            # 4. Evaluate results
            evaluation = await self._evaluate_improvement(improvement_result, current_state)
            
            # 5. Evolve the system if needed
            await self._evolve_system(evaluation)
            
            # 6. Record everything
            master_record = {
                "timestamp": datetime.now().isoformat(),
                "system_used": system_choice,
                "evolution_stage": self.evolution_stage,
                "creativity_level": self.creativity_level,
                "current_state": current_state,
                "improvement_result": improvement_result,
                "evaluation": evaluation,
                "breakthrough_count": self.breakthrough_count
            }
            
            self.improvement_history.append(master_record)
            
            # 7. Check for stage evolution
            await self._check_stage_evolution()
            
            logger.info(f"✅ Master improvement cycle completed using {system_choice}")
            return master_record
            
        except Exception as e:
            logger.error(f"❌ Error in master improvement cycle: {e}")
            return {"success": False, "error": str(e)}
    
    async def _analyze_current_state(self) -> Dict[str, Any]:
        """Analyze the current state of the system."""
        
        prompt = f"""
        Analyze the current state of this AI system and provide insights for improvement.
        
        System context:
        - Modular AI assistant with multiple LLM providers
        - Features: file management, cloud storage, plugins, auto-improvement
        - Architecture: providers, routing, core modules, plugins
        - Recent improvements: {len(self.improvement_history)} cycles completed
        
        Provide analysis in JSON format:
        {{
            "maturity_level": "beginner/intermediate/advanced/expert",
            "complexity_score": 1-10,
            "innovation_potential": 1-10,
            "technical_debt": 1-10,
            "creativity_opportunities": 1-10,
            "strengths": [...],
            "weaknesses": [...],
            "improvement_priorities": [...],
            "next_focus_areas": [...]
        }}
        """
        
        try:
            response = await superchat(prompt)
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                return json.loads(json_match.group(0))
            else:
                return self._generate_fallback_analysis()
        except Exception as e:
            logger.error(f"Error analyzing current state: {e}")
            return self._generate_fallback_analysis()
    
    def _generate_fallback_analysis(self) -> Dict[str, Any]:
        """Generate fallback analysis if AI fails."""
        return {
            "maturity_level": "intermediate",
            "complexity_score": 6,
            "innovation_potential": 8,
            "technical_debt": 4,
            "creativity_opportunities": 9,
            "strengths": ["Modular design", "Multiple providers", "Cloud integration"],
            "weaknesses": ["Limited creativity", "Basic automation"],
            "improvement_priorities": ["Enhanced creativity", "Advanced automation"],
            "next_focus_areas": ["AI evolution", "Boundary pushing"]
        }
    
    async def _choose_improvement_system(self, current_state: Dict[str, Any]) -> str:
        """Choose which improvement system to use based on current state."""
        
        maturity = current_state.get("maturity_level", "intermediate")
        complexity = current_state.get("complexity_score", 5)
        innovation_potential = current_state.get("innovation_potential", 5)
        
        # Decision logic based on system state
        if maturity == "beginner" or complexity < 4:
            return "basic"
        elif maturity == "intermediate" and innovation_potential < 7:
            return "advanced"
        elif maturity in ["advanced", "expert"] or innovation_potential >= 7:
            return "creative"
        else:
            # Random choice with weighted probabilities
            weights = {
                "basic": 0.2,
                "advanced": 0.4,
                "creative": 0.4
            }
            return random.choices(list(weights.keys()), weights=list(weights.values()))[0]
    
    async def _run_improvement_system(self, system_name: str) -> Dict[str, Any]:
        """Run the specified improvement system."""
        logger.info(f"Running {system_name} improvement system...")
        
        try:
            script_path = self.improvement_systems[system_name]
            result = subprocess.run([
                sys.executable, script_path, "--once"
            ], capture_output=True, text=True, cwd=Path(__file__).parent.parent)
            
            # Parse JSON output
            try:
                output = json.loads(result.stdout)
                return {
                    "system": system_name,
                    "success": result.returncode == 0,
                    "output": output,
                    "stderr": result.stderr
                }
            except json.JSONDecodeError:
                return {
                    "system": system_name,
                    "success": result.returncode == 0,
                    "output": result.stdout,
                    "stderr": result.stderr
                }
        except Exception as e:
            logger.error(f"Error running {system_name} system: {e}")
            return {
                "system": system_name,
                "success": False,
                "error": str(e)
            }
    
    async def _evaluate_improvement(self, improvement_result: Dict[str, Any], current_state: Dict[str, Any]) -> Dict[str, Any]:
        """Evaluate the improvement results."""
        
        prompt = f"""
        Evaluate this improvement result and provide insights:
        
        Improvement Result:
        {json.dumps(improvement_result, indent=2)}
        
        Previous State:
        {json.dumps(current_state, indent=2)}
        
        Provide evaluation in JSON format:
        {{
            "success_rating": 1-10,
            "creativity_rating": 1-10,
            "impact_rating": 1-10,
            "innovation_level": "incremental/improvement/breakthrough/revolutionary",
            "key_achievements": [...],
            "areas_for_improvement": [...],
            "next_recommendations": [...],
            "overall_assessment": "..."
        }}
        """
        
        try:
            response = await superchat(prompt)
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                return json.loads(json_match.group(0))
            else:
                return self._generate_fallback_evaluation(improvement_result)
        except Exception as e:
            logger.error(f"Error evaluating improvement: {e}")
            return self._generate_fallback_evaluation(improvement_result)
    
    def _generate_fallback_evaluation(self, improvement_result: Dict[str, Any]) -> Dict[str, Any]:
        """Generate fallback evaluation if AI fails."""
        return {
            "success_rating": 7 if improvement_result.get("success", False) else 3,
            "creativity_rating": 6,
            "impact_rating": 5,
            "innovation_level": "improvement",
            "key_achievements": ["System ran successfully"],
            "areas_for_improvement": ["Increase creativity", "Enhance impact"],
            "next_recommendations": ["Try creative system", "Push boundaries"],
            "overall_assessment": "Moderate improvement achieved"
        }
    
    async def _evolve_system(self, evaluation: Dict[str, Any]):
        """Evolve the system based on evaluation results."""
        
        creativity_rating = evaluation.get("creativity_rating", 5)
        impact_rating = evaluation.get("impact_rating", 5)
        innovation_level = evaluation.get("innovation_level", "incremental")
        
        # Boost creativity if ratings are high
        if creativity_rating >= 8 or impact_rating >= 8:
            self.creativity_level = min(10.0, self.creativity_level + 0.5)
            logger.info(f"🚀 Creativity level boosted to {self.creativity_level}")
        
        # Check for breakthrough
        if innovation_level in ["breakthrough", "revolutionary"]:
            self.breakthrough_count += 1
            logger.info(f"🎉 BREAKTHROUGH #{self.breakthrough_count} detected!")
    
    async def _check_stage_evolution(self):
        """Check if the system should evolve to the next stage."""
        
        total_cycles = len(self.improvement_history)
        breakthrough_rate = self.breakthrough_count / max(total_cycles, 1)
        avg_creativity = sum([r.get("creativity_level", 1) for r in self.improvement_history[-10:]]) / min(10, len(self.improvement_history))
        
        # Stage evolution logic
        if self.evolution_stage == "exploration" and total_cycles >= 10 and breakthrough_rate >= 0.1:
            self.evolution_stage = "growth"
            logger.info("🌱 Evolving to GROWTH stage!")
        elif self.evolution_stage == "growth" and total_cycles >= 25 and breakthrough_rate >= 0.2:
            self.evolution_stage = "mastery"
            logger.info("🎯 Evolving to MASTERY stage!")
        elif self.evolution_stage == "mastery" and total_cycles >= 50 and breakthrough_rate >= 0.3:
            self.evolution_stage = "transcendence"
            logger.info("🌟 Evolving to TRANSCENDENCE stage!")
    
    async def run_continuous_master_improvement(self, interval_minutes: int = 20):
        """Run continuous master improvement."""
        logger.info("🚀 Starting CONTINUOUS MASTER improvement...")
        logger.info(f"Improvement interval: {interval_minutes} minutes")
        logger.info(f"Current evolution stage: {self.evolution_stage}")
        logger.info(f"Creativity level: {self.creativity_level}")
        
        while True:
            try:
                result = await self.run_master_improvement_cycle()
                
                # Log significant events
                if result.get("evaluation", {}).get("innovation_level") in ["breakthrough", "revolutionary"]:
                    logger.info("🎉 BREAKTHROUGH ACHIEVED!")
                
                if result.get("breakthrough_count", 0) > self.breakthrough_count:
                    logger.info(f"🎯 Total breakthroughs: {result['breakthrough_count']}")
                
                # Wait for next cycle
                await asyncio.sleep(interval_minutes * 60)
                
            except KeyboardInterrupt:
                logger.info("Stopping master improvement...")
                break
            except Exception as e:
                logger.error(f"Unexpected error in master improvement: {e}")
                await asyncio.sleep(60)
    
    def get_master_status(self) -> Dict[str, Any]:
        """Get comprehensive master system status."""
        return {
            "total_cycles": len(self.improvement_history),
            "evolution_stage": self.evolution_stage,
            "creativity_level": self.creativity_level,
            "breakthrough_count": self.breakthrough_count,
            "last_improvement": self.improvement_history[-1]["timestamp"] if self.improvement_history else None,
            "recent_breakthroughs": len([r for r in self.improvement_history[-10:] if r.get("evaluation", {}).get("innovation_level") in ["breakthrough", "revolutionary"]]),
            "system_usage": {
                "basic": len([r for r in self.improvement_history if r.get("system_used") == "basic"]),
                "advanced": len([r for r in self.improvement_history if r.get("system_used") == "advanced"]),
                "creative": len([r for r in self.improvement_history if r.get("system_used") == "creative"])
            }
        }

async def main():
    """Main entry point."""
    system = MasterAutoImprovementSystem()
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--status":
            status = system.get_master_status()
            print(json.dumps(status, indent=2))
            return
        elif sys.argv[1] == "--once":
            result = await system.run_master_improvement_cycle()
            print(json.dumps(result, indent=2))
            return
        elif sys.argv[1].startswith("--interval="):
            interval = int(sys.argv[1].split("=")[1])
            await system.run_continuous_master_improvement(interval)
            return
        elif sys.argv[1] == "--creative-only":
            # Run only creative system
            result = await system._run_improvement_system("creative")
            print(json.dumps(result, indent=2))
            return
    
    # Default: run master improvement every 20 minutes
    await system.run_continuous_master_improvement(20)

if __name__ == "__main__":
    asyncio.run(main())
