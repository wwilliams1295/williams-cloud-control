#!/usr/bin/env python3
"""
AI Governance and Safety System
==============================

A comprehensive AI governance system that ensures:
- Safe AI operations and modifications
- Ethical decision making
- Risk assessment and mitigation
- Compliance monitoring
- Human oversight and intervention
"""

import asyncio
import json
import logging
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
import hashlib
import re

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.agent import superchat
from memory.memory_system import get_memory

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/ai_governance.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class SafetyRule:
    """Individual safety rule definition."""
    
    def __init__(self, rule_id: str, description: str, severity: str, 
                 pattern: str, action: str, category: str):
        self.rule_id = rule_id
        self.description = description
        self.severity = severity  # low, medium, high, critical
        self.pattern = pattern
        self.action = action  # block, warn, log, require_approval
        self.category = category
        self.created_at = datetime.now().isoformat()
        self.violations = 0
        self.last_violation = None
    
    def check_violation(self, content: str) -> Tuple[bool, str]:
        """Check if content violates this rule."""
        if re.search(self.pattern, content, re.IGNORECASE):
            self.violations += 1
            self.last_violation = datetime.now().isoformat()
            return True, f"Rule {self.rule_id} violated: {self.description}"
        return False, ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'rule_id': self.rule_id,
            'description': self.description,
            'severity': self.severity,
            'pattern': self.pattern,
            'action': self.action,
            'category': self.category,
            'created_at': self.created_at,
            'violations': self.violations,
            'last_violation': self.last_violation
        }

class EthicalGuideline:
    """Ethical guideline for AI behavior."""
    
    def __init__(self, guideline_id: str, title: str, description: str, 
                 category: str, priority: int):
        self.guideline_id = guideline_id
        self.title = title
        self.description = description
        self.category = category
        self.priority = priority
        self.created_at = datetime.now().isoformat()
        self.compliance_score = 0.0
        self.last_assessment = None
    
    def assess_compliance(self, action: str, context: Dict[str, Any]) -> Tuple[float, str]:
        """Assess compliance with this guideline."""
        # Simplified compliance assessment
        # In production, would use more sophisticated NLP analysis
        
        compliance_score = 1.0
        issues = []
        
        # Check for ethical concerns
        if "harm" in action.lower() or "damage" in action.lower():
            compliance_score -= 0.3
            issues.append("Potential harm detected")
        
        if "bias" in action.lower() or "discrimination" in action.lower():
            compliance_score -= 0.4
            issues.append("Potential bias detected")
        
        if "privacy" in action.lower() and "consent" not in action.lower():
            compliance_score -= 0.2
            issues.append("Privacy concern without consent")
        
        self.compliance_score = max(0.0, compliance_score)
        self.last_assessment = datetime.now().isoformat()
        
        return self.compliance_score, "; ".join(issues) if issues else "Compliant"
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return {
            'guideline_id': self.guideline_id,
            'title': self.title,
            'description': self.description,
            'category': self.category,
            'priority': self.priority,
            'created_at': self.created_at,
            'compliance_score': self.compliance_score,
            'last_assessment': self.last_assessment
        }

class RiskAssessment:
    """Risk assessment for AI actions."""
    
    def __init__(self, action: str, context: Dict[str, Any]):
        self.action = action
        self.context = context
        self.risk_level = "low"
        self.risk_factors = []
        self.mitigation_strategies = []
        self.assessed_at = datetime.now().isoformat()
    
    def assess_risk(self) -> Dict[str, Any]:
        """Assess the risk level of an action."""
        risk_score = 0
        
        # Check for high-risk patterns
        high_risk_patterns = [
            r"delete.*system",
            r"modify.*core",
            r"access.*credentials",
            r"execute.*command",
            r"modify.*database"
        ]
        
        for pattern in high_risk_patterns:
            if re.search(pattern, self.action, re.IGNORECASE):
                risk_score += 2
                self.risk_factors.append(f"High-risk pattern detected: {pattern}")
        
        # Check context for additional risk factors
        if self.context.get('file_path', '').endswith(('.py', '.js', '.sql')):
            risk_score += 1
            self.risk_factors.append("Code modification detected")
        
        if self.context.get('requires_restart', False):
            risk_score += 1
            self.risk_factors.append("System restart required")
        
        if self.context.get('affects_multiple_users', False):
            risk_score += 2
            self.risk_factors.append("Multi-user impact")
        
        # Determine risk level
        if risk_score >= 5:
            self.risk_level = "critical"
        elif risk_score >= 3:
            self.risk_level = "high"
        elif risk_score >= 1:
            self.risk_level = "medium"
        else:
            self.risk_level = "low"
        
        # Generate mitigation strategies
        self._generate_mitigation_strategies()
        
        return {
            'risk_level': self.risk_level,
            'risk_score': risk_score,
            'risk_factors': self.risk_factors,
            'mitigation_strategies': self.mitigation_strategies,
            'assessed_at': self.assessed_at
        }
    
    def _generate_mitigation_strategies(self):
        """Generate mitigation strategies based on risk level."""
        if self.risk_level == "critical":
            self.mitigation_strategies = [
                "Require human approval before execution",
                "Create full system backup",
                "Test in isolated environment first",
                "Implement rollback plan",
                "Monitor system closely during execution"
            ]
        elif self.risk_level == "high":
            self.mitigation_strategies = [
                "Create backup before execution",
                "Test in staging environment",
                "Implement monitoring",
                "Prepare rollback plan"
            ]
        elif self.risk_level == "medium":
            self.mitigation_strategies = [
                "Create backup",
                "Monitor execution",
                "Test thoroughly"
            ]
        else:
            self.mitigation_strategies = [
                "Standard monitoring",
                "Log execution"
            ]

class HumanOversight:
    """Human oversight and intervention system."""
    
    def __init__(self):
        self.pending_approvals = {}
        self.approval_history = []
        self.human_operators = []
        self.escalation_rules = []
    
    def request_approval(self, action: str, context: Dict[str, Any], 
                        risk_assessment: RiskAssessment) -> str:
        """Request human approval for an action."""
        approval_id = hashlib.md5(f"{action}{time.time()}".encode()).hexdigest()[:8]
        
        self.pending_approvals[approval_id] = {
            'action': action,
            'context': context,
            'risk_assessment': risk_assessment.assess_risk(),
            'requested_at': datetime.now().isoformat(),
            'status': 'pending',
            'approved_by': None,
            'approved_at': None
        }
        
        logger.info(f"Human approval requested for action: {action}")
        return approval_id
    
    def approve_action(self, approval_id: str, operator: str, 
                      conditions: List[str] = None) -> bool:
        """Approve a pending action."""
        if approval_id not in self.pending_approvals:
            return False
        
        approval = self.pending_approvals[approval_id]
        approval['status'] = 'approved'
        approval['approved_by'] = operator
        approval['approved_at'] = datetime.now().isoformat()
        approval['conditions'] = conditions or []
        
        self.approval_history.append(approval)
        del self.pending_approvals[approval_id]
        
        logger.info(f"Action approved by {operator}: {approval['action']}")
        return True
    
    def reject_action(self, approval_id: str, operator: str, reason: str) -> bool:
        """Reject a pending action."""
        if approval_id not in self.pending_approvals:
            return False
        
        approval = self.pending_approvals[approval_id]
        approval['status'] = 'rejected'
        approval['approved_by'] = operator
        approval['approved_at'] = datetime.now().isoformat()
        approval['rejection_reason'] = reason
        
        self.approval_history.append(approval)
        del self.pending_approvals[approval_id]
        
        logger.info(f"Action rejected by {operator}: {reason}")
        return True
    
    def get_pending_approvals(self) -> List[Dict[str, Any]]:
        """Get all pending approvals."""
        return list(self.pending_approvals.values())

class ComplianceMonitor:
    """Monitor compliance with regulations and standards."""
    
    def __init__(self):
        self.compliance_frameworks = {
            'GDPR': self._check_gdpr_compliance,
            'CCPA': self._check_ccpa_compliance,
            'HIPAA': self._check_hipaa_compliance,
            'SOX': self._check_sox_compliance
        }
        self.compliance_scores = {}
        self.violations = []
    
    def check_compliance(self, action: str, context: Dict[str, Any]) -> Dict[str, Any]:
        """Check compliance with all applicable frameworks."""
        results = {}
        
        for framework, check_function in self.compliance_frameworks.items():
            try:
                score, issues = check_function(action, context)
                results[framework] = {
                    'score': score,
                    'issues': issues,
                    'compliant': score >= 0.8
                }
                
                if score < 0.8:
                    self.violations.append({
                        'framework': framework,
                        'action': action,
                        'issues': issues,
                        'timestamp': datetime.now().isoformat()
                    })
                    
            except Exception as e:
                logger.error(f"Compliance check failed for {framework}: {e}")
                results[framework] = {
                    'score': 0.0,
                    'issues': [f"Check failed: {e}"],
                    'compliant': False
                }
        
        return results
    
    def _check_gdpr_compliance(self, action: str, context: Dict[str, Any]) -> Tuple[float, List[str]]:
        """Check GDPR compliance."""
        score = 1.0
        issues = []
        
        # Check for data processing
        if any(keyword in action.lower() for keyword in ['process', 'store', 'collect', 'personal']):
            if 'consent' not in action.lower() and 'lawful_basis' not in context:
                score -= 0.3
                issues.append("Data processing without lawful basis")
            
            if 'data_protection' not in action.lower():
                score -= 0.2
                issues.append("Missing data protection measures")
        
        return score, issues
    
    def _check_ccpa_compliance(self, action: str, context: Dict[str, Any]) -> Tuple[float, List[str]]:
        """Check CCPA compliance."""
        score = 1.0
        issues = []
        
        # Check for consumer rights
        if 'personal_information' in action.lower():
            if 'opt_out' not in action.lower():
                score -= 0.2
                issues.append("Missing opt-out mechanism")
        
        return score, issues
    
    def _check_hipaa_compliance(self, action: str, context: Dict[str, Any]) -> Tuple[float, List[str]]:
        """Check HIPAA compliance."""
        score = 1.0
        issues = []
        
        # Check for PHI handling
        if any(keyword in action.lower() for keyword in ['health', 'medical', 'patient', 'phi']):
            if 'encryption' not in action.lower():
                score -= 0.4
                issues.append("PHI handling without encryption")
            
            if 'access_control' not in action.lower():
                score -= 0.3
                issues.append("Missing access controls for PHI")
        
        return score, issues
    
    def _check_sox_compliance(self, action: str, context: Dict[str, Any]) -> Tuple[float, List[str]]:
        """Check SOX compliance."""
        score = 1.0
        issues = []
        
        # Check for financial data handling
        if any(keyword in action.lower() for keyword in ['financial', 'accounting', 'audit']):
            if 'audit_trail' not in action.lower():
                score -= 0.3
                issues.append("Financial data without audit trail")
        
        return score, issues

class AIGovernanceSystem:
    """Main AI governance system."""
    
    def __init__(self):
        self.safety_rules = self._initialize_safety_rules()
        self.ethical_guidelines = self._initialize_ethical_guidelines()
        self.human_oversight = HumanOversight()
        self.compliance_monitor = ComplianceMonitor()
        self.governance_log = []
        self.memory = get_memory()
    
    def _initialize_safety_rules(self) -> List[SafetyRule]:
        """Initialize core safety rules."""
        rules = [
            SafetyRule(
                "RULE_001",
                "No direct system command execution",
                "critical",
                r"os\.system|subprocess\.call|exec\(|eval\(",
                "block",
                "code_execution"
            ),
            SafetyRule(
                "RULE_002",
                "No modification of core authentication",
                "critical",
                r"auth|login|password|credential",
                "require_approval",
                "security"
            ),
            SafetyRule(
                "RULE_003",
                "No deletion of critical files",
                "high",
                r"delete.*\.py|remove.*core|unlink.*system",
                "require_approval",
                "file_operations"
            ),
            SafetyRule(
                "RULE_004",
                "No access to sensitive data without encryption",
                "high",
                r"private|secret|key|token",
                "warn",
                "data_access"
            ),
            SafetyRule(
                "RULE_005",
                "No modification of database schema",
                "medium",
                r"alter.*table|drop.*table|create.*table",
                "require_approval",
                "database"
            )
        ]
        return rules
    
    def _initialize_ethical_guidelines(self) -> List[EthicalGuideline]:
        """Initialize ethical guidelines."""
        guidelines = [
            EthicalGuideline(
                "ETH_001",
                "Do No Harm",
                "AI should not cause harm to humans or systems",
                "safety",
                1
            ),
            EthicalGuideline(
                "ETH_002",
                "Transparency",
                "AI decisions should be transparent and explainable",
                "transparency",
                2
            ),
            EthicalGuideline(
                "ETH_003",
                "Privacy Protection",
                "AI should protect user privacy and data",
                "privacy",
                1
            ),
            EthicalGuideline(
                "ETH_004",
                "Fairness",
                "AI should treat all users fairly without bias",
                "fairness",
                2
            ),
            EthicalGuideline(
                "ETH_005",
                "Accountability",
                "AI actions should be accountable and auditable",
                "accountability",
                3
            )
        ]
        return guidelines
    
    async def evaluate_action(self, action: str, context: Dict[str, Any]) -> Dict[str, Any]:
        """Evaluate an action for safety, ethics, and compliance."""
        evaluation = {
            'action': action,
            'context': context,
            'timestamp': datetime.now().isoformat(),
            'approved': False,
            'requires_approval': False,
            'safety_violations': [],
            'ethical_assessment': {},
            'risk_assessment': {},
            'compliance_check': {},
            'recommendations': []
        }
        
        # Check safety rules
        for rule in self.safety_rules:
            violated, message = rule.check_violation(action)
            if violated:
                evaluation['safety_violations'].append({
                    'rule_id': rule.rule_id,
                    'severity': rule.severity,
                    'message': message,
                    'action': rule.action
                })
                
                if rule.action == "block":
                    evaluation['approved'] = False
                    evaluation['recommendations'].append(f"Action blocked by safety rule: {rule.description}")
                    break
                elif rule.action == "require_approval":
                    evaluation['requires_approval'] = True
        
        # Assess risk
        risk_assessment = RiskAssessment(action, context)
        evaluation['risk_assessment'] = risk_assessment.assess_risk()
        
        # Check ethical guidelines
        for guideline in self.ethical_guidelines:
            score, issues = guideline.assess_compliance(action, context)
            evaluation['ethical_assessment'][guideline.guideline_id] = {
                'title': guideline.title,
                'score': score,
                'issues': issues
            }
        
        # Check compliance
        evaluation['compliance_check'] = self.compliance_monitor.check_compliance(action, context)
        
        # Determine final approval
        if not evaluation['safety_violations'] or all(
            v['action'] != 'block' for v in evaluation['safety_violations']
        ):
            if evaluation['requires_approval'] or evaluation['risk_assessment']['risk_level'] in ['high', 'critical']:
                # Request human approval
                approval_id = self.human_oversight.request_approval(
                    action, context, risk_assessment
                )
                evaluation['approval_id'] = approval_id
                evaluation['status'] = 'pending_approval'
            else:
                evaluation['approved'] = True
                evaluation['status'] = 'approved'
        else:
            evaluation['status'] = 'blocked'
        
        # Log evaluation
        self.governance_log.append(evaluation)
        
        return evaluation
    
    def get_governance_status(self) -> Dict[str, Any]:
        """Get current governance system status."""
        return {
            'timestamp': datetime.now().isoformat(),
            'safety_rules': len(self.safety_rules),
            'ethical_guidelines': len(self.ethical_guidelines),
            'pending_approvals': len(self.human_oversight.pending_approvals),
            'total_evaluations': len(self.governance_log),
            'recent_violations': len([
                v for v in self.compliance_monitor.violations
                if datetime.fromisoformat(v['timestamp']) > datetime.now() - timedelta(hours=24)
            ])
        }
    
    def get_safety_report(self) -> Dict[str, Any]:
        """Get safety report."""
        return {
            'timestamp': datetime.now().isoformat(),
            'safety_rules': [rule.to_dict() for rule in self.safety_rules],
            'recent_violations': [
                v for v in self.governance_log
                if v.get('safety_violations') and 
                datetime.fromisoformat(v['timestamp']) > datetime.now() - timedelta(hours=24)
            ]
        }

# Main execution
async def main():
    """Main execution function."""
    governance = AIGovernanceSystem()
    
    # Example: Evaluate an action
    action = "Create a new plugin for user authentication"
    context = {
        'file_path': 'src/plugins/auth_plugin.py',
        'requires_restart': False,
        'affects_multiple_users': True
    }
    
    evaluation = await governance.evaluate_action(action, context)
    print(f"Action evaluation: {json.dumps(evaluation, indent=2)}")
    
    # Get governance status
    status = governance.get_governance_status()
    print(f"Governance status: {json.dumps(status, indent=2)}")

if __name__ == "__main__":
    asyncio.run(main())
