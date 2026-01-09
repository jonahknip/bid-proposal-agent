"""
Bid Analyzer - Compare bid proposals against requirements and extracted quantities
Analyzes completeness, accuracy, and generates recommendations
"""

import os
import json
import re
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
from dataclasses import dataclass, field
from difflib import SequenceMatcher

from openai import OpenAI


@dataclass
class AnalysisResult:
    """Result of analyzing a bid against requirements"""
    completeness_score: float = 0.0
    accuracy_score: float = 0.0
    overall_score: float = 0.0
    matches: List[Dict[str, Any]] = field(default_factory=list)
    discrepancies: List[Dict[str, Any]] = field(default_factory=list)
    missing_items: List[Dict[str, Any]] = field(default_factory=list)
    extra_items: List[Dict[str, Any]] = field(default_factory=list)
    recommendations: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    critical_issues: List[str] = field(default_factory=list)


class BidAnalyzer:
    """
    Analyzes bid proposals against RFP requirements and plan quantities.
    Identifies discrepancies, missing items, and provides recommendations.
    """
    
    # Bid analysis prompt
    BID_ANALYSIS_PROMPT = """You are an expert civil engineering bid analyst reviewing a bid proposal.

Compare the PROPOSED BID against the RFP REQUIREMENTS and PLAN QUANTITIES.

Analyze for:
1. COMPLETENESS - Are all required line items included?
2. ACCURACY - Do quantities match the plans?
3. PRICING - Are unit prices reasonable for the work?
4. REQUIREMENTS - Are all RFP requirements addressed?

For each discrepancy found, assess:
- Severity: CRITICAL (bid could be rejected), WARNING (should fix), INFO (minor issue)
- Impact: Cost impact, schedule impact, compliance risk
- Recommendation: How to resolve

Provide specific, actionable feedback.

Return a JSON object:
{
    "summary": {
        "completeness_score": 0-100,
        "accuracy_score": 0-100,
        "overall_assessment": "brief assessment",
        "recommendation": "go/no-go/revise"
    },
    "line_item_analysis": [
        {
            "item": "description",
            "status": "match/discrepancy/missing/extra",
            "proposed_qty": 0,
            "required_qty": 0,
            "variance_pct": 0,
            "severity": "critical/warning/info",
            "notes": "explanation"
        }
    ],
    "critical_issues": ["list of critical issues that must be addressed"],
    "warnings": ["list of warnings to consider"],
    "recommendations": ["list of recommendations to improve bid"],
    "missing_requirements": ["list of RFP requirements not addressed"],
    "cost_analysis": {
        "total_proposed": 0,
        "estimated_should_be": 0,
        "variance": 0,
        "notes": ""
    }
}

Only return the JSON object, no other text."""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize the bid analyzer with OpenAI API key."""
        self.api_key = api_key or os.environ.get('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OpenAI API key is required")
        self.client = OpenAI(api_key=self.api_key)
        self.model = "gpt-4o"
    
    def analyze_bid(
        self,
        proposal_requirements: Dict[str, Any],
        bid_proposal: Dict[str, Any],
        plan_quantities: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Analyze a bid proposal against requirements and plan quantities.
        
        Args:
            proposal_requirements: Extracted requirements from RFP
            bid_proposal: The bid proposal being reviewed
            plan_quantities: Optional quantities extracted from plans
            
        Returns:
            Analysis results with scores and recommendations
        """
        # Build analysis context
        context = self._build_analysis_context(
            proposal_requirements,
            bid_proposal,
            plan_quantities
        )
        
        # Use GPT for comprehensive analysis
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "user", "content": f"{self.BID_ANALYSIS_PROMPT}\n\n{context}"}
            ],
            max_tokens=4000,
            temperature=0.3
        )
        
        try:
            result_text = response.choices[0].message.content
            if result_text.startswith('```'):
                result_text = result_text.split('```')[1]
                if result_text.startswith('json'):
                    result_text = result_text[4:]
            if result_text.endswith('```'):
                result_text = result_text[:-3]
            
            ai_analysis = json.loads(result_text.strip())
            
        except json.JSONDecodeError:
            content_text = response.choices[0].message.content
            start = content_text.find('{')
            end = content_text.rfind('}') + 1
            if start != -1 and end > start:
                ai_analysis = json.loads(content_text[start:end])
            else:
                ai_analysis = {
                    'summary': {'completeness_score': 0, 'accuracy_score': 0},
                    'error': 'Failed to analyze bid'
                }
        
        # Add rule-based analysis
        rule_analysis = self._rule_based_analysis(
            proposal_requirements,
            bid_proposal,
            plan_quantities
        )
        
        # Combine analyses
        combined = self._combine_analyses(ai_analysis, rule_analysis)
        combined['timestamp'] = datetime.now().isoformat()
        
        return combined
    
    def _build_analysis_context(
        self,
        requirements: Dict[str, Any],
        bid: Dict[str, Any],
        quantities: Optional[Dict[str, Any]]
    ) -> str:
        """Build context string for AI analysis."""
        
        context = "=== RFP REQUIREMENTS ===\n"
        
        # Project info
        project = requirements.get('project_info', {})
        context += f"Project: {project.get('project_name', 'N/A')}\n"
        context += f"Owner: {project.get('owner', 'N/A')}\n"
        context += f"Location: {project.get('location', 'N/A')}\n\n"
        
        # Bid schedule
        schedule = requirements.get('bid_schedule', {})
        if schedule:
            context += f"Bid Due: {schedule.get('bid_date', 'N/A')}\n\n"
        
        # Required line items
        req_items = requirements.get('line_items', []) or requirements.get('combined_line_items', [])
        context += "REQUIRED LINE ITEMS:\n"
        for item in req_items[:50]:  # Limit for token efficiency
            context += f"- {item.get('item_number', '')}: {item.get('description', '')} "
            context += f"| Qty: {item.get('quantity', 'N/A')} {item.get('unit', '')}\n"
        
        # Requirements
        reqs = requirements.get('requirements', {})
        if reqs:
            context += "\nSPECIAL REQUIREMENTS:\n"
            if reqs.get('bonding'):
                context += f"- Bonding: {reqs['bonding']}\n"
            if reqs.get('insurance'):
                context += f"- Insurance: {reqs['insurance']}\n"
            for qual in reqs.get('qualifications', [])[:5]:
                context += f"- Qualification: {qual}\n"
        
        context += "\n=== BID PROPOSAL ===\n"
        
        # Bid line items
        bid_items = bid.get('line_items', []) or bid.get('combined_line_items', [])
        context += "PROPOSED LINE ITEMS:\n"
        for item in bid_items[:50]:
            context += f"- {item.get('item_number', '')}: {item.get('description', '')} "
            context += f"| Qty: {item.get('quantity', 'N/A')} {item.get('unit', '')} "
            if item.get('unit_price'):
                context += f"@ ${item.get('unit_price', 0)}/unit"
            context += "\n"
        
        # Plan quantities if available
        if quantities:
            context += "\n=== PLAN QUANTITIES (from drawings) ===\n"
            all_qty = quantities.get('all_quantities', [])
            for qty in all_qty[:30]:
                context += f"- {qty.get('description', '')}: "
                context += f"{qty.get('quantity', 'N/A')} {qty.get('unit', '')}\n"
        
        return context
    
    def _rule_based_analysis(
        self,
        requirements: Dict[str, Any],
        bid: Dict[str, Any],
        quantities: Optional[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Perform rule-based analysis to supplement AI."""
        
        analysis = {
            'matches': [],
            'discrepancies': [],
            'missing': [],
            'extra': []
        }
        
        req_items = requirements.get('line_items', []) or requirements.get('combined_line_items', [])
        bid_items = bid.get('line_items', []) or bid.get('combined_line_items', [])
        
        # Normalize for matching
        def normalize(text):
            return re.sub(r'[^a-z0-9]', '', str(text).lower())
        
        req_lookup = {}
        for item in req_items:
            key = normalize(item.get('description', ''))
            req_lookup[key] = item
        
        bid_lookup = {}
        for item in bid_items:
            key = normalize(item.get('description', ''))
            bid_lookup[key] = item
        
        # Check required items against bid
        for key, req in req_lookup.items():
            if key in bid_lookup:
                bid_item = bid_lookup[key]
                
                req_qty = float(req.get('quantity', 0) or 0)
                bid_qty = float(bid_item.get('quantity', 0) or 0)
                
                if req_qty > 0:
                    variance = abs(bid_qty - req_qty) / req_qty
                else:
                    variance = 0
                
                if variance <= 0.05:  # Within 5%
                    analysis['matches'].append({
                        'description': req.get('description'),
                        'required_qty': req_qty,
                        'proposed_qty': bid_qty,
                        'unit': req.get('unit'),
                        'variance_pct': round(variance * 100, 1)
                    })
                else:
                    analysis['discrepancies'].append({
                        'description': req.get('description'),
                        'required_qty': req_qty,
                        'proposed_qty': bid_qty,
                        'unit': req.get('unit'),
                        'variance_pct': round(variance * 100, 1),
                        'difference': round(bid_qty - req_qty, 2)
                    })
            else:
                # Try fuzzy matching
                best_match = None
                best_ratio = 0
                for bid_key in bid_lookup.keys():
                    ratio = SequenceMatcher(None, key, bid_key).ratio()
                    if ratio > best_ratio and ratio > 0.7:
                        best_ratio = ratio
                        best_match = bid_key
                
                if best_match:
                    bid_item = bid_lookup[best_match]
                    analysis['discrepancies'].append({
                        'description': req.get('description'),
                        'required_qty': float(req.get('quantity', 0) or 0),
                        'proposed_qty': float(bid_item.get('quantity', 0) or 0),
                        'unit': req.get('unit'),
                        'note': f'Fuzzy match: {bid_item.get("description")}'
                    })
                else:
                    analysis['missing'].append(req)
        
        # Check for extra items in bid
        for key, bid_item in bid_lookup.items():
            if key not in req_lookup:
                # Check if it's a fuzzy match
                is_matched = False
                for req_key in req_lookup.keys():
                    if SequenceMatcher(None, key, req_key).ratio() > 0.7:
                        is_matched = True
                        break
                
                if not is_matched:
                    analysis['extra'].append(bid_item)
        
        # Calculate scores
        total_required = len(req_items)
        matched = len(analysis['matches'])
        
        analysis['completeness_score'] = round(
            (matched + len(analysis['discrepancies'])) / max(total_required, 1) * 100, 1
        )
        analysis['accuracy_score'] = round(
            matched / max(matched + len(analysis['discrepancies']), 1) * 100, 1
        )
        
        return analysis
    
    def _combine_analyses(
        self,
        ai_analysis: Dict[str, Any],
        rule_analysis: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Combine AI and rule-based analyses."""
        
        combined = {
            'summary': ai_analysis.get('summary', {}),
            'line_item_analysis': ai_analysis.get('line_item_analysis', []),
            'critical_issues': ai_analysis.get('critical_issues', []),
            'warnings': ai_analysis.get('warnings', []),
            'recommendations': ai_analysis.get('recommendations', []),
            'missing_requirements': ai_analysis.get('missing_requirements', []),
            'cost_analysis': ai_analysis.get('cost_analysis', {}),
            'rule_based': {
                'matches': rule_analysis.get('matches', []),
                'discrepancies': rule_analysis.get('discrepancies', []),
                'missing': rule_analysis.get('missing', []),
                'extra': rule_analysis.get('extra', []),
                'completeness_score': rule_analysis.get('completeness_score', 0),
                'accuracy_score': rule_analysis.get('accuracy_score', 0)
            }
        }
        
        # Add rule-based findings to warnings if not captured by AI
        if rule_analysis.get('missing'):
            for item in rule_analysis['missing'][:5]:
                warning = f"Missing item: {item.get('description', 'Unknown')}"
                if warning not in combined['warnings']:
                    combined['warnings'].append(warning)
        
        return combined
    
    def compare_quantities(
        self,
        proposal_quantities: List[Dict[str, Any]],
        plan_quantities: List[Dict[str, Any]],
        tolerance: float = 0.10
    ) -> Dict[str, Any]:
        """
        Compare quantities from proposal against quantities from plans.
        
        Args:
            proposal_quantities: Quantities from the bid proposal
            plan_quantities: Quantities extracted from plans
            tolerance: Acceptable variance (0.10 = 10%)
            
        Returns:
            Comparison results
        """
        results = {
            'matches': [],
            'over_estimated': [],
            'under_estimated': [],
            'not_on_plans': [],
            'not_in_proposal': [],
            'summary': {}
        }
        
        def normalize(text):
            return re.sub(r'[^a-z0-9]', '', str(text).lower())
        
        prop_lookup = {normalize(p.get('description', '')): p for p in proposal_quantities}
        plan_lookup = {normalize(p.get('description', '')): p for p in plan_quantities}
        
        # Compare each proposal quantity against plans
        for key, prop in prop_lookup.items():
            if key in plan_lookup:
                plan = plan_lookup[key]
                prop_qty = float(prop.get('quantity', 0) or 0)
                plan_qty = float(plan.get('quantity', 0) or 0)
                
                if plan_qty > 0:
                    variance = (prop_qty - plan_qty) / plan_qty
                else:
                    variance = 1.0 if prop_qty > 0 else 0.0
                
                item_result = {
                    'description': prop.get('description'),
                    'proposal_qty': prop_qty,
                    'plan_qty': plan_qty,
                    'unit': prop.get('unit'),
                    'variance_pct': round(variance * 100, 1),
                    'difference': round(prop_qty - plan_qty, 2)
                }
                
                if abs(variance) <= tolerance:
                    results['matches'].append(item_result)
                elif variance > tolerance:
                    results['over_estimated'].append(item_result)
                else:
                    results['under_estimated'].append(item_result)
            else:
                results['not_on_plans'].append(prop)
        
        # Find items on plans but not in proposal
        for key, plan in plan_lookup.items():
            if key not in prop_lookup:
                results['not_in_proposal'].append(plan)
        
        # Summary
        total = len(proposal_quantities)
        results['summary'] = {
            'total_items': total,
            'matches': len(results['matches']),
            'over_estimated': len(results['over_estimated']),
            'under_estimated': len(results['under_estimated']),
            'match_rate': round(len(results['matches']) / max(total, 1) * 100, 1),
            'items_not_on_plans': len(results['not_on_plans']),
            'items_missing_from_proposal': len(results['not_in_proposal'])
        }
        
        return results
    
    def generate_recommendations(self, analysis: Dict[str, Any]) -> List[Dict[str, str]]:
        """
        Generate prioritized recommendations based on analysis.
        
        Args:
            analysis: Complete analysis results
            
        Returns:
            List of prioritized recommendations
        """
        recommendations = []
        
        # Critical issues first
        for issue in analysis.get('critical_issues', []):
            recommendations.append({
                'priority': 'CRITICAL',
                'action': issue,
                'reason': 'May result in bid rejection'
            })
        
        # Under-estimated quantities
        under = analysis.get('rule_based', {}).get('discrepancies', [])
        for item in under:
            if item.get('variance_pct', 0) < -10:
                recommendations.append({
                    'priority': 'HIGH',
                    'action': f"Review quantity for: {item.get('description')}",
                    'reason': f"Proposal is {abs(item.get('variance_pct', 0)):.1f}% under requirement"
                })
        
        # Missing items
        missing = analysis.get('rule_based', {}).get('missing', [])
        for item in missing[:5]:
            recommendations.append({
                'priority': 'HIGH',
                'action': f"Add missing item: {item.get('description')}",
                'reason': 'Required by RFP but not in proposal'
            })
        
        # Warnings
        for warning in analysis.get('warnings', []):
            if warning not in [r['action'] for r in recommendations]:
                recommendations.append({
                    'priority': 'MEDIUM',
                    'action': warning,
                    'reason': 'Potential issue identified'
                })
        
        # AI recommendations
        for rec in analysis.get('recommendations', []):
            if rec not in [r['action'] for r in recommendations]:
                recommendations.append({
                    'priority': 'LOW',
                    'action': rec,
                    'reason': 'Suggested improvement'
                })
        
        return recommendations
    
    def get_bid_status(self, analysis: Dict[str, Any]) -> Dict[str, str]:
        """
        Determine overall bid status and recommendation.
        
        Args:
            analysis: Complete analysis results
            
        Returns:
            Status assessment
        """
        summary = analysis.get('summary', {})
        rule_based = analysis.get('rule_based', {})
        
        completeness = summary.get('completeness_score', 0) or rule_based.get('completeness_score', 0)
        accuracy = summary.get('accuracy_score', 0) or rule_based.get('accuracy_score', 0)
        
        critical_count = len(analysis.get('critical_issues', []))
        warning_count = len(analysis.get('warnings', []))
        missing_count = len(rule_based.get('missing', []))
        
        # Determine status
        if critical_count > 0:
            status = 'NOT_READY'
            color = 'red'
            message = f'{critical_count} critical issue(s) must be resolved before submission'
        elif completeness < 80:
            status = 'INCOMPLETE'
            color = 'orange'
            message = f'Bid is {completeness:.0f}% complete - {missing_count} items missing'
        elif accuracy < 80:
            status = 'NEEDS_REVIEW'
            color = 'orange'
            message = f'Quantity accuracy is {accuracy:.0f}% - review discrepancies'
        elif warning_count > 5:
            status = 'REVIEW_WARNINGS'
            color = 'yellow'
            message = f'{warning_count} warnings to review before submission'
        else:
            status = 'READY'
            color = 'green'
            message = 'Bid appears complete and ready for final review'
        
        return {
            'status': status,
            'color': color,
            'message': message,
            'completeness_score': completeness,
            'accuracy_score': accuracy,
            'critical_issues': critical_count,
            'warnings': warning_count
        }
