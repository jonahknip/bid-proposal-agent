"""
Bid Analyzer - Expert civil engineering bid analysis and recommendations
Acts as an experienced estimator reviewing and improving bids
"""

import os
import json
import re
from typing import List, Dict, Any, Optional
from datetime import datetime

from openai import OpenAI


class BidAnalyzer:
    """
    Expert civil engineering bid analyzer.
    Reviews proposals, identifies issues, and provides recommendations.
    """
    
    EXPERT_ANALYSIS_PROMPT = """You are a senior civil engineering estimator with 25+ years of experience winning competitive bids for municipal infrastructure projects.

You are reviewing a bid proposal to ensure it is competitive, complete, and profitable.

Analyze the proposal and provide expert feedback:

1. COMPLETENESS CHECK:
   - Are all required bid items included?
   - Are quantities reasonable for the scope?
   - Are any typical items missing?

2. PRICING ANALYSIS:
   For each line item, assess:
   - Is the unit price competitive for the local market?
   - Material cost: Is it accurate for current market?
   - Labor cost: Does it reflect prevailing wages if applicable?
   - Equipment cost: Is it reasonable for the work?
   - Is the markup appropriate (typically 15-25% OH&P)?

3. RISK ASSESSMENT:
   - Identify potential cost risks
   - Flag any unusual specifications
   - Note items that may need clarification
   - Highlight scope gaps or ambiguities

4. COMPETITIVE POSITIONING:
   - Overall bid competitiveness (1-10 scale)
   - Items that may be too high (risk losing bid)
   - Items that may be too low (risk losing money)
   - Recommended adjustments

5. STRATEGIC RECOMMENDATIONS:
   - Which items to sharpen pricing on
   - Which items need more contingency
   - Value engineering opportunities
   - Bid strategy suggestions

Return a JSON object:
{
    "overall_assessment": {
        "status": "ready/needs_work/not_ready",
        "competitiveness_score": 1-10,
        "confidence_score": 1-10,
        "summary": "Executive summary of the bid"
    },
    "completeness": {
        "score": 0-100,
        "missing_items": [{"item": "", "estimated_cost": 0, "impact": "high/medium/low"}],
        "incomplete_items": [{"item": "", "issue": ""}]
    },
    "pricing_analysis": {
        "total_bid": 0,
        "recommended_total": 0,
        "variance_pct": 0,
        "line_items": [
            {
                "item": "",
                "current_price": 0,
                "recommended_price": 0,
                "status": "good/high/low/review",
                "material_assessment": "",
                "labor_assessment": "",
                "equipment_assessment": "",
                "notes": ""
            }
        ]
    },
    "risks": [
        {
            "risk": "",
            "severity": "high/medium/low",
            "potential_cost": 0,
            "mitigation": ""
        }
    ],
    "cost_breakdown": {
        "material_pct": 0,
        "labor_pct": 0,
        "equipment_pct": 0,
        "overhead_profit_pct": 0,
        "assessment": "typical/material_heavy/labor_heavy/etc"
    },
    "recommendations": [
        {
            "priority": "critical/high/medium/low",
            "action": "",
            "rationale": "",
            "estimated_impact": ""
        }
    ],
    "bid_strategy": {
        "approach": "",
        "key_focus_areas": [],
        "items_to_sharpen": [],
        "items_needing_contingency": [],
        "value_engineering_opportunities": []
    },
    "final_recommendation": "submit/revise/do_not_bid"
}

Only return the JSON object, no other text."""

    START_PROPOSAL_PROMPT = """You are an expert civil engineering estimator helping to START a new bid proposal.

Based on the bid documents provided, create a complete preliminary bid estimate with:

1. All line items from the bid schedule
2. Recommended quantities (verify against documents)
3. Material, labor, and equipment cost estimates
4. Unit prices with appropriate markup
5. Total bid amount

Use current market rates for the region. Apply typical civil engineering cost factors.

For each line item, provide:
- Material cost breakdown
- Labor cost (with crew and production rate)
- Equipment cost
- Overhead and profit (use 18% as standard)

Return a JSON object with the complete bid proposal:
{
    "project_info": {
        "project_name": "",
        "project_number": "",
        "owner": "",
        "location": "",
        "bid_date": ""
    },
    "bid_items": [
        {
            "item_number": "",
            "description": "",
            "quantity": 0,
            "unit": "",
            "material": {"cost": 0, "description": ""},
            "labor": {"cost": 0, "crew": "", "production": ""},
            "equipment": {"cost": 0, "items": []},
            "overhead_profit": 0,
            "unit_price": 0,
            "total_price": 0,
            "notes": ""
        }
    ],
    "summary": {
        "subtotal": 0,
        "contingency_pct": 5,
        "contingency_amt": 0,
        "total_bid": 0
    },
    "assumptions": [],
    "clarifications_needed": [],
    "risks": []
}

Only return the JSON object, no other text."""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize the bid analyzer with OpenAI API key."""
        self.api_key = api_key or os.environ.get('OPENAI_API_KEY') or os.environ.get('OPENAI_KEY')
        if not self.api_key:
            raise ValueError("OpenAI API key not found. Set OPENAI_API_KEY environment variable.")
        self.client = OpenAI(api_key=self.api_key.strip())
        self.model = "gpt-4o"
    
    def analyze_proposal(self, proposal_data: Dict[str, Any], bid_docs_data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Analyze a bid proposal with expert feedback.
        
        Args:
            proposal_data: The proposal being reviewed
            bid_docs_data: Original bid document requirements (optional)
            
        Returns:
            Expert analysis with recommendations
        """
        context = f"PROPOSAL BEING REVIEWED:\n{json.dumps(proposal_data, indent=2)[:30000]}"
        
        if bid_docs_data:
            context += f"\n\nORIGINAL BID REQUIREMENTS:\n{json.dumps(bid_docs_data, indent=2)[:20000]}"
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": "You are a senior civil engineering estimator with expert knowledge of competitive bidding."
                },
                {
                    "role": "user",
                    "content": f"{self.EXPERT_ANALYSIS_PROMPT}\n\n{context}"
                }
            ],
            max_tokens=6000,
            temperature=0.3
        )
        
        return self._parse_response(response.choices[0].message.content)
    
    def start_proposal(self, bid_docs_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Start a new proposal based on bid documents.
        
        Args:
            bid_docs_data: Parsed bid document data
            
        Returns:
            Complete preliminary bid proposal
        """
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert civil engineering estimator creating competitive bid proposals."
                },
                {
                    "role": "user",
                    "content": f"{self.START_PROPOSAL_PROMPT}\n\nBID DOCUMENTS:\n{json.dumps(bid_docs_data, indent=2)[:40000]}"
                }
            ],
            max_tokens=8000,
            temperature=0.3
        )
        
        return self._parse_response(response.choices[0].message.content)
    
    def _parse_response(self, content: str) -> Dict[str, Any]:
        """Parse GPT response to JSON."""
        try:
            if content.startswith('```'):
                content = content.split('```')[1]
                if content.startswith('json'):
                    content = content[4:]
            if content.endswith('```'):
                content = content[:-3]
            
            return json.loads(content.strip())
            
        except json.JSONDecodeError:
            start = content.find('{')
            end = content.rfind('}') + 1
            if start != -1 and end > start:
                try:
                    return json.loads(content[start:end])
                except:
                    pass
            
            return {
                'error': 'Failed to parse response',
                'raw_content': content[:2000]
            }
    
    def get_bid_status(self, analysis: Dict[str, Any]) -> Dict[str, Any]:
        """Get simplified bid status from analysis."""
        overall = analysis.get('overall_assessment', {})
        
        status = overall.get('status', 'needs_work')
        comp_score = overall.get('competitiveness_score', 5)
        
        if status == 'ready' or comp_score >= 8:
            color = 'green'
            message = 'Bid is competitive and ready for submission'
        elif status == 'not_ready' or comp_score <= 4:
            color = 'red'
            message = 'Significant revisions needed before submission'
        else:
            color = 'orange'
            message = 'Review recommendations before submission'
        
        return {
            'status': status.upper().replace('_', ' '),
            'color': color,
            'message': message,
            'competitiveness_score': comp_score,
            'confidence_score': overall.get('confidence_score', 5),
            'final_recommendation': analysis.get('final_recommendation', 'revise')
        }
    
    def generate_recommendations(self, analysis: Dict[str, Any]) -> List[Dict[str, str]]:
        """Generate prioritized recommendations list."""
        recommendations = []
        
        # Critical items first
        for rec in analysis.get('recommendations', []):
            if rec.get('priority') == 'critical':
                recommendations.append({
                    'priority': 'CRITICAL',
                    'action': rec.get('action', ''),
                    'rationale': rec.get('rationale', ''),
                    'impact': rec.get('estimated_impact', '')
                })
        
        # Missing items
        for item in analysis.get('completeness', {}).get('missing_items', []):
            if item.get('impact') == 'high':
                recommendations.append({
                    'priority': 'HIGH',
                    'action': f"Add missing item: {item.get('item', '')}",
                    'rationale': 'Required item not in proposal',
                    'impact': f"Est. ${item.get('estimated_cost', 0):,.0f}"
                })
        
        # Pricing issues
        for item in analysis.get('pricing_analysis', {}).get('line_items', []):
            if item.get('status') == 'low':
                recommendations.append({
                    'priority': 'HIGH',
                    'action': f"Review pricing: {item.get('item', '')}",
                    'rationale': 'Price may be too low - risk of losing money',
                    'impact': item.get('notes', '')
                })
            elif item.get('status') == 'high':
                recommendations.append({
                    'priority': 'MEDIUM',
                    'action': f"Consider reducing: {item.get('item', '')}",
                    'rationale': 'Price may be too high - risk of losing bid',
                    'impact': item.get('notes', '')
                })
        
        # High severity risks
        for risk in analysis.get('risks', []):
            if risk.get('severity') == 'high':
                recommendations.append({
                    'priority': 'HIGH',
                    'action': f"Address risk: {risk.get('risk', '')}",
                    'rationale': risk.get('mitigation', ''),
                    'impact': f"Potential cost: ${risk.get('potential_cost', 0):,.0f}"
                })
        
        # Remaining recommendations
        for rec in analysis.get('recommendations', []):
            if rec.get('priority') in ['high', 'medium']:
                if not any(r['action'] == rec.get('action') for r in recommendations):
                    recommendations.append({
                        'priority': rec.get('priority', 'MEDIUM').upper(),
                        'action': rec.get('action', ''),
                        'rationale': rec.get('rationale', ''),
                        'impact': rec.get('estimated_impact', '')
                    })
        
        return recommendations[:15]  # Limit to top 15
    
    def format_currency(self, amount: float) -> str:
        """Format number as currency."""
        return f"${amount:,.2f}"
