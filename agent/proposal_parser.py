"""
Proposal Parser - Extract bid specifications, requirements, and line items from bid documents
Optimized for civil engineering design/survey work bidding
"""

import os
import io
import json
import re
from typing import List, Dict, Any, Optional
from datetime import datetime

import fitz  # PyMuPDF
from openai import OpenAI
import openpyxl


class ProposalParser:
    """
    Parses RFP/bid documents to extract specifications and requirements
    for accurate civil engineering bid estimates.
    """
    
    BID_DOC_EXTRACTION_PROMPT = """You are an expert civil engineering bid analyst extracting information from bid documents.

Extract ALL relevant information for preparing an accurate bid proposal:

1. PROJECT IDENTIFICATION:
   - Project name and number
   - Owner/Agency
   - Location (city, county, state)
   - Engineer of Record
   - Project type (road, utility, site development, survey, etc.)

2. BID SCHEDULE & DEADLINES:
   - Bid due date and time
   - Pre-bid meeting (date, time, location, mandatory?)
   - Site visit information
   - Question deadline
   - Award date (if stated)

3. SCOPE OF WORK:
   - Detailed description of work
   - Project limits (stations, addresses, boundaries)
   - Major work elements
   - Phases or sequences required
   - Working days or calendar days allowed

4. LINE ITEMS / BID SCHEDULE:
   For EACH pay item, extract:
   - Item number
   - Description (exact wording)
   - Quantity
   - Unit of measure
   - Any notes or specifications

5. SPECIFICATIONS & STANDARDS:
   - Referenced specs (MDOT, local agency, etc.)
   - Material requirements
   - Testing requirements
   - Quality standards
   - Special provisions

6. SPECIAL CONDITIONS:
   - Prevailing wage requirements
   - DBE/MBE goals
   - Bonding requirements
   - Insurance requirements
   - Liquidated damages
   - Retainage
   - Traffic control requirements
   - Working hour restrictions
   - Environmental constraints
   - Permit requirements

7. CONTACTS:
   - Project manager
   - Bid contact
   - Questions to

Return a JSON object:
{
    "project_info": {
        "project_name": "",
        "project_number": "",
        "owner": "",
        "location": "",
        "engineer": "",
        "project_type": ""
    },
    "bid_schedule": {
        "bid_date": "",
        "bid_time": "",
        "pre_bid_meeting": {"date": "", "time": "", "location": "", "mandatory": false},
        "question_deadline": "",
        "site_visit": ""
    },
    "scope": {
        "description": "",
        "limits": "",
        "major_elements": [],
        "duration": "",
        "phases": []
    },
    "line_items": [
        {
            "item_number": "",
            "description": "",
            "quantity": 0,
            "unit": "",
            "spec_reference": "",
            "notes": ""
        }
    ],
    "specifications": {
        "standard_specs": [],
        "special_provisions": [],
        "material_requirements": [],
        "testing_requirements": []
    },
    "requirements": {
        "prevailing_wage": false,
        "dbe_goal": "",
        "bonding": "",
        "insurance": "",
        "liquidated_damages": "",
        "retainage": "",
        "working_hours": "",
        "traffic_control": "",
        "permits_required": []
    },
    "contacts": [
        {"name": "", "title": "", "email": "", "phone": ""}
    ],
    "key_dates": [
        {"event": "", "date": ""}
    ],
    "risks_notes": []
}

Only return the JSON object, no other text."""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize the proposal parser with OpenAI API key."""
        self.api_key = api_key or os.environ.get('OPENAI_API_KEY') or os.environ.get('OPENAI_KEY')
        if not self.api_key:
            available_keys = [k for k in os.environ.keys() if 'OPENAI' in k.upper() or 'API' in k.upper()]
            raise ValueError(f"OpenAI API key not found. Set OPENAI_API_KEY environment variable. Found env vars with API/OPENAI: {available_keys}")
        self.client = OpenAI(api_key=self.api_key.strip())
        self.model = "gpt-4o"
    
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract all text from a PDF."""
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text() + "\n"
        doc.close()
        return text
    
    def extract_from_excel(self, excel_path: str) -> str:
        """Extract content from Excel file."""
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        text = ""
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            text += f"\n=== SHEET: {sheet_name} ===\n"
            
            for row in ws.iter_rows(values_only=True):
                if any(cell is not None for cell in row):
                    row_text = " | ".join([str(cell) if cell is not None else "" for cell in row])
                    text += row_text + "\n"
        
        wb.close()
        return text
    
    def parse_bid_document(self, file_path: str) -> Dict[str, Any]:
        """
        Parse a single bid document.
        
        Args:
            file_path: Path to the document
            
        Returns:
            Extracted bid information
        """
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.pdf':
            text = self.extract_text_from_pdf(file_path)
        elif ext in ['.xlsx', '.xls', '.xlsm']:
            text = self.extract_from_excel(file_path)
        else:
            raise ValueError(f"Unsupported file type: {ext}")
        
        # Limit text length
        if len(text) > 50000:
            text = text[:50000] + "\n[... truncated ...]"
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "user",
                    "content": f"{self.BID_DOC_EXTRACTION_PROMPT}\n\nDOCUMENT CONTENT:\n{text}"
                }
            ],
            max_tokens=6000,
            temperature=0.2
        )
        
        result = self._parse_response(response.choices[0].message.content)
        result['source_file'] = os.path.basename(file_path)
        return result
    
    def parse_multiple_documents(self, file_paths: List[str]) -> Dict[str, Any]:
        """
        Parse multiple bid documents and combine results.
        
        Args:
            file_paths: List of document paths
            
        Returns:
            Combined extracted information
        """
        # Combine all document text
        combined_text = ""
        for path in file_paths:
            ext = os.path.splitext(path)[1].lower()
            combined_text += f"\n--- Document: {os.path.basename(path)} ---\n"
            
            if ext == '.pdf':
                combined_text += self.extract_text_from_pdf(path)
            elif ext in ['.xlsx', '.xls', '.xlsm']:
                combined_text += self.extract_from_excel(path)
        
        # Limit text
        if len(combined_text) > 60000:
            combined_text = combined_text[:60000] + "\n[... truncated ...]"
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "user",
                    "content": f"{self.BID_DOC_EXTRACTION_PROMPT}\n\nDOCUMENT CONTENT:\n{combined_text}"
                }
            ],
            max_tokens=6000,
            temperature=0.2
        )
        
        result = self._parse_response(response.choices[0].message.content)
        result['source_files'] = [os.path.basename(p) for p in file_paths]
        result['files_processed'] = len(file_paths)
        return result
    
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
                'error': 'Failed to parse document',
                'raw_content': content[:2000]
            }
    
    def extract_line_items_table(self, result: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Extract line items in table format."""
        items = result.get('line_items', [])
        
        clean_items = []
        for i, item in enumerate(items, 1):
            clean_items.append({
                'No.': item.get('item_number', str(i)),
                'Description': item.get('description', ''),
                'Quantity': item.get('quantity', 0),
                'Unit': item.get('unit', ''),
                'Spec': item.get('spec_reference', ''),
                'Notes': item.get('notes', '')
            })
        
        return clean_items
    
    def generate_bid_summary(self, result: Dict[str, Any]) -> str:
        """Generate a text summary of the bid requirements."""
        project = result.get('project_info', {})
        schedule = result.get('bid_schedule', {})
        scope = result.get('scope', {})
        requirements = result.get('requirements', {})
        
        summary = f"""
PROJECT: {project.get('project_name', 'N/A')}
PROJECT NO: {project.get('project_number', 'N/A')}
OWNER: {project.get('owner', 'N/A')}
LOCATION: {project.get('location', 'N/A')}
TYPE: {project.get('project_type', 'N/A')}

BID DUE: {schedule.get('bid_date', 'N/A')} at {schedule.get('bid_time', 'N/A')}

SCOPE: {scope.get('description', 'N/A')}
DURATION: {scope.get('duration', 'N/A')}

LINE ITEMS: {len(result.get('line_items', []))} items

KEY REQUIREMENTS:
- Prevailing Wage: {'Yes' if requirements.get('prevailing_wage') else 'No/Not Specified'}
- DBE Goal: {requirements.get('dbe_goal', 'N/A')}
- Bonding: {requirements.get('bonding', 'N/A')}
- Liquidated Damages: {requirements.get('liquidated_damages', 'N/A')}
"""
        return summary.strip()
    
    def get_key_dates(self, result: Dict[str, Any]) -> List[Dict[str, str]]:
        """Extract all key dates from parsed results."""
        dates = result.get('key_dates', [])
        
        # Add bid schedule dates
        schedule = result.get('bid_schedule', {})
        if schedule.get('bid_date'):
            dates.append({'event': 'Bid Due', 'date': f"{schedule.get('bid_date')} {schedule.get('bid_time', '')}"})
        
        pre_bid = schedule.get('pre_bid_meeting', {})
        if pre_bid.get('date'):
            mandatory = " (MANDATORY)" if pre_bid.get('mandatory') else ""
            dates.append({'event': f'Pre-Bid Meeting{mandatory}', 'date': f"{pre_bid.get('date')} {pre_bid.get('time', '')}"})
        
        if schedule.get('question_deadline'):
            dates.append({'event': 'Question Deadline', 'date': schedule.get('question_deadline')})
        
        return dates
