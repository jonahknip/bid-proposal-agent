"""
Quantity Calculator - AI-powered quantity extraction from civil engineering plans
Extracts earthwork, paving, utilities, survey, and structure quantities using GPT-4 Vision
"""

import os
import io
import base64
import json
import re
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field, asdict
from decimal import Decimal, ROUND_HALF_UP

import fitz  # PyMuPDF
from PIL import Image
from openai import OpenAI


@dataclass
class QuantityItem:
    """A single quantity line item"""
    item_number: str = ""
    description: str = ""
    quantity: float = 0.0
    unit: str = ""
    category: str = ""  # earthwork, paving, utilities, survey, structures, misc
    subcategory: str = ""  # e.g., excavation, fill, HMA, concrete, storm, sanitary
    sheet_reference: str = ""
    notes: str = ""
    confidence: float = 0.0


@dataclass 
class QuantitySummary:
    """Summary of quantities by category"""
    earthwork: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    paving: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    utilities: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    survey: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    structures: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    erosion_control: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    traffic_control: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    landscaping: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    misc: Dict[str, List[QuantityItem]] = field(default_factory=dict)
    total_items: int = 0


class QuantityCalculator:
    """
    AI-powered quantity calculator for civil engineering plans.
    Uses GPT-4 Vision to extract and calculate quantities from plan sheets.
    """
    
    # Standard units by category
    STANDARD_UNITS = {
        'earthwork': {
            'excavation': 'CY',
            'fill': 'CY',
            'topsoil': 'CY',
            'subgrade': 'SY',
            'grading': 'SY',
            'embankment': 'CY',
            'borrow': 'CY',
            'unsuitable': 'CY',
        },
        'paving': {
            'hma': 'TON',  # Hot Mix Asphalt
            'asphalt': 'TON',
            'concrete': 'SY',
            'base': 'TON',
            'aggregate': 'TON',
            'subbase': 'CY',
            'curb': 'LF',
            'sidewalk': 'SF',
            'driveway': 'SY',
        },
        'utilities': {
            'storm_pipe': 'LF',
            'sanitary_pipe': 'LF',
            'water_main': 'LF',
            'manhole': 'EA',
            'catch_basin': 'EA',
            'inlet': 'EA',
            'hydrant': 'EA',
            'valve': 'EA',
            'service': 'EA',
            'cleanout': 'EA',
        },
        'survey': {
            'boundary': 'LS',
            'topographic': 'AC',
            'construction_staking': 'LS',
            'control': 'EA',
            'right_of_way': 'LF',
            'easement': 'AC',
        },
        'structures': {
            'retaining_wall': 'SF',
            'headwall': 'EA',
            'endwall': 'EA',
            'culvert': 'LF',
            'bridge': 'LS',
            'box_culvert': 'LF',
            'guardrail': 'LF',
        },
        'erosion_control': {
            'silt_fence': 'LF',
            'inlet_protection': 'EA',
            'stabilized_construction': 'SY',
            'seeding': 'AC',
            'mulch': 'AC',
            'erosion_blanket': 'SY',
            'check_dam': 'EA',
        },
        'traffic_control': {
            'signs': 'EA',
            'pavement_marking': 'LF',
            'striping': 'LF',
            'signal': 'LS',
            'barricade': 'EA',
            'drums': 'EA',
        },
        'landscaping': {
            'trees': 'EA',
            'shrubs': 'EA',
            'sod': 'SY',
            'seed': 'LB',
            'mulch': 'CY',
            'irrigation': 'LS',
        }
    }
    
    # Quantity extraction prompt for GPT-4 Vision
    QUANTITY_EXTRACTION_PROMPT = """You are an expert civil engineering quantity takeoff specialist. Analyze this plan sheet and extract ALL quantities visible.

IMPORTANT: Look for quantity tables, schedules, callouts, dimensions, and any numerical data that represents construction quantities.

For each quantity found, provide:
1. Item description (match standard pay item descriptions when possible)
2. Numeric quantity value
3. Unit of measure (CY, LF, SY, SF, TON, EA, AC, LS, etc.)
4. Category: earthwork, paving, utilities, survey, structures, erosion_control, traffic_control, landscaping, or misc
5. Subcategory (e.g., for utilities: storm_pipe, sanitary_pipe, water_main, manhole, etc.)
6. Where on the sheet you found this (e.g., "quantity table", "profile note", "detail callout")

QUANTITY CATEGORIES TO LOOK FOR:

EARTHWORK:
- Excavation, Fill, Embankment (CY)
- Topsoil removal/replacement (CY)
- Subgrade preparation (SY)
- Unsuitable material removal (CY)

PAVING:
- Hot Mix Asphalt / HMA (TON or SY)
- Concrete pavement (SY)
- Aggregate base (TON or CY)
- Curb & gutter (LF)
- Sidewalk (SF or SY)
- Driveways (SY)

UTILITIES:
- Storm sewer pipe by diameter (LF)
- Sanitary sewer pipe by diameter (LF)
- Water main by diameter (LF)
- Manholes, catch basins, inlets (EA)
- Fire hydrants, valves (EA)
- Services and connections (EA)

STRUCTURES:
- Retaining walls (SF)
- Culverts (LF)
- Headwalls, endwalls (EA)
- Guardrail (LF)

EROSION CONTROL:
- Silt fence (LF)
- Inlet protection (EA)
- Seeding/mulching (AC or SY)
- Erosion blanket (SY)

TRAFFIC CONTROL:
- Signs (EA)
- Pavement markings (LF or SF)
- Temporary traffic control items

Return a JSON object:
{
    "sheet_info": {
        "sheet_number": "extracted sheet number",
        "sheet_title": "extracted sheet title",
        "sheet_type": "grading/utility/paving/detail/etc"
    },
    "quantities": [
        {
            "item_number": "pay item number if visible",
            "description": "item description",
            "quantity": numeric_value,
            "unit": "unit of measure",
            "category": "category",
            "subcategory": "subcategory",
            "location_on_sheet": "where found",
            "confidence": 0.0 to 1.0
        }
    ],
    "calculations_visible": true/false,
    "notes": "any relevant notes about quantities on this sheet"
}

Only return the JSON object, no other text."""

    def __init__(self, api_key: Optional[str] = None):
        """Initialize the quantity calculator with OpenAI API key."""
        self.api_key = api_key or os.environ.get('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OpenAI API key is required")
        self.client = OpenAI(api_key=self.api_key)
        self.model = "gpt-4o"
        
    def extract_sheets_from_pdf(self, pdf_path: str, dpi: int = 150) -> List[Dict[str, Any]]:
        """Extract each page from a PDF as an image for analysis."""
        sheets = []
        doc = fitz.open(pdf_path)
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Render page to image
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            # Resize if too large (max 2048px for API efficiency)
            max_dim = 2048
            if max(img.size) > max_dim:
                ratio = max_dim / max(img.size)
                new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
                img = img.resize(new_size, Image.LANCZOS)
            
            # Convert to base64
            buffer = io.BytesIO()
            img.save(buffer, format="PNG")
            base64_image = base64.b64encode(buffer.getvalue()).decode('utf-8')
            
            # Also extract text
            text_content = page.get_text()
            
            sheets.append({
                'page_num': page_num + 1,
                'pdf_name': os.path.basename(pdf_path),
                'image_base64': base64_image,
                'text_content': text_content,
                'width': img.size[0],
                'height': img.size[1]
            })
        
        doc.close()
        return sheets
    
    def extract_quantities_from_sheet(self, sheet: Dict[str, Any]) -> Dict[str, Any]:
        """Use GPT-4 Vision to extract quantities from a single plan sheet."""
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": self.QUANTITY_EXTRACTION_PROMPT},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{sheet['image_base64']}",
                                "detail": "high"
                            }
                        }
                    ]
                }
            ],
            max_tokens=4000,
            temperature=0.2  # Low temperature for more consistent extraction
        )
        
        try:
            content = response.choices[0].message.content
            # Clean up any markdown code blocks
            if content.startswith('```'):
                content = content.split('```')[1]
                if content.startswith('json'):
                    content = content[4:]
            if content.endswith('```'):
                content = content[:-3]
            
            result = json.loads(content.strip())
            result['page_num'] = sheet['page_num']
            result['pdf_name'] = sheet['pdf_name']
            return result
            
        except json.JSONDecodeError:
            # Try to extract JSON from response
            content = response.choices[0].message.content
            start = content.find('{')
            end = content.rfind('}') + 1
            if start != -1 and end > start:
                try:
                    result = json.loads(content[start:end])
                    result['page_num'] = sheet['page_num']
                    result['pdf_name'] = sheet['pdf_name']
                    return result
                except:
                    pass
            
            return {
                'page_num': sheet['page_num'],
                'pdf_name': sheet['pdf_name'],
                'sheet_info': {'sheet_number': str(sheet['page_num']), 'sheet_title': 'Unknown', 'sheet_type': 'unknown'},
                'quantities': [],
                'calculations_visible': False,
                'notes': 'Unable to extract quantities from this sheet',
                'error': True
            }
    
    def extract_quantities_from_pdf(self, pdf_path: str, max_sheets: Optional[int] = None) -> Dict[str, Any]:
        """
        Extract quantities from all sheets in a PDF.
        
        Args:
            pdf_path: Path to the PDF file
            max_sheets: Maximum number of sheets to process (None for all)
            
        Returns:
            Dictionary with all extracted quantities organized by category
        """
        sheets = self.extract_sheets_from_pdf(pdf_path)
        
        if max_sheets:
            sheets = sheets[:max_sheets]
        
        all_quantities = []
        sheet_results = []
        
        for sheet in sheets:
            result = self.extract_quantities_from_sheet(sheet)
            sheet_results.append(result)
            
            for qty in result.get('quantities', []):
                qty['sheet_reference'] = f"{result['pdf_name']} - Sheet {result['page_num']}"
                all_quantities.append(qty)
        
        # Organize by category
        summary = self._organize_quantities(all_quantities)
        
        return {
            'pdf_name': os.path.basename(pdf_path),
            'total_sheets': len(sheets),
            'sheets_processed': len(sheet_results),
            'sheet_results': sheet_results,
            'quantity_summary': summary,
            'all_quantities': all_quantities
        }
    
    def extract_quantities_from_multiple_pdfs(self, pdf_paths: List[str], max_sheets_per_pdf: Optional[int] = None) -> Dict[str, Any]:
        """Extract quantities from multiple PDF files."""
        all_results = []
        combined_quantities = []
        
        for pdf_path in pdf_paths:
            result = self.extract_quantities_from_pdf(pdf_path, max_sheets_per_pdf)
            all_results.append(result)
            combined_quantities.extend(result.get('all_quantities', []))
        
        # Combine and organize all quantities
        summary = self._organize_quantities(combined_quantities)
        
        return {
            'pdf_files': [os.path.basename(p) for p in pdf_paths],
            'total_pdfs': len(pdf_paths),
            'individual_results': all_results,
            'combined_summary': summary,
            'all_quantities': combined_quantities
        }
    
    def _organize_quantities(self, quantities: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Organize quantities by category and subcategory."""
        organized = {
            'earthwork': {},
            'paving': {},
            'utilities': {},
            'survey': {},
            'structures': {},
            'erosion_control': {},
            'traffic_control': {},
            'landscaping': {},
            'misc': {}
        }
        
        for qty in quantities:
            category = qty.get('category', 'misc').lower().replace(' ', '_')
            subcategory = qty.get('subcategory', 'general').lower().replace(' ', '_')
            
            if category not in organized:
                category = 'misc'
            
            if subcategory not in organized[category]:
                organized[category][subcategory] = []
            
            organized[category][subcategory].append(qty)
        
        # Calculate totals for each category
        totals = {}
        for category, subcats in organized.items():
            totals[category] = {
                'item_count': sum(len(items) for items in subcats.values()),
                'subcategories': list(subcats.keys())
            }
        
        return {
            'by_category': organized,
            'totals': totals,
            'total_items': sum(t['item_count'] for t in totals.values())
        }
    
    def aggregate_quantities(self, quantities: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Aggregate similar quantities (same description + unit) into single line items.
        Useful for combining quantities from multiple sheets.
        """
        aggregated = {}
        
        for qty in quantities:
            key = f"{qty.get('description', '').lower().strip()}|{qty.get('unit', '').upper()}"
            
            if key not in aggregated:
                aggregated[key] = {
                    'description': qty.get('description', ''),
                    'quantity': 0.0,
                    'unit': qty.get('unit', ''),
                    'category': qty.get('category', 'misc'),
                    'subcategory': qty.get('subcategory', ''),
                    'sources': []
                }
            
            try:
                aggregated[key]['quantity'] += float(qty.get('quantity', 0))
            except (ValueError, TypeError):
                pass
            
            aggregated[key]['sources'].append(qty.get('sheet_reference', 'Unknown'))
        
        # Convert to list and round quantities
        result = []
        for item in aggregated.values():
            item['quantity'] = round(item['quantity'], 2)
            item['source_count'] = len(item['sources'])
            result.append(item)
        
        # Sort by category, then description
        result.sort(key=lambda x: (x['category'], x['description']))
        
        return result
    
    def calculate_earthwork(self, cut_areas: List[Tuple[float, float]], fill_areas: List[Tuple[float, float]], 
                           station_interval: float = 50.0) -> Dict[str, float]:
        """
        Calculate earthwork volumes using average end area method.
        
        Args:
            cut_areas: List of (station, area) tuples for cut sections
            fill_areas: List of (station, area) tuples for fill sections
            station_interval: Default interval between stations (feet)
            
        Returns:
            Dictionary with cut and fill volumes in cubic yards
        """
        def avg_end_area(areas: List[Tuple[float, float]]) -> float:
            if len(areas) < 2:
                return 0.0
            
            areas.sort(key=lambda x: x[0])  # Sort by station
            total_volume = 0.0
            
            for i in range(len(areas) - 1):
                sta1, area1 = areas[i]
                sta2, area2 = areas[i + 1]
                distance = sta2 - sta1
                avg_area = (area1 + area2) / 2
                volume = avg_area * distance
                total_volume += volume
            
            # Convert to cubic yards (from cubic feet)
            return total_volume / 27
        
        return {
            'cut_cy': round(avg_end_area(cut_areas), 1),
            'fill_cy': round(avg_end_area(fill_areas), 1),
            'net_cy': round(avg_end_area(fill_areas) - avg_end_area(cut_areas), 1)
        }
    
    def calculate_paving_area(self, length: float, width: float, depth_inches: float = 0) -> Dict[str, float]:
        """
        Calculate paving quantities.
        
        Args:
            length: Length in feet
            width: Width in feet
            depth_inches: Depth in inches (for tonnage calc)
            
        Returns:
            Area in SF, SY, and estimated tonnage if depth provided
        """
        area_sf = length * width
        area_sy = area_sf / 9
        
        result = {
            'area_sf': round(area_sf, 1),
            'area_sy': round(area_sy, 1)
        }
        
        if depth_inches > 0:
            # HMA: approximately 110 lbs/SY per inch of thickness
            tons = (area_sy * depth_inches * 110) / 2000
            result['estimated_tons'] = round(tons, 1)
        
        return result
    
    def calculate_pipe_quantity(self, segments: List[Dict[str, Any]]) -> Dict[str, float]:
        """
        Calculate pipe quantities by diameter.
        
        Args:
            segments: List of dicts with 'diameter' (inches) and 'length' (feet)
            
        Returns:
            Dictionary with total length by diameter
        """
        by_diameter = {}
        
        for seg in segments:
            dia = seg.get('diameter', 0)
            length = seg.get('length', 0)
            
            key = f"{dia}\" Pipe"
            if key not in by_diameter:
                by_diameter[key] = 0
            by_diameter[key] += length
        
        return {k: round(v, 1) for k, v in by_diameter.items()}
    
    def verify_quantities(self, proposed: List[Dict[str, Any]], extracted: List[Dict[str, Any]], 
                         tolerance: float = 0.10) -> Dict[str, Any]:
        """
        Compare proposed quantities against extracted quantities.
        
        Args:
            proposed: List of proposed quantities from bid
            extracted: List of extracted quantities from plans
            tolerance: Acceptable variance (0.10 = 10%)
            
        Returns:
            Verification results with matches, discrepancies, and missing items
        """
        results = {
            'matches': [],
            'discrepancies': [],
            'missing_in_proposal': [],
            'extra_in_proposal': [],
            'summary': {}
        }
        
        # Normalize descriptions for matching
        def normalize(desc):
            return re.sub(r'[^a-z0-9]', '', desc.lower())
        
        proposed_lookup = {normalize(p.get('description', '')): p for p in proposed}
        extracted_lookup = {normalize(e.get('description', '')): e for e in extracted}
        
        # Check each extracted quantity
        for key, ext in extracted_lookup.items():
            if key in proposed_lookup:
                prop = proposed_lookup[key]
                prop_qty = float(prop.get('quantity', 0))
                ext_qty = float(ext.get('quantity', 0))
                
                if ext_qty > 0:
                    variance = abs(prop_qty - ext_qty) / ext_qty
                else:
                    variance = 1.0 if prop_qty > 0 else 0.0
                
                if variance <= tolerance:
                    results['matches'].append({
                        'description': ext.get('description'),
                        'proposed_qty': prop_qty,
                        'extracted_qty': ext_qty,
                        'unit': ext.get('unit'),
                        'variance_pct': round(variance * 100, 1)
                    })
                else:
                    results['discrepancies'].append({
                        'description': ext.get('description'),
                        'proposed_qty': prop_qty,
                        'extracted_qty': ext_qty,
                        'unit': ext.get('unit'),
                        'variance_pct': round(variance * 100, 1),
                        'difference': round(prop_qty - ext_qty, 2)
                    })
            else:
                results['missing_in_proposal'].append(ext)
        
        # Check for extra items in proposal
        for key, prop in proposed_lookup.items():
            if key not in extracted_lookup:
                results['extra_in_proposal'].append(prop)
        
        # Summary
        results['summary'] = {
            'total_compared': len(extracted),
            'matches': len(results['matches']),
            'discrepancies': len(results['discrepancies']),
            'missing_in_proposal': len(results['missing_in_proposal']),
            'extra_in_proposal': len(results['extra_in_proposal']),
            'match_rate': round(len(results['matches']) / max(len(extracted), 1) * 100, 1)
        }
        
        return results
    
    def export_to_dict(self, quantities: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Convert quantities to a clean dictionary format for export."""
        return [{
            'Item No.': qty.get('item_number', ''),
            'Description': qty.get('description', ''),
            'Quantity': qty.get('quantity', 0),
            'Unit': qty.get('unit', ''),
            'Category': qty.get('category', ''),
            'Subcategory': qty.get('subcategory', ''),
            'Sheet Reference': qty.get('sheet_reference', ''),
            'Notes': qty.get('notes', ''),
            'Confidence': qty.get('confidence', 0)
        } for qty in quantities]
