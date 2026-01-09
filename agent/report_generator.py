"""
Report Generator - Generate Word documents and Excel spreadsheets for bid analysis
"""

import io
from typing import Dict, Any, List, Optional
from datetime import datetime

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ReportGenerator:
    """
    Generates bid analysis reports in Word and Excel formats.
    """
    
    # Brand colors
    NAVY = RGBColor(27, 54, 93)
    RED = RGBColor(200, 16, 46)
    GREEN = RGBColor(40, 167, 69)
    ORANGE = RGBColor(230, 126, 34)
    GRAY = RGBColor(108, 117, 125)
    
    # Excel colors
    EXCEL_NAVY = "1B365D"
    EXCEL_RED = "C8102E"
    EXCEL_GREEN = "28A745"
    EXCEL_ORANGE = "E67E22"
    EXCEL_LIGHT = "F5F5F5"
    
    def __init__(self):
        pass
    
    def _set_cell_shading(self, cell, color_hex: str):
        """Set background shading for a table cell."""
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), color_hex)
        cell._tc.get_or_add_tcPr().append(shading)
    
    def generate_bid_analysis_report(self, analysis: Dict[str, Any], project_name: str = "") -> io.BytesIO:
        """Generate a Word document report for bid analysis."""
        doc = Document()
        
        # Set up styles
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        
        # Title
        title = doc.add_heading('BID PROPOSAL ANALYSIS', level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in title.runs:
            run.font.color.rgb = self.NAVY
        
        if project_name:
            subtitle = doc.add_paragraph(project_name)
            subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
            subtitle.runs[0].font.size = Pt(14)
            subtitle.runs[0].font.color.rgb = self.RED
        
        doc.add_paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}")
        doc.add_paragraph()
        
        # Executive Summary
        self._add_section_header(doc, "EXECUTIVE SUMMARY")
        
        overall = analysis.get('overall_assessment', {})
        status = analysis.get('status', {})
        
        summary_para = doc.add_paragraph()
        summary_para.add_run(f"Status: {status.get('status', 'N/A')}\n").bold = True
        summary_para.add_run(f"Competitiveness Score: {overall.get('competitiveness_score', 'N/A')}/10\n")
        summary_para.add_run(f"Confidence Score: {overall.get('confidence_score', 'N/A')}/10\n")
        summary_para.add_run(f"\n{overall.get('summary', '')}")
        
        # Pricing Analysis
        pricing = analysis.get('pricing_analysis', {})
        if pricing:
            self._add_section_header(doc, "PRICING SUMMARY")
            
            pricing_para = doc.add_paragraph()
            pricing_para.add_run(f"Total Bid: ${pricing.get('total_bid', 0):,.2f}\n").bold = True
            pricing_para.add_run(f"Recommended Total: ${pricing.get('recommended_total', 0):,.2f}\n")
            
            variance = pricing.get('variance_pct', 0)
            var_text = f"Variance: {variance:+.1f}%"
            run = pricing_para.add_run(var_text)
            if variance < -5:
                run.font.color.rgb = self.RED
            elif variance > 5:
                run.font.color.rgb = self.ORANGE
        
        # Risks
        risks = analysis.get('risks', [])
        if risks:
            self._add_section_header(doc, "RISKS")
            
            for risk in risks:
                para = doc.add_paragraph(style='List Bullet')
                severity = risk.get('severity', 'medium').upper()
                para.add_run(f"[{severity}] ").bold = True
                para.add_run(risk.get('risk', ''))
                if risk.get('mitigation'):
                    para.add_run(f"\n  Mitigation: {risk.get('mitigation')}")
        
        # Recommendations
        recommendations = analysis.get('prioritized_recommendations', [])
        if recommendations:
            self._add_section_header(doc, "RECOMMENDATIONS")
            
            for i, rec in enumerate(recommendations[:10], 1):
                para = doc.add_paragraph()
                para.add_run(f"{i}. [{rec.get('priority', '')}] ").bold = True
                para.add_run(rec.get('action', ''))
                if rec.get('rationale'):
                    para.add_run(f"\n   {rec.get('rationale')}")
        
        # Bid Strategy
        strategy = analysis.get('bid_strategy', {})
        if strategy:
            self._add_section_header(doc, "BID STRATEGY")
            
            if strategy.get('approach'):
                doc.add_paragraph(strategy.get('approach'))
            
            if strategy.get('items_to_sharpen'):
                doc.add_paragraph().add_run("Items to Sharpen Pricing:").bold = True
                for item in strategy.get('items_to_sharpen', []):
                    doc.add_paragraph(f"  - {item}", style='List Bullet')
            
            if strategy.get('value_engineering_opportunities'):
                doc.add_paragraph().add_run("Value Engineering Opportunities:").bold = True
                for item in strategy.get('value_engineering_opportunities', []):
                    doc.add_paragraph(f"  - {item}", style='List Bullet')
        
        # Footer
        doc.add_paragraph()
        footer = doc.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer.add_run("Generated by Bid Proposal Agent - Abonmarche")
        run.font.size = Pt(9)
        run.font.color.rgb = self.GRAY
        
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        return buffer
    
    def _add_section_header(self, doc, text: str):
        """Add a styled section header."""
        heading = doc.add_heading(text, level=2)
        for run in heading.runs:
            run.font.color.rgb = self.NAVY
    
    def generate_bid_excel(
        self,
        items: List[Dict[str, Any]],
        project_name: str = "",
        summary: Optional[Dict[str, Any]] = None
    ) -> io.BytesIO:
        """Generate an Excel bid estimate spreadsheet."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bid Estimate"
        
        # Styling
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color=self.EXCEL_NAVY, end_color=self.EXCEL_NAVY, fill_type="solid")
        currency_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Title
        ws.merge_cells('A1:K1')
        ws['A1'] = f"BID ESTIMATE - {project_name}" if project_name else "BID ESTIMATE"
        ws['A1'].font = Font(bold=True, size=14, color=self.EXCEL_NAVY)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        ws['A2'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ws['A2'].font = Font(size=9, color=self.EXCEL_ORANGE)
        
        # Headers
        headers = [
            'Item No.', 'Description', 'Qty', 'Unit',
            'Material', 'Labor', 'Equipment', 'OH&P',
            'Unit Price', 'Total', 'Notes'
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        
        # Data rows
        row_num = 5
        subtotal = 0
        
        for item in items:
            ws.cell(row=row_num, column=1, value=item.get('item_number', '')).border = thin_border
            ws.cell(row=row_num, column=2, value=item.get('description', '')).border = thin_border
            
            qty_cell = ws.cell(row=row_num, column=3, value=item.get('quantity', 0))
            qty_cell.border = thin_border
            qty_cell.number_format = '#,##0.00'
            
            ws.cell(row=row_num, column=4, value=item.get('unit', '')).border = thin_border
            
            # Cost breakdown
            mat_cost = item.get('material', {}).get('cost', 0) if isinstance(item.get('material'), dict) else item.get('material', 0)
            lab_cost = item.get('labor', {}).get('cost', 0) if isinstance(item.get('labor'), dict) else item.get('labor', 0)
            equip_cost = item.get('equipment', {}).get('cost', 0) if isinstance(item.get('equipment'), dict) else item.get('equipment', 0)
            ohp = item.get('overhead_profit', 0)
            
            mat_cell = ws.cell(row=row_num, column=5, value=mat_cost)
            mat_cell.border = thin_border
            mat_cell.number_format = currency_format
            
            lab_cell = ws.cell(row=row_num, column=6, value=lab_cost)
            lab_cell.border = thin_border
            lab_cell.number_format = currency_format
            
            equip_cell = ws.cell(row=row_num, column=7, value=equip_cost)
            equip_cell.border = thin_border
            equip_cell.number_format = currency_format
            
            ohp_cell = ws.cell(row=row_num, column=8, value=ohp)
            ohp_cell.border = thin_border
            ohp_cell.number_format = currency_format
            
            unit_price = item.get('unit_price', 0)
            up_cell = ws.cell(row=row_num, column=9, value=unit_price)
            up_cell.border = thin_border
            up_cell.number_format = currency_format
            
            total = item.get('total_price', 0)
            total_cell = ws.cell(row=row_num, column=10, value=total)
            total_cell.border = thin_border
            total_cell.number_format = currency_format
            total_cell.font = Font(bold=True)
            
            subtotal += total
            
            ws.cell(row=row_num, column=11, value=item.get('notes', '')).border = thin_border
            
            row_num += 1
        
        # Summary
        row_num += 1
        ws.cell(row=row_num, column=9, value="SUBTOTAL:").font = Font(bold=True)
        subtotal_cell = ws.cell(row=row_num, column=10, value=subtotal)
        subtotal_cell.font = Font(bold=True)
        subtotal_cell.number_format = currency_format
        
        if summary:
            cont_pct = summary.get('contingency_pct', 5)
            cont_amt = summary.get('contingency_amt', subtotal * cont_pct / 100)
            
            row_num += 1
            ws.cell(row=row_num, column=9, value=f"Contingency ({cont_pct}%):").font = Font(bold=True)
            cont_cell = ws.cell(row=row_num, column=10, value=cont_amt)
            cont_cell.number_format = currency_format
            
            row_num += 1
            total_bid = summary.get('total_bid', subtotal + cont_amt)
            ws.cell(row=row_num, column=9, value="TOTAL BID:").font = Font(bold=True, color=self.EXCEL_NAVY)
            total_cell = ws.cell(row=row_num, column=10, value=total_bid)
            total_cell.font = Font(bold=True, size=12, color=self.EXCEL_NAVY)
            total_cell.number_format = currency_format
        
        # Adjust column widths
        widths = [10, 40, 10, 8, 12, 12, 12, 12, 12, 14, 30]
        for i, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        return buffer
    
    def generate_html_report(self, analysis: Dict[str, Any], project_name: str = "") -> str:
        """Generate an HTML report for display in the web UI."""
        status = analysis.get('status', {})
        overall = analysis.get('overall_assessment', {})
        pricing = analysis.get('pricing_analysis', {})
        
        status_color = '#28a745' if status.get('color') == 'green' else '#dc3545' if status.get('color') == 'red' else '#ffc107'
        
        html = f'''
<div class="bid-analysis-report">
    <div class="report-header" style="background: #1B365D; color: white; padding: 20px; border-radius: 8px 8px 0 0;">
        <h1 style="margin: 0;">Bid Analysis Report</h1>
        <p style="margin: 5px 0 0 0; opacity: 0.9;">{project_name or 'Project Analysis'}</p>
    </div>
    
    <div class="status-banner" style="background: {status_color}20; border-left: 4px solid {status_color}; padding: 15px; margin: 20px 0;">
        <h2 style="color: {status_color}; margin: 0;">{status.get('status', 'REVIEW')}</h2>
        <p style="margin: 10px 0 0 0;">{status.get('message', '')}</p>
        <p style="margin: 5px 0 0 0;"><strong>Recommendation:</strong> {analysis.get('final_recommendation', 'revise').upper()}</p>
    </div>
    
    <div class="scores-grid" style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin: 20px 0;">
        <div style="background: #e3f2fd; padding: 20px; border-radius: 8px; text-align: center;">
            <div style="font-size: 2rem; font-weight: bold; color: #1B365D;">{overall.get('competitiveness_score', 'N/A')}/10</div>
            <div style="font-size: 0.85rem; color: #666;">Competitiveness</div>
        </div>
        <div style="background: #e8f5e9; padding: 20px; border-radius: 8px; text-align: center;">
            <div style="font-size: 2rem; font-weight: bold; color: #28a745;">{overall.get('confidence_score', 'N/A')}/10</div>
            <div style="font-size: 0.85rem; color: #666;">Confidence</div>
        </div>
        <div style="background: #fff3e0; padding: 20px; border-radius: 8px; text-align: center;">
            <div style="font-size: 2rem; font-weight: bold; color: #E67E22;">${pricing.get('total_bid', 0):,.0f}</div>
            <div style="font-size: 0.85rem; color: #666;">Total Bid</div>
        </div>
    </div>
    
    <div class="summary" style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <h3 style="color: #1B365D; margin-top: 0;">Summary</h3>
        <p>{overall.get('summary', 'Analysis in progress...')}</p>
    </div>
'''
        
        # Risks
        risks = analysis.get('risks', [])
        if risks:
            html += '''
    <div class="risks" style="margin: 20px 0;">
        <h3 style="color: #1B365D; border-bottom: 2px solid #1B365D; padding-bottom: 8px;">Risks</h3>
        <ul style="list-style: none; padding: 0;">
'''
            for risk in risks[:5]:
                severity = risk.get('severity', 'medium')
                color = '#dc3545' if severity == 'high' else '#ffc107' if severity == 'medium' else '#6c757d'
                html += f'''
            <li style="padding: 10px; background: {color}15; border-left: 4px solid {color}; margin-bottom: 8px; border-radius: 0 4px 4px 0;">
                <strong style="color: {color};">[{severity.upper()}]</strong> {risk.get('risk', '')}
                {f"<br><small style='color: #666;'>Mitigation: {risk.get('mitigation', '')}</small>" if risk.get('mitigation') else ""}
            </li>
'''
            html += '''
        </ul>
    </div>
'''
        
        # Recommendations
        recommendations = analysis.get('prioritized_recommendations', [])
        if recommendations:
            html += '''
    <div class="recommendations" style="margin: 20px 0;">
        <h3 style="color: #1B365D; border-bottom: 2px solid #1B365D; padding-bottom: 8px;">Recommendations</h3>
        <ol style="padding-left: 20px;">
'''
            for rec in recommendations[:8]:
                priority = rec.get('priority', 'MEDIUM')
                color = '#dc3545' if priority == 'CRITICAL' else '#E67E22' if priority == 'HIGH' else '#1B365D'
                html += f'''
            <li style="padding: 8px 0;">
                <span style="color: {color}; font-weight: bold;">[{priority}]</span> {rec.get('action', '')}
                {f"<br><small style='color: #666;'>{rec.get('rationale', '')}</small>" if rec.get('rationale') else ""}
            </li>
'''
            html += '''
        </ol>
    </div>
'''
        
        # Bid Strategy
        strategy = analysis.get('bid_strategy', {})
        if strategy and strategy.get('approach'):
            html += f'''
    <div class="strategy" style="margin: 20px 0;">
        <h3 style="color: #1B365D; border-bottom: 2px solid #1B365D; padding-bottom: 8px;">Bid Strategy</h3>
        <p>{strategy.get('approach', '')}</p>
    </div>
'''
        
        html += '''
    <div class="footer" style="text-align: center; color: #999; font-size: 12px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd;">
        Generated by Bid Proposal Agent - Abonmarche
    </div>
</div>
'''
        
        return html
