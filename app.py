"""
Bid Proposal Agent - Web Application
A tool for analyzing and reviewing civil engineering bid proposals
"""

import os
import tempfile
import logging
import traceback
import json
import uuid
from datetime import datetime
from pathlib import Path
from io import BytesIO

from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename

from agent.quantity_calculator import QuantityCalculator
from agent.proposal_parser import ProposalParser
from agent.bid_analyzer import BidAnalyzer
from agent.report_generator import ReportGenerator

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max upload
app.secret_key = os.environ.get('SECRET_KEY', 'bid-proposal-agent-secret-key-change-in-prod')

# In-memory storage for session data
session_data = {}

# Allowed extensions
ALLOWED_EXTENSIONS = {'.pdf', '.xlsx', '.xls', '.xlsm'}


def allowed_file(filename: str) -> bool:
    """Check if file has allowed extension"""
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def get_session_id():
    """Get or create session ID"""
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
    return session['session_id']


def get_session_data():
    """Get session data for current user"""
    sid = get_session_id()
    if sid not in session_data:
        session_data[sid] = {
            'proposal_requirements': None,
            'bid_proposal': None,
            'plan_quantities': None,
            'analysis_results': None,
            'history': []
        }
    return session_data[sid]


@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'service': 'bid-proposal-agent'})


@app.route('/api/parse-proposal', methods=['POST'])
def parse_proposal_documents():
    """
    Parse proposal/RFP documents to extract requirements and line items.
    Accepts PDF and Excel files.
    """
    temp_files = []
    
    try:
        if 'files' not in request.files:
            return jsonify({
                'success': False,
                'error': 'No files uploaded'
            }), 400
        
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            return jsonify({
                'success': False,
                'error': 'No files selected'
            }), 400
        
        # Save uploaded files
        file_paths = []
        for file in files:
            if file.filename and allowed_file(file.filename):
                temp_dir = tempfile.mkdtemp()
                filename = secure_filename(file.filename)
                temp_path = os.path.join(temp_dir, filename)
                file.save(temp_path)
                file_paths.append(temp_path)
                temp_files.append(temp_path)
                logger.info(f"Saved proposal document: {filename}")
            else:
                logger.warning(f"Skipped invalid file: {file.filename}")
        
        if not file_paths:
            return jsonify({
                'success': False,
                'error': 'No valid files uploaded. Supported formats: PDF, Excel (.xlsx, .xls)'
            }), 400
        
        # Parse documents
        parser = ProposalParser()
        
        if len(file_paths) == 1:
            result = parser.parse_bid_document(file_paths[0])
        else:
            result = parser.parse_multiple_documents(file_paths)
        
        # Store in session
        data = get_session_data()
        data['proposal_requirements'] = result
        
        # Extract clean line items for response
        line_items = parser.extract_line_items_table(result)
        requirements_checklist = parser.generate_requirements_checklist(result)
        
        return jsonify({
            'success': True,
            'project_info': result.get('project_info', {}),
            'bid_schedule': result.get('bid_schedule', {}),
            'scope_summary': result.get('scope_summary', ''),
            'line_items': line_items,
            'line_item_count': len(line_items),
            'requirements_checklist': requirements_checklist,
            'files_processed': len(file_paths)
        })
        
    except Exception as e:
        logger.error(f"Parse proposal error: {e}\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500
    
    finally:
        # Clean up temp files
        for path in temp_files:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    os.rmdir(os.path.dirname(path))
            except Exception:
                pass


@app.route('/api/parse-bid', methods=['POST'])
def parse_bid_proposal():
    """
    Parse the working bid proposal document.
    Accepts PDF and Excel files.
    """
    temp_files = []
    
    try:
        if 'files' not in request.files:
            return jsonify({
                'success': False,
                'error': 'No files uploaded'
            }), 400
        
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            return jsonify({
                'success': False,
                'error': 'No files selected'
            }), 400
        
        # Save uploaded files
        file_paths = []
        for file in files:
            if file.filename and allowed_file(file.filename):
                temp_dir = tempfile.mkdtemp()
                filename = secure_filename(file.filename)
                temp_path = os.path.join(temp_dir, filename)
                file.save(temp_path)
                file_paths.append(temp_path)
                temp_files.append(temp_path)
                logger.info(f"Saved bid document: {filename}")
        
        if not file_paths:
            return jsonify({
                'success': False,
                'error': 'No valid files uploaded'
            }), 400
        
        # Parse documents
        parser = ProposalParser()
        
        if len(file_paths) == 1:
            result = parser.parse_bid_document(file_paths[0])
        else:
            result = parser.parse_multiple_documents(file_paths)
        
        # Store in session
        data = get_session_data()
        data['bid_proposal'] = result
        
        # Extract clean line items
        line_items = parser.extract_line_items_table(result)
        
        return jsonify({
            'success': True,
            'line_items': line_items,
            'line_item_count': len(line_items),
            'totals': result.get('totals', {}),
            'files_processed': len(file_paths)
        })
        
    except Exception as e:
        logger.error(f"Parse bid error: {e}\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500
    
    finally:
        for path in temp_files:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    os.rmdir(os.path.dirname(path))
            except Exception:
                pass


@app.route('/api/extract-quantities', methods=['POST'])
def extract_quantities():
    """
    Extract quantities from plan sheets using AI vision.
    Accepts PDF files only.
    """
    temp_files = []
    
    try:
        if 'files' not in request.files:
            return jsonify({
                'success': False,
                'error': 'No files uploaded'
            }), 400
        
        files = request.files.getlist('files')
        pdf_paths = []
        
        for file in files:
            if file.filename and file.filename.lower().endswith('.pdf'):
                temp_dir = tempfile.mkdtemp()
                filename = secure_filename(file.filename)
                temp_path = os.path.join(temp_dir, filename)
                file.save(temp_path)
                pdf_paths.append(temp_path)
                temp_files.append(temp_path)
                logger.info(f"Saved plan file: {filename}")
        
        if not pdf_paths:
            return jsonify({
                'success': False,
                'error': 'No PDF files uploaded. Plan sheets must be PDF format.'
            }), 400
        
        # Get max sheets parameter
        max_sheets = request.form.get('max_sheets')
        max_sheets = int(max_sheets) if max_sheets else None
        
        # Extract quantities
        calculator = QuantityCalculator()
        
        if len(pdf_paths) == 1:
            result = calculator.extract_quantities_from_pdf(pdf_paths[0], max_sheets)
        else:
            result = calculator.extract_quantities_from_multiple_pdfs(pdf_paths, max_sheets)
        
        # Aggregate quantities
        all_quantities = result.get('all_quantities', [])
        aggregated = calculator.aggregate_quantities(all_quantities)
        
        # Store in session
        data = get_session_data()
        data['plan_quantities'] = {
            'raw': result,
            'aggregated': aggregated
        }
        
        return jsonify({
            'success': True,
            'quantities': aggregated,
            'quantity_count': len(aggregated),
            'sheets_processed': result.get('sheets_processed', 0),
            'summary': result.get('quantity_summary', {})
        })
        
    except Exception as e:
        logger.error(f"Quantity extraction error: {e}\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500
    
    finally:
        for path in temp_files:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    os.rmdir(os.path.dirname(path))
            except Exception:
                pass


@app.route('/api/analyze', methods=['POST'])
def analyze_bid():
    """
    Analyze bid proposal against requirements and plan quantities.
    """
    try:
        data = get_session_data()
        
        # Check if we have requirements
        if not data.get('proposal_requirements'):
            return jsonify({
                'success': False,
                'error': 'No proposal documents uploaded. Please upload RFP/bid documents first.'
            }), 400
        
        # Check if we have a bid
        if not data.get('bid_proposal'):
            return jsonify({
                'success': False,
                'error': 'No bid proposal uploaded. Please upload your working bid.'
            }), 400
        
        # Run analysis
        analyzer = BidAnalyzer()
        
        plan_quantities = None
        if data.get('plan_quantities'):
            plan_quantities = data['plan_quantities'].get('raw')
        
        analysis = analyzer.analyze_bid(
            data['proposal_requirements'],
            data['bid_proposal'],
            plan_quantities
        )
        
        # Get recommendations
        recommendations = analyzer.generate_recommendations(analysis)
        analysis['prioritized_recommendations'] = recommendations
        
        # Get status
        status = analyzer.get_bid_status(analysis)
        analysis['status'] = status
        
        # Store results
        data['analysis_results'] = analysis
        
        # Add to history
        history_entry = {
            'id': str(uuid.uuid4()),
            'timestamp': datetime.now().isoformat(),
            'project_name': data['proposal_requirements'].get('project_info', {}).get('project_name', 'Unknown'),
            'status': status,
            'summary': analysis.get('summary', {})
        }
        data['history'].insert(0, history_entry)
        data['history'] = data['history'][:20]  # Keep last 20
        
        # Generate HTML report
        report_gen = ReportGenerator()
        project_name = data['proposal_requirements'].get('project_info', {}).get('project_name', '')
        html_report = report_gen.generate_html_report(analysis, project_name)
        
        return jsonify({
            'success': True,
            'status': status,
            'summary': analysis.get('summary', {}),
            'critical_issues': analysis.get('critical_issues', []),
            'warnings': analysis.get('warnings', []),
            'recommendations': recommendations[:10],
            'html_report': html_report,
            'analysis_id': history_entry['id']
        })
        
    except Exception as e:
        logger.error(f"Analysis error: {e}\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/api/compare-quantities', methods=['POST'])
def compare_quantities():
    """
    Compare bid quantities against plan quantities.
    """
    try:
        data = get_session_data()
        
        if not data.get('bid_proposal'):
            return jsonify({
                'success': False,
                'error': 'No bid proposal uploaded'
            }), 400
        
        if not data.get('plan_quantities'):
            return jsonify({
                'success': False,
                'error': 'No plan quantities extracted. Please upload plan sheets first.'
            }), 400
        
        # Get quantities
        bid_items = data['bid_proposal'].get('line_items', []) or data['bid_proposal'].get('combined_line_items', [])
        plan_items = data['plan_quantities'].get('aggregated', [])
        
        # Compare
        analyzer = BidAnalyzer()
        comparison = analyzer.compare_quantities(bid_items, plan_items)
        
        return jsonify({
            'success': True,
            'comparison': comparison
        })
        
    except Exception as e:
        logger.error(f"Comparison error: {e}\n{traceback.format_exc()}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/api/export/word', methods=['POST'])
def export_word():
    """Export analysis report as Word document"""
    try:
        data = get_session_data()
        
        if not data.get('analysis_results'):
            return jsonify({
                'success': False,
                'error': 'No analysis results to export. Please run analysis first.'
            }), 400
        
        report_gen = ReportGenerator()
        project_name = data.get('proposal_requirements', {}).get('project_info', {}).get('project_name', 'Bid Analysis')
        
        buffer = report_gen.generate_bid_analysis_report(
            data['analysis_results'],
            project_name
        )
        
        filename = f"{project_name.replace(' ', '_')}_Bid_Analysis_{datetime.now().strftime('%Y%m%d')}.docx"
        
        return send_file(
            buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f"Word export error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/excel', methods=['POST'])
def export_excel():
    """Export quantities as Excel spreadsheet"""
    try:
        data = get_session_data()
        export_type = request.json.get('type', 'quantities') if request.is_json else 'quantities'
        
        report_gen = ReportGenerator()
        project_name = data.get('proposal_requirements', {}).get('project_info', {}).get('project_name', 'Bid')
        
        if export_type == 'comparison' and data.get('plan_quantities') and data.get('bid_proposal'):
            # Export comparison
            analyzer = BidAnalyzer()
            bid_items = data['bid_proposal'].get('line_items', []) or data['bid_proposal'].get('combined_line_items', [])
            plan_items = data['plan_quantities'].get('aggregated', [])
            comparison = analyzer.compare_quantities(bid_items, plan_items)
            
            buffer = report_gen.generate_comparison_excel(comparison, project_name)
            filename = f"{project_name.replace(' ', '_')}_Quantity_Comparison_{datetime.now().strftime('%Y%m%d')}.xlsx"
            
        elif data.get('plan_quantities'):
            # Export plan quantities
            quantities = data['plan_quantities'].get('aggregated', [])
            buffer = report_gen.generate_quantity_excel(quantities, project_name)
            filename = f"{project_name.replace(' ', '_')}_Quantities_{datetime.now().strftime('%Y%m%d')}.xlsx"
            
        elif data.get('bid_proposal'):
            # Export bid quantities
            items = data['bid_proposal'].get('line_items', []) or data['bid_proposal'].get('combined_line_items', [])
            buffer = report_gen.generate_quantity_excel(items, project_name)
            filename = f"{project_name.replace(' ', '_')}_Bid_Items_{datetime.now().strftime('%Y%m%d')}.xlsx"
            
        else:
            return jsonify({
                'success': False,
                'error': 'No data to export'
            }), 400
        
        return send_file(
            buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f"Excel export error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/history', methods=['GET'])
def get_history():
    """Get analysis history for current session"""
    data = get_session_data()
    return jsonify({
        'success': True,
        'history': data.get('history', [])
    })


@app.route('/api/clear', methods=['POST'])
def clear_session():
    """Clear session data for a fresh start"""
    sid = get_session_id()
    if sid in session_data:
        session_data[sid] = {
            'proposal_requirements': None,
            'bid_proposal': None,
            'plan_quantities': None,
            'analysis_results': None,
            'history': session_data[sid].get('history', [])  # Keep history
        }
    
    return jsonify({'success': True, 'message': 'Session cleared'})


@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current session status - what's been uploaded"""
    data = get_session_data()
    
    return jsonify({
        'success': True,
        'has_proposal_docs': data.get('proposal_requirements') is not None,
        'has_bid_proposal': data.get('bid_proposal') is not None,
        'has_plan_quantities': data.get('plan_quantities') is not None,
        'has_analysis': data.get('analysis_results') is not None,
        'project_name': data.get('proposal_requirements', {}).get('project_info', {}).get('project_name', '') if data.get('proposal_requirements') else ''
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
