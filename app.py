"""
Bid Proposal Agent - Web Application
Expert civil engineering bid analysis and proposal tool
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

# Lazy load heavy modules to speed up startup
_bid_estimator = None
_proposal_parser = None
_bid_analyzer = None
_report_generator = None

def get_bid_estimator():
    global _bid_estimator
    if _bid_estimator is None:
        from agent.quantity_calculator import BidEstimator
        _bid_estimator = BidEstimator
    return _bid_estimator()

def get_proposal_parser():
    global _proposal_parser
    if _proposal_parser is None:
        from agent.proposal_parser import ProposalParser
        _proposal_parser = ProposalParser
    return _proposal_parser()

def get_bid_analyzer():
    global _bid_analyzer
    if _bid_analyzer is None:
        from agent.bid_analyzer import BidAnalyzer
        _bid_analyzer = BidAnalyzer
    return _bid_analyzer()

def get_report_generator():
    global _report_generator
    if _report_generator is None:
        from agent.report_generator import ReportGenerator
        _report_generator = ReportGenerator
    return _report_generator()

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
            'bid_docs': None,
            'current_proposal': None,
            'estimate': None,
            'analysis': None,
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


@app.route('/api/parse-bid-docs', methods=['POST'])
def parse_bid_documents():
    """
    Parse bid documents (RFP, bid schedule, specs) to extract requirements.
    """
    temp_files = []
    
    try:
        if 'files' not in request.files:
            return jsonify({'success': False, 'error': 'No files uploaded'}), 400
        
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            return jsonify({'success': False, 'error': 'No files selected'}), 400
        
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
                'error': 'No valid files uploaded. Supported: PDF, Excel'
            }), 400
        
        # Parse documents
        parser = get_proposal_parser()
        
        if len(file_paths) == 1:
            result = parser.parse_bid_document(file_paths[0])
        else:
            result = parser.parse_multiple_documents(file_paths)
        
        # Store in session
        data = get_session_data()
        data['bid_docs'] = result
        
        # Extract summary
        line_items = parser.extract_line_items_table(result)
        key_dates = parser.get_key_dates(result)
        summary = parser.generate_bid_summary(result)
        
        return jsonify({
            'success': True,
            'project_info': result.get('project_info', {}),
            'scope': result.get('scope', {}),
            'line_items': line_items,
            'line_item_count': len(line_items),
            'requirements': result.get('requirements', {}),
            'specifications': result.get('specifications', {}),
            'key_dates': key_dates,
            'summary': summary,
            'files_processed': len(file_paths)
        })
        
    except Exception as e:
        logger.error(f"Parse bid docs error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500
    
    finally:
        for path in temp_files:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    os.rmdir(os.path.dirname(path))
            except Exception:
                pass


@app.route('/api/parse-proposal', methods=['POST'])
def parse_current_proposal():
    """
    Parse an existing proposal being worked on for review.
    """
    temp_files = []
    
    try:
        if 'files' not in request.files:
            return jsonify({'success': False, 'error': 'No files uploaded'}), 400
        
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            return jsonify({'success': False, 'error': 'No files selected'}), 400
        
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
        
        if not file_paths:
            return jsonify({'success': False, 'error': 'No valid files uploaded'}), 400
        
        # Parse using estimator
        estimator = get_bid_estimator()
        result = estimator.analyze_bid_documents(file_paths)
        
        # Store in session
        data = get_session_data()
        data['current_proposal'] = result
        
        return jsonify({
            'success': True,
            'project_info': result.get('project_summary', {}),
            'line_items': result.get('line_items', []),
            'line_item_count': len(result.get('line_items', [])),
            'bid_total': result.get('bid_total', {}),
            'files_processed': len(file_paths)
        })
        
    except Exception as e:
        logger.error(f"Parse proposal error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500
    
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
    Analyze bid documents and/or proposal with expert feedback.
    Works with just bid docs - no proposal required.
    """
    try:
        data = get_session_data()
        
        # Check if we have bid docs
        if not data.get('bid_docs'):
            return jsonify({
                'success': False,
                'error': 'No bid documents uploaded. Please upload bid docs first.'
            }), 400
        
        # Run analysis using the bid estimator
        estimator = get_bid_estimator()
        analyzer = get_bid_analyzer()
        report_gen = get_report_generator()
        
        # Generate estimate from bid docs
        estimate = analyzer.start_proposal(data['bid_docs'])
        data['estimate'] = estimate
        
        # If we have a current proposal, analyze it against the bid docs
        proposal_data = data.get('current_proposal') or estimate
        
        # Run expert analysis
        analysis = analyzer.analyze_proposal(proposal_data, data['bid_docs'])
        
        # Get status and recommendations
        status = analyzer.get_bid_status(analysis)
        recommendations = analyzer.generate_recommendations(analysis)
        
        analysis['status'] = status
        analysis['prioritized_recommendations'] = recommendations
        analysis['estimate'] = estimate
        
        # Store results
        data['analysis'] = analysis
        
        # Add to history
        project_name = data['bid_docs'].get('project_info', {}).get('project_name', 'Unknown Project')
        history_entry = {
            'id': str(uuid.uuid4()),
            'timestamp': datetime.now().isoformat(),
            'project_name': project_name,
            'status': status,
            'total_bid': estimate.get('summary', {}).get('total_bid', 0)
        }
        data['history'].insert(0, history_entry)
        data['history'] = data['history'][:20]
        
        # Generate HTML report
        html_report = report_gen.generate_html_report(analysis, project_name)
        
        return jsonify({
            'success': True,
            'status': status,
            'overall_assessment': analysis.get('overall_assessment', {}),
            'completeness': analysis.get('completeness', {}),
            'pricing_analysis': analysis.get('pricing_analysis', {}),
            'risks': analysis.get('risks', []),
            'recommendations': recommendations[:10],
            'bid_strategy': analysis.get('bid_strategy', {}),
            'estimate': estimate,
            'html_report': html_report,
            'project_name': project_name
        })
        
    except Exception as e:
        logger.error(f"Analysis error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/pdf', methods=['POST'])
def export_pdf():
    """Export analysis report as PDF"""
    try:
        data = get_session_data()
        
        if not data.get('analysis'):
            return jsonify({
                'success': False,
                'error': 'No analysis to export. Run analysis first.'
            }), 400
        
        report_gen = get_report_generator()
        project_name = data.get('bid_docs', {}).get('project_info', {}).get('project_name', 'Bid Analysis')
        
        buffer = report_gen.generate_pdf_report(data['analysis'], project_name)
        
        filename = f"{project_name.replace(' ', '_')}_Bid_Analysis_{datetime.now().strftime('%Y%m%d')}.pdf"
        
        return send_file(
            buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f"PDF export error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/export/word', methods=['POST'])
def export_word():
    """Export analysis report as Word document"""
    try:
        data = get_session_data()
        
        if not data.get('analysis'):
            return jsonify({
                'success': False,
                'error': 'No analysis to export. Run analysis first.'
            }), 400
        
        report_gen = get_report_generator()
        project_name = data.get('bid_docs', {}).get('project_info', {}).get('project_name', 'Bid Analysis')
        
        buffer = report_gen.generate_bid_analysis_report(data['analysis'], project_name)
        
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
    """Export bid estimate as Excel spreadsheet"""
    try:
        data = get_session_data()
        
        # Get items from estimate, analysis, or bid docs
        estimate = data.get('analysis', {}).get('estimate') or data.get('estimate') or data.get('current_proposal')
        
        if not estimate:
            # Try to get line items from bid docs
            bid_docs = data.get('bid_docs')
            if bid_docs:
                items = bid_docs.get('line_items', [])
                project_name = bid_docs.get('project_info', {}).get('project_name', 'Bid')
            else:
                return jsonify({
                    'success': False,
                    'error': 'No data to export. Upload bid docs or run analysis first.'
                }), 400
        else:
            items = estimate.get('bid_items', []) or estimate.get('line_items', [])
            project_name = data.get('bid_docs', {}).get('project_info', {}).get('project_name', 'Bid')
        
        report_gen = get_report_generator()
        buffer = report_gen.generate_bid_excel(items, project_name, estimate.get('summary', {}) if estimate else {})
        
        filename = f"{project_name.replace(' ', '_')}_Bid_Estimate_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        return send_file(
            buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        logger.error(f"Excel export error: {e}\n{traceback.format_exc()}")
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current session status"""
    data = get_session_data()
    
    return jsonify({
        'success': True,
        'has_bid_docs': data.get('bid_docs') is not None,
        'has_proposal': data.get('current_proposal') is not None,
        'has_estimate': data.get('estimate') is not None,
        'has_analysis': data.get('analysis') is not None,
        'project_name': data.get('bid_docs', {}).get('project_info', {}).get('project_name', '') if data.get('bid_docs') else ''
    })


@app.route('/api/clear', methods=['POST'])
def clear_session():
    """Clear session data for a fresh start"""
    sid = get_session_id()
    if sid in session_data:
        session_data[sid] = {
            'bid_docs': None,
            'current_proposal': None,
            'estimate': None,
            'analysis': None,
            'history': session_data[sid].get('history', [])
        }
    
    return jsonify({'success': True, 'message': 'Session cleared'})


@app.route('/api/history', methods=['GET'])
def get_history():
    """Get analysis history"""
    data = get_session_data()
    return jsonify({
        'success': True,
        'history': data.get('history', [])
    })


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
