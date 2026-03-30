import re
import io
import time
import uuid
import logging
import traceback
import openpyxl
from flask import Flask, request, jsonify, send_file
from pydantic import ValidationError
from schemas.models import parse_request, parse_multi_request
from core.generator import ExcelGeneratorService

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024 # 20MB

@app.errorhandler(413)
def request_entity_too_large(e):
    req_id = str(uuid.uuid4())
    logger.error(f"[ReqID: {req_id}] Payload too large (>20MB)")
    return jsonify({"error": "Request entity too large", "max_bytes": 20 * 1024 * 1024}), 413
        
@app.after_request
def append_request_id(response):
    req_id = getattr(request, 'req_id', None)
    if req_id: response.headers['X-Request-ID'] = req_id
    return response

@app.route('/api/v1/generate-excel', methods=['POST'])
def generate_excel():
    req_id = str(uuid.uuid4())
    request.req_id = req_id
    logger.info(f"[ReqID: {req_id}] Received single sheet request")
    start_time = time.time()
    
    try:
        data = request.get_json(force=True, silent=True)
        if not isinstance(data, dict):
            return jsonify({
                "error": "Invalid JSON payload",
                "hint": "Ensure Content-Type is application/json and body is a JSON object"
            }), 400
            
        gen_request = parse_request(data)
        parse_time = time.time() - start_time
        logger.info(f"[ReqID: {req_id}] Parsed request successfully in {parse_time:.4f}s")
        
        render_start = time.time()
        generator = ExcelGeneratorService(gen_request)
        excel_io = generator.generate()
        render_time = time.time() - render_start
        
        raw_name = gen_request.template.sheet_name
        safe_name = re.sub(r'[^\w\-. ]', '_', raw_name).strip()
        if not safe_name:
            safe_name = 'output'
        filename = f"{safe_name.replace(' ', '_')}.xlsx"
        logger.info(f"[ReqID: {req_id}] Excel rendered in {render_time:.4f}s. Output: {filename}")
        
        return send_file(
            excel_io,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except ValidationError as e:
        logger.error(f"[ReqID: {req_id}] Pydantic validation error: {e}")
        return jsonify({"error": "Validation Error", "details": e.errors()}), 422
    except Exception as e:
        logger.error(f"[ReqID: {req_id}] Internal Server Error: {e}\n{traceback.format_exc()}")
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

@app.route('/api/v1/generate-excel-multi', methods=['POST'])
def generate_excel_multi():
    req_id = str(uuid.uuid4())
    request.req_id = req_id
    logger.info(f"[ReqID: {req_id}] Received multi-sheet request")
    start_time = time.time()
    
    try:
        data = request.get_json(force=True, silent=True)
        if not isinstance(data, dict):
            return jsonify({
                "error": "Invalid JSON payload",
                "hint": "Ensure Content-Type is application/json and body is a JSON object"
            }), 400
            
        multi_req = parse_multi_request(data)
        logger.info(f"[ReqID: {req_id}] Parsed {len(multi_req.sheets)} sheets configuration")
        
        render_start = time.time()
        wb = openpyxl.Workbook()
        
        for sheet_req in multi_req.sheets:
             svc = ExcelGeneratorService(sheet_req, workbook=wb)
             svc.generate(save=False) 
             
        final_io = io.BytesIO()
        wb.save(final_io)
        final_io.seek(0)
        
        render_time = time.time() - render_start
        
        filename = request.args.get('filename', 'output_multi.xlsx')
        filename = re.sub(r'[^\w\-. ]', '_', filename).strip()
        if not filename:
            filename = 'output_multi.xlsx'
        if not filename.endswith('.xlsx'): filename += '.xlsx'
            
        logger.info(f"[ReqID: {req_id}] Multi-sheet Excel rendered in {render_time:.4f}s. Output: {filename}")
        
        return send_file(
            final_io,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except ValidationError as e:
        logger.error(f"[ReqID: {req_id}] Pydantic validation error: {e}")
        return jsonify({"error": "Validation Error", "details": e.errors()}), 422
    except Exception as e:
        logger.error(f"[ReqID: {req_id}] Internal Server Error: {e}\n{traceback.format_exc()}")
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy"}), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5016, debug=True)
