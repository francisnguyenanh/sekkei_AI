import re
import io
import time
import uuid
import logging
import traceback
import openpyxl
from flask import Flask, request, jsonify, send_file, render_template
from pydantic import ValidationError
from schemas.models import parse_request, parse_multi_request, TemplateConfig
from core.generator import ExcelGeneratorService

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024 # 20MB

# ── In-memory template store ──────────────────────────────────────────────────
TEMPLATE_STORE: dict = {}

def _template_summary(name: str, cfg: dict) -> dict:
    """Return lightweight summary for list endpoints."""
    anchors = cfg.get("mapping_anchors", {})
    return {
        "template_name": name,
        "sheet_name": cfg.get("sheet_name", ""),
        "anchor_keys": list(anchors.keys()),
    }

_TABLE_HINTS = ["_start", "_table", "_data", "table"]

def _build_ai_prompt(template_name: str, cfg: dict) -> str:
    anchors = cfg.get("mapping_anchors", {})
    single_keys = {k: v for k, v in anchors.items() if not any(t in k for t in _TABLE_HINTS)}
    table_keys  = {k: v for k, v in anchors.items() if     any(t in k for t in _TABLE_HINTS)}

    single_schema = "\n".join(
        f'    "{k}": "<value>",  // → cell {v}' for k, v in single_keys.items()
    )
    if table_keys:
        table_lines = []
        for k, v in table_keys.items():
            table_lines.append(
                f'    "{k}": [  // starts at cell {v}\n'
                f'      ["col1_header", "col2_header", ...],  // header row\n'
                f'      ["row1_val1",  "row1_val2",  ...]     // data rows...\n'
                f'    ],'
            )
        table_schema = "\n".join(table_lines)
    else:
        table_schema = ""

    field_ref = "\n".join(f"- `{k}` → cell {v}" for k, v in anchors.items())

    return f"""You are a data-fill assistant for an Excel generation system.

## Your task
Generate a JSON object that fills data into the Excel template named "{template_name}" (sheet: "{cfg.get('sheet_name', '')}").

## Output format (STRICT — output ONLY valid JSON, no markdown, no explanation)
{{
  "single_values": {{
{single_schema}
  }},
  "table_data": {{
{table_schema if table_schema else '    // (no table data in this template)'}
  }}
}}

## Rules
1. Output ONLY the JSON object above. Do NOT wrap in ```json``` or add any explanation.
2. Replace every "<value>" with the actual data for that field.
3. For table_data: each entry is an array-of-arrays. Each inner array is one Excel row.
   The number of columns per row must be consistent throughout the table.
4. All values must be JSON primitives: string, number, boolean, or null.
5. Date values: use "YYYY-MM-DD" format strings.
6. Number values: use JSON numbers (no quotes) for pure numeric cells.
7. Do NOT add keys that are not listed above — they will be ignored or cause errors.

## Field reference
{field_ref}

## Example of valid output (2 single values + 1 table):
{{
  "single_values": {{
    "screen_id": "3075",
    "screen_name": "メニュー"
  }},
  "table_data": {{
    "menu_table_start": [
      ["画面番号", "アイコン№", "リボン表示名"],
      ["3075", "001", "メインメニュー"]
    ]
  }}
}}
"""

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

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

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

# ── TEMPLATE STORE ROUTES ─────────────────────────────────────────────────────

@app.route('/api/v1/templates', methods=['GET'])
def list_templates():
    return jsonify({"templates": [_template_summary(n, c) for n, c in TEMPLATE_STORE.items()]}), 200

@app.route('/api/v1/templates', methods=['POST'])
def import_template():
    req_id = str(uuid.uuid4())
    logger.info(f"[ReqID: {req_id}] Import template request")
    data = request.get_json(force=True, silent=True)
    if not isinstance(data, dict):
        return jsonify({"error": "Invalid JSON payload"}), 400
    name = data.get("template_name", "").strip()
    if not name:
        return jsonify({"error": "Missing or empty 'template_name'"}), 400
    tpl_raw = data.get("template")
    if not isinstance(tpl_raw, dict):
        return jsonify({"error": "Missing 'template' object"}), 400
    try:
        TemplateConfig.model_validate(tpl_raw)
    except ValidationError as e:
        return jsonify({"error": "Validation Error", "details": e.errors()}), 422
    TEMPLATE_STORE[name] = tpl_raw
    logger.info(f"[ReqID: {req_id}] Template '{name}' stored")
    return jsonify(_template_summary(name, tpl_raw)), 200

@app.route('/api/v1/templates/<template_name>', methods=['GET'])
def get_template(template_name: str):
    cfg = TEMPLATE_STORE.get(template_name)
    if cfg is None:
        return jsonify({"error": f"Template '{template_name}' not found"}), 404
    return jsonify({"template_name": template_name, "template": cfg}), 200

@app.route('/api/v1/templates/<template_name>', methods=['DELETE'])
def delete_template(template_name: str):
    if template_name not in TEMPLATE_STORE:
        return jsonify({"error": f"Template '{template_name}' not found"}), 404
    del TEMPLATE_STORE[template_name]
    return jsonify({"status": "deleted", "template_name": template_name}), 200

# ── AI PROMPT ROUTE ───────────────────────────────────────────────────────────

@app.route('/api/v1/templates/<template_name>/prompt', methods=['GET'])
def get_ai_prompt(template_name: str):
    cfg = TEMPLATE_STORE.get(template_name)
    if cfg is None:
        return jsonify({"error": f"Template '{template_name}' not found"}), 404
    prompt = _build_ai_prompt(template_name, cfg)
    anchors = cfg.get("mapping_anchors", {})
    return jsonify({
        "prompt": prompt,
        "template_name": template_name,
        "sheet_name": cfg.get("sheet_name", ""),
        "anchor_keys": list(anchors.keys()),
    }), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5016, debug=True)
