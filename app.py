import re
import io
import json
import time
import uuid
import logging
import traceback
import openpyxl
from pathlib import Path
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

# ── Persistent template store ─────────────────────────────────────────────────
TEMPLATE_DIR = Path(__file__).parent / 'template_store'
TEMPLATE_DIR.mkdir(exist_ok=True)

TEMPLATE_STORE: dict = {}

def _load_templates_from_disk():
    """Load all .json files from TEMPLATE_DIR into TEMPLATE_STORE on startup."""
    for f in TEMPLATE_DIR.glob('*.json'):
        try:
            data = json.loads(f.read_text(encoding='utf-8'))
            TEMPLATE_STORE[f.stem] = data
            logger.info(f"Loaded template '{f.stem}' from disk")
        except Exception as e:
            logger.warning(f"Failed to load template '{f.name}': {e}")

def _save_template_to_disk(name: str, cfg: dict):
    path = TEMPLATE_DIR / f'{name}.json'
    path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')

def _delete_template_from_disk(name: str):
    path = TEMPLATE_DIR / f'{name}.json'
    if path.exists():
        path.unlink()

def _template_summary(name: str, cfg: dict) -> dict:
    """Return lightweight summary for list endpoints."""
    anchors = cfg.get("mapping_anchors", {})
    return {
        "template_name": name,
        "sheet_name": cfg.get("sheet_name", ""),
        "anchor_keys": list(anchors.keys()),
    }

_TABLE_HINTS = ["_start", "_table", "_data", "table"]

# ── Table column definitions per template type ────────────────────────────────
_TABLE_COLUMNS = {
    "menu_table_start": [
        ("画面番号",                       '例: "①" "②" ... の連番'),
        ("アイコン№",                      '例: "G03ichiran_02"'),
        ("リボン表示名",                   'メニューに表示される名称'),
        ("権限タイプ",                     '例: "D"'),
        ("画面遷移先（対応シート名）",     '遷移先シート名'),
        ("画面名",                         '遷移先の画面名称'),
        ("追加位置 大分類",               'メニュー配置の大分類'),
        ("追加位置 中分類",               'メニュー配置の中分類'),
        ("グループ名",                     'グループ名（なければ "－"）'),
        ("変更区分",                       '"追加" / "変更" / "削除" / "－"'),
    ],
    "item_table_start": [
        ("画面番号",           '"①" "②" ...'),
        ("項目名",             '項目の表示名'),
        ("項目種類",           '"ラベル" / "テキスト" / "コンボ" / "ボタン" / "チェック" / "グリッド"'),
        ("編集",               '"○" / "×" / "－"'),
        ("文字種",             '"全角" / "半角" / "数字" / "－"'),
        ("IME",                '"ON" / "OFF" / "－"'),
        ("入力文字数 全角",    '桁数（数値）または null'),
        ("入力文字数 半角",    '桁数（数値）または null'),
        ("入力文字数 整数",    '桁数（数値）または null'),
        ("入力文字数 小数",    '桁数（数値）または null'),
        ("重複チェック",       'チェック内容または "－"'),
        ("初期表示",           '初期値または "－"'),
        ("書式",               '書式文字列または "－"'),
        ("変更区分",           '"追加" / "変更" / "削除" / "－"'),
        ("参照先",             '参照マスタ・画面名または null'),
        ("備考",               '補足・制約事項または null'),
    ],
    "list_table_start": [
        ("画面番号",                    '"①" "②" ...'),
        ("項目名（表示名）",           '一覧の列ヘッダー名'),
        ("伝種区分/廃棄区分",          '例: "1.受入" / "2.出荷" / "全て"'),
        ("表示区分",                   '"伝票" / "行" / "合計"'),
        ("必須",                       '"○" / "×" / "－"'),
        ("出力元画面名（入力）",       'データ取得元の画面名'),
        ("出力元項目名（入力）",       'データ取得元の項目名'),
        ("出力元画面名（入力）2",      '第2取得元の画面名'),
        ("出力元項目名（入力）2",      '第2取得元の項目名'),
        ("表示備考",                   '表示に関する補足'),
        ("変更区分",                   '"追加" / "変更" / "削除" / "－"'),
        ("備考",                       'その他備考'),
    ],
    "display_order_start": [
        ("定義区分",  '定義の種類'),
        ("表示順",    '表示順序番号（数値）'),
        ("指定なし",  '"○" / "－"'),
        ("備考",      '備考'),
    ],
    "csv_table_start": [
        ("項目番号",               '連番（数値）'),
        ("帳票項目名",             'CSV出力項目名'),
        ("文字種",                 '"全角" / "半角" / "数字"'),
        ("変更区分",               '"追加" / "変更" / "削除" / "－"'),
        ("出力元画面名（入力）",   'データ取得元の画面名'),
        ("出力元項目名（入力）",   'データ取得元の項目名'),
        ("出力元画面名（入力）2",  '第2取得元の画面名'),
        ("出力元項目名（入力）2",  '第2取得元の項目名'),
    ],
    "ipo_table_start": [
        ("入力画面・入力項目 画面",  '入力元の画面名'),
        ("入力画面・入力項目 項目",  '入力元の項目名'),
        ("プロセス（処理内容）",
         '処理ロジックを以下の形式で記述（\\n改行）:\n'
         '①データ取得元と抽出条件\n②振り分けロジック・優先順位\n'
         '③集計・計算方法\n④按分ロジック（複数作業者がいる場合）\n'
         '⑤出力先・表示方法\n※補足・注意事項'),
        ("出力項目",  '出力先の項目名'),
        ("備考",      '補足事項'),
    ],
    "voucher_table_start": [
        ("帳票番号",                    '"①" "②" ...'),
        ("帳票項目名",                  '帳票上の項目名'),
        ("項目種類",                    '"ラベル" / "データ" / "集計"'),
        ("文字種",                      '"全角" / "半角" / "数字" / "－"'),
        ("表示文字数/幅 全角",          '数値または null'),
        ("表示文字数/幅 半角",          '数値または null'),
        ("表示文字数/幅 整数",          '数値または null'),
        ("表示文字数/幅 小数",          '数値または null'),
        ("フォント サイズ",             '数値または null'),
        ("フォント 太字",               '"○" / "－"'),
        ("フォント 下線",               '"○" / "－"'),
        ("書式",                        '書式文字列または "－"'),
        ("変更区分",                    '"追加" / "変更" / "削除" / "－"'),
        ("出力元画面（入力） 画面名",   'データ取得元の画面名'),
        ("出力元画面（入力） 項目名",   'データ取得元の項目名'),
    ],
}

# ── single_value field descriptions ──────────────────────────────────────────
_SINGLE_FIELD_DESC = {
    "system_name":      '固定値 → "環境将軍R(A1)"',
    "project_number":   '固定値 → 20230927  ※整数、引用符なし',
    "customer_name":    '固定値 → "株式会社ISC"',
    "screen_id":        '画面ID（例: "G055"）',
    "screen_name":      '画面名・帳票名',
    "screen_ver":       '画面バージョン（例: "2.33.6"）',
    "version":          '版数（例: "初回" / "更新"）。なければ "初回"',
    "create_date":      '作成日 → "YYYY-MM-DD" 形式',
    "author":           '作成者名',
    "csv_name":         'CSV出力ファイル名（CSV設計書のみ）',
    "requirements":     (
        'requirements は文字列（配列ではない）。\\n で改行。形式:\n'
        '■{機能名}　{画面種別}\n'
        '①{セクション名}\n　・{項目1}\n　・{項目2}\n'
        '②{セクション名}\n　・{項目1}'
    ),
}

def _build_ai_prompt(template_name: str, cfg: dict) -> str:
    anchors = cfg.get("mapping_anchors", {})
    sheet_name = cfg.get("sheet_name", "")
    single_keys = {k: v for k, v in anchors.items() if not any(t in k for t in _TABLE_HINTS)}
    table_keys  = {k: v for k, v in anchors.items() if     any(t in k for t in _TABLE_HINTS)}

    # ── Build single_values schema block ─────────────────────────────────────
    sv_lines = []
    for k in single_keys:
        desc = _SINGLE_FIELD_DESC.get(k, f"→ cell {anchors[k]}")
        sv_lines.append(f'    "{k}": <value>,  // {desc}')
    single_schema = "\n".join(sv_lines) if sv_lines else '    // (no single_values in this template)'

    # ── Build table_data schema block + column reference ─────────────────────
    table_schema_lines = []
    table_col_reference = []
    for tk in table_keys:
        cell = anchors[tk]
        col_defs = _TABLE_COLUMNS.get(tk)
        if col_defs:
            headers = json.dumps([c[0] for c in col_defs], ensure_ascii=False)
            col_ref_block = "\n".join(
                f"  列{i}: {c[0]} — {c[1]}" for i, c in enumerate(col_defs)
            )
            table_col_reference.append(
                f"### `{tk}` (starts at cell {cell})  — {len(col_defs)}列\n{col_ref_block}"
            )
            table_schema_lines.append(
                f'    "{tk}": [  // starts at cell {cell}\n'
                f'      {headers},  // ← Row 0: header (MUST match exactly)\n'
                f'      [/* row 1 values, {len(col_defs)} columns */],\n'
                f'      [/* row 2 values, {len(col_defs)} columns */]\n'
                f'    ],'
            )
        else:
            table_schema_lines.append(
                f'    "{tk}": [  // starts at cell {cell}\n'
                f'      ["col1_header", "col2_header", ...],\n'
                f'      ["row1_val1",  "row1_val2",  ...]\n'
                f'    ],'
            )
            table_col_reference.append(f"### `{tk}` (starts at cell {cell})\n  (カスタムテーブル — 適切な列を定義してください)")

    table_schema = "\n".join(table_schema_lines) if table_schema_lines else '    // (no table_data in this template)'
    col_ref_section = "\n\n".join(table_col_reference) if table_col_reference else "(テーブルデータなし)"

    # ── Build example ─────────────────────────────────────────────────────────
    example_sv: dict = {}
    for k in single_keys:
        if k == "system_name":       example_sv[k] = "環境将軍R(A1)"
        elif k == "project_number":  example_sv[k] = 20230927
        elif k == "customer_name":   example_sv[k] = "株式会社ISC"
        elif k == "version":         example_sv[k] = "初回"
        elif k == "create_date":     example_sv[k] = "2025-04-19"
        elif k == "screen_id":       example_sv[k] = "XXXX"
        elif k == "screen_name":     example_sv[k] = sheet_name or "画面名"
        elif k == "screen_ver":      example_sv[k] = "1.0.0"
        elif k == "requirements":    example_sv[k] = "■画面名　機能名\n①抽出条件\n　・条件1\n　・条件2"
        else:                        example_sv[k] = f"<{k}の値>"

    example_td: dict = {}
    for tk in table_keys:
        col_defs = _TABLE_COLUMNS.get(tk)
        if col_defs:
            example_td[tk] = [
                [c[0] for c in col_defs],
                ["<値>" if "数値" not in c[1] and "null" not in c[1] else None for c in col_defs],
            ]
        else:
            example_td[tk] = [["col1", "col2"], ["val1", "val2"]]

    example_json = json.dumps(
        {"single_values": example_sv, "table_data": example_td},
        ensure_ascii=False, indent=2
    )

    return f"""あなたは日本語ソフトウェア設計書（カスタマイズ設計書）の**ロジックデータ作成AI**です。
提供された仕様書・要件書を読み取り、**`logic` JSONのみを出力**してください。
ExcelファイルやHTMLは一切出力しません。

---

## コンテキスト

システム：**環境将軍R(A1)**（株式会社ISC、案件番号：20230927）
テンプレート名：**{template_name}**（シート：{sheet_name}）

あなたが出力する `logic` JSON は、システム側でテンプレートと合成されてExcelに変換されます。

---

## 出力スキーマ（厳守）

以下のJSONのみを出力してください。コードブロック・説明文・前置き・後書きは一切不要。

{{
  "single_values": {{
{single_schema}
  }},
  "table_data": {{
{table_schema}
  }}
}}

---

## single_values の入力ルール

- `system_name` → 固定値 `"環境将軍R(A1)"`
- `project_number` → 固定値 `20230927`（整数、引用符なし）
- `customer_name` → 固定値 `"株式会社ISC"`
- `version` → なければ `"初回"`
- `create_date` → `"YYYY-MM-DD"` 形式
- `requirements` → 文字列（配列ではない）。`\\n` で改行。形式：
    ■{{機能名}}　{{画面種別}}
    ①{{セクション名}}
    　・{{項目1}}
    ②{{セクション名}}
    　・{{項目1}}

---

## table_data の列定義

{col_ref_section}

---

## 値の型ルール（厳守）

| 値の種類 | 正しい書き方 | ❌ 間違い |
|---------|------------|---------|
| 数値 | `15000` | `"15000"` |
| 案件番号 | `20230927` | `"20230927"` |
| 日付 | `"2025-04-19"` | `"2025/04/19"` |
| 空欄 | `null` | `""` / `"-"` |
| フラグ（該当なし） | `"－"` | `"-"` / `null` |
| 改行 | `\\n`（文字列内） | 配列 |
| boolean的フラグ | `"○"` または `"－"` | `true` / `false` |

---

## 出力例

{example_json}

---

## 出力前の自己チェック

1. `mapping_anchors` にないキーを追加していないか
2. 全テーブル行の列数がヘッダー行と一致しているか
3. `project_number` が整数 `20230927`（引用符なし）か
4. 空セルが `null`（`""` でも `"-"` でもない）か
5. `requirements` が文字列（配列ではない）か
6. 出力がJSON **のみ**（コードブロックや説明文がない）か
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
    _save_template_to_disk(name, tpl_raw)
    logger.info(f"[ReqID: {req_id}] Template '{name}' stored to disk")
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
    _delete_template_from_disk(template_name)
    return jsonify({"status": "deleted", "template_name": template_name}), 200

@app.route('/api/v1/templates/<template_name>/download', methods=['GET'])
def download_template(template_name: str):
    path = TEMPLATE_DIR / f'{template_name}.json'
    if not path.exists():
        return jsonify({"error": f"Template '{template_name}' not found on disk"}), 404
    return send_file(
        path,
        mimetype='application/json',
        as_attachment=True,
        download_name=f'{template_name}.json'
    )

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

_load_templates_from_disk()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5016, debug=True)
