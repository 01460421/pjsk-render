#!/usr/bin/env python3
"""
render_server.py — 雲端渲染 API 伺服器
接收 JSON 渲染請求 → 回傳 PNG 圖片

啟動: python3 render_server.py
預設 port: 5100
"""
import os, json, traceback
from io import BytesIO
from flask import Flask, request, send_file, jsonify

# 渲染函數
from img_render import (
    render_table_image, render_info_card, render_message_box,
    render_help_image, render_line_chart, Theme
)
from render_funcs import (
    create_push_plan_image, create_ranking_detail_image,
    create_ranking_list_image, create_ranking_chart,
    create_schedule_image, create_member_table_image,
    create_hours_table_image, find_push_plans,
    SONG_DB
)

app = Flask(__name__)

# API 金鑰 (可選, 從環境變數讀取)
API_KEY = os.getenv('RENDER_API_KEY', '')

# Theme 顏色對照表 (讓客戶端傳字串)
THEME_COLORS = {name: getattr(Theme, name) for name in dir(Theme) if not name.startswith('_')}

def resolve_color(val):
    """將 Theme.RED 等字串解析為實際顏色值"""
    if isinstance(val, str) and val.startswith('Theme.'):
        attr = val[6:]
        return THEME_COLORS.get(attr, val)
    return val

def resolve_colors_deep(obj):
    """遞迴解析 dict/list 中的顏色引用"""
    if isinstance(obj, dict):
        return {k: resolve_colors_deep(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [resolve_colors_deep(v) for v in obj]
    elif isinstance(obj, str) and obj.startswith('Theme.'):
        return resolve_color(obj)
    return obj

def check_auth():
    """驗證 API 金鑰 (若有設定)"""
    if not API_KEY:
        return True
    key = request.headers.get('X-API-Key', '')
    return key == API_KEY

# 函數對照表
FUNC_MAP = {
    # img_render 基礎函數
    'render_table_image': render_table_image,
    'render_info_card': render_info_card,
    'render_message_box': render_message_box,
    'render_help_image': render_help_image,
    'render_line_chart': render_line_chart,
    # render_funcs 複合函數
    'create_push_plan_image': create_push_plan_image,
    'create_ranking_detail_image': create_ranking_detail_image,
    'create_ranking_list_image': create_ranking_list_image,
    'create_ranking_chart': create_ranking_chart,
    'create_schedule_image': create_schedule_image,
    'create_member_table_image': create_member_table_image,
    'create_hours_table_image': create_hours_table_image,
}

@app.route('/health', methods=['GET'])
def health():
    return jsonify(status='ok', songs=len(SONG_DB), funcs=list(FUNC_MAP.keys()))

@app.route('/render', methods=['POST'])
def render():
    if not check_auth():
        return jsonify(error='Unauthorized'), 401
    
    data = request.json
    if not data:
        return jsonify(error='No JSON body'), 400
    
    func_name = data.get('func')
    kwargs = data.get('kwargs', {})
    
    func = FUNC_MAP.get(func_name)
    if not func:
        return jsonify(error=f'Unknown function: {func_name}', available=list(FUNC_MAP.keys())), 400
    
    try:
        # 解析顏色引用
        kwargs = resolve_colors_deep(kwargs)
        
        # 特殊處理: col_colors 的 key 需要轉 int (JSON 不支援 int key)
        if 'col_colors' in kwargs and isinstance(kwargs['col_colors'], dict):
            kwargs['col_colors'] = {int(k): v for k, v in kwargs['col_colors'].items()}
        if 'row_highlights' in kwargs and isinstance(kwargs['row_highlights'], dict):
            kwargs['row_highlights'] = {int(k): v for k, v in kwargs['row_highlights'].items()}
        
        # 特殊處理: figsize tuple
        if 'figsize' in kwargs and isinstance(kwargs['figsize'], list):
            kwargs['figsize'] = tuple(kwargs['figsize'])
        
        # 特殊處理: y_formatter (不能序列化，用預設)
        if 'y_formatter' in kwargs:
            fmt_type = kwargs.pop('y_formatter')
            if fmt_type == 'score_w':
                kwargs['y_formatter'] = lambda v,_: f"{v:,.0f}W" if v>=1 else f"{v:.1f}W"
        
        result = func(**kwargs)
        
        if result is None:
            return jsonify(error='No image generated'), 204
        
        if isinstance(result, BytesIO):
            result.seek(0)
            return send_file(result, mimetype='image/png')
        
        return jsonify(error='Unexpected result type'), 500
        
    except Exception as e:
        tb = traceback.format_exc()
        print(f"[render_server] Error in {func_name}: {tb}")
        return jsonify(error=str(e), traceback=tb), 500

@app.route('/push_plans', methods=['POST'])
def push_plans():
    """計算肘人方案 (不含圖片)"""
    if not check_auth():
        return jsonify(error='Unauthorized'), 401
    
    data = request.json
    try:
        plans = find_push_plans(**data)
        return jsonify(plans=plans)
    except Exception as e:
        return jsonify(error=str(e)), 500

if __name__ == '__main__':
    port = int(os.getenv('RENDER_PORT', 5100))
    debug = os.getenv('RENDER_DEBUG', '0') == '1'
    print(f"[Render Server] Starting on port {port}")
    print(f"[Render Server] Songs: {len(SONG_DB)}")
    print(f"[Render Server] Functions: {list(FUNC_MAP.keys())}")
    print(f"[Render Server] Auth: {'enabled' if API_KEY else 'disabled'}")
    app.run(host='0.0.0.0', port=port, debug=debug)
