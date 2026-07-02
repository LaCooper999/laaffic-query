"""
NodeSMS RCS 任务查询工具
运行后在浏览器打开: http://localhost:5051

获取 Token 方法:
  1. 打开 https://clientcdn.nodesms.com/#/campaign/rcsTask
  2. 按 F12 → 网络(Network) → 找任意请求
  3. 复制请求头里的 Authorization 值（Bearer xxxx）

获取 API 端点方法:
  在网络面板里找加载任务列表的那条 POST 请求，复制完整 URL
"""

from flask import Flask, request, jsonify, render_template_string
import requests, os, json

SHEET_ID = "14Ykfo6bD56z14pWTjwWCzFFVlmT8p51GX2a49PU1hLw"

app = Flask(__name__)

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NodeSMS RCS 查询</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif; background: #f0f2f5; min-height: 100vh; padding: 24px; }
.container { max-width: 1100px; margin: 0 auto; }
h1 { font-size: 22px; color: #1a1a2e; margin-bottom: 20px; font-weight: 700; }
.card { background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
label { display: block; font-size: 13px; color: #666; margin-bottom: 6px; font-weight: 500; }
input, textarea { width: 100%; border: 1.5px solid #e0e0e0; border-radius: 8px; padding: 10px 14px; font-size: 14px; outline: none; transition: border-color 0.2s; color: #333; }
input:focus, textarea:focus { border-color: #0ea5e9; }
textarea { resize: vertical; min-height: 80px; font-family: monospace; font-size: 12px; }
.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }
.row3 { display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 16px; margin-bottom: 16px; }
.full { margin-bottom: 16px; }
.btn { border: none; border-radius: 8px; padding: 12px 28px; font-size: 15px; cursor: pointer; font-weight: 600; transition: background 0.2s; }
.btn-query  { background: #0ea5e9; color: white; }
.btn-query:hover  { background: #0284c7; }
.btn-query:disabled { background: #7dd3fc; cursor: not-allowed; }
.btn-export { background: #059669; color: white; padding: 10px 20px; font-size: 14px; }
.btn-export:hover { background: #047857; }
.btn-sync   { background: #2563eb; color: white; padding: 10px 20px; font-size: 14px; }
.btn-sync:hover   { background: #1d4ed8; }
.btn-sync:disabled { background: #93c5fd; cursor: not-allowed; }
.btn-row { display: flex; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 16px; }
.stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 20px; }
.stat-card { border-radius: 10px; padding: 18px; text-align: center; }
.stat-card.sub  { background: #e0f2fe; }
.stat-card.send { background: #d1fae5; }
.stat-card.cnt  { background: #fef3c7; }
.stat-label { font-size: 12px; color: #555; margin-bottom: 8px; }
.stat-num   { font-size: 30px; font-weight: 700; }
.stat-card.sub  .stat-num { color: #0369a1; }
.stat-card.send .stat-num { color: #065f46; }
.stat-card.cnt  .stat-num { color: #92400e; }
table { width: 100%; border-collapse: collapse; }
th { background: #f0f9ff; color: #0369a1; font-size: 13px; padding: 12px 14px; text-align: left; border-bottom: 2px solid #bae6fd; white-space: nowrap; }
td { padding: 10px 14px; border-bottom: 1px solid #f0f0f0; font-size: 13px; color: #333; }
tr:hover td { background: #f0f9ff; }
.num-sub  { font-weight: 700; color: #0369a1; }
.num-send { font-weight: 700; color: #059669; }
.empty  { text-align: center; color: #999; padding: 48px; font-size: 15px; }
.loading { text-align: center; color: #0ea5e9; padding: 24px; font-size: 15px; }
.error-box { color: #dc2626; padding: 12px 16px; background: #fef2f2; border-radius: 8px; margin-top: 14px; border: 1px solid #fecaca; }
.tip  { font-size: 12px; color: #888; margin-top: 5px; line-height: 1.6; }
.table-wrap { overflow-x: auto; }
.sync-ok  { color: #059669; padding: 10px 14px; background: #d1fae5; border-radius: 8px; margin-bottom: 12px; font-size: 14px; }
.sync-err { color: #dc2626; padding: 10px 14px; background: #fef2f2; border-radius: 8px; margin-bottom: 12px; font-size: 14px; }
.section-title { font-size: 13px; font-weight: 600; color: #555; margin-bottom: 10px; padding-bottom: 6px; border-bottom: 1px solid #eee; }
details summary { cursor: pointer; font-size: 13px; color: #0ea5e9; margin-bottom: 10px; }
</style>
</head>
<body>
<div class="container">
  <h1>📱 NodeSMS RCS 任务查询</h1>

  <div class="card">
    <!-- Token -->
    <div class="full">
      <label>🔑 Bearer Token</label>
      <textarea id="token" placeholder="粘贴 Token（带不带 Bearer 前缀都行）..."></textarea>
      <div class="tip">获取方法：打开 clientcdn.nodesms.com → F12 → 网络 → 任意请求 → 复制 Authorization 头的值</div>
    </div>

    <!-- 筛选条件 -->
    <div class="row2">
      <div>
        <label>任务/文件名（模糊搜索，留空显示全部）</label>
        <input type="text" id="taskName" placeholder="输入关键词...">
      </div>
      <div>
        <label>日期</label>
        <input type="date" id="queryDate">
      </div>
    </div>

    <button class="btn btn-query" onclick="fetchData()" id="btn">🔍 查 询</button>
    <div id="error"></div>
  </div>

  <div class="card" id="result" style="display:none">
    <div id="content"></div>
  </div>

</div>

<script>
// 初始化日期（今天）
const now = new Date();
const pad = n => String(n).padStart(2,'0');
const todayStr = now.getFullYear()+'-'+pad(now.getMonth()+1)+'-'+pad(now.getDate());
document.getElementById('queryDate').value = todayStr;

function toApiTime(v, isEnd) {
  if (!v) return '';
  return isEnd ? v + ' 23:59:59' : v + ' 00:00:00';
}

async function fetchData() {
  const rawToken = document.getElementById('token').value.trim();
  if (!rawToken) { showError('请先填写 Bearer Token'); return; }

  const token          = rawToken.startsWith('Bearer ') ? rawToken.slice(7) : rawToken;
  const taskNameFilter = document.getElementById('taskName').value.trim();
  const queryDate      = document.getElementById('queryDate').value;
  const startDate      = queryDate;
  const endDate        = queryDate;

  const btn = document.getElementById('btn');
  btn.disabled = true; btn.textContent = '查询中...';
  document.getElementById('error').innerHTML = '';
  document.getElementById('result').style.display = 'block';
  document.getElementById('content').innerHTML = '<div class="loading">⏳ 正在加载数据...</div>';

  try {
    const resp = await fetch('/api/query', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        __token__:     token,
        __taskName__:  taskNameFilter,
        __startDate__: toApiTime(startDate, false),
        __endDate__:   toApiTime(endDate, true),
      })
    });
    const data = await resp.json();
    if (!data.ok) { showError('接口错误: ' + (data.msg || '未知错误')); return; }
    window._lastRecords = data.records || [];
    renderTable(window._lastRecords, startDate);
  } catch(e) {
    showError('请求失败: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '🔍 查 询';
  }
}

function renderTable(records, date) {
  if (!records.length) {
    document.getElementById('content').innerHTML = '<div class="empty">🔍 没有找到匹配的任务</div>';
    return;
  }

  let totalSub = 0, totalSend = 0;
  records.forEach(r => { totalSub += r.submitCount||0; totalSend += r.sendCount||0; });

  const savedTab = localStorage.getItem('nodesms_sheetTab') || '';
  let html = '<div class="btn-row">'
    + '<span style="font-size:14px;color:#333;font-weight:600">共 ' + records.length + ' 条任务</span>'
    + '<button class="btn btn-export" onclick="exportExcel()">📥 导出 Excel</button>'
    + '<button class="btn btn-sync" onclick="syncSheets()" id="syncBtn">📊 同步到 Google Sheet</button>'
    + '<input type="text" id="sheetTab" placeholder="标签页名称" value="' + savedTab + '" oninput="localStorage.setItem(\'nodesms_sheetTab\',this.value)" style="width:140px;padding:10px 12px;font-size:14px;border:1.5px solid #c7d2fe;border-radius:8px;outline:none;color:#333;">'
    + '</div>'
    + '<div id="syncMsg"></div>'
    + '<div class="stats">'
    + '<div class="stat-card sub"><div class="stat-label">提交数（合计）</div><div class="stat-num">' + totalSub.toLocaleString() + '</div></div>'
    + '<div class="stat-card send"><div class="stat-label">发送数/条数（合计）</div><div class="stat-num">' + totalSend.toLocaleString() + '</div></div>'
    + '<div class="stat-card cnt"><div class="stat-label">任务数量</div><div class="stat-num">' + records.length + '</div></div>'
    + '</div>'
    + '<div class="table-wrap"><table><thead><tr>'
    + '<th>#</th><th>文件名称</th><th>提交</th><th>发送</th><th>重复</th><th>错误</th><th>状态</th><th>时间</th>'
    + '</tr></thead><tbody>';

  records.forEach((r, i) => {
    html += '<tr>'
      + '<td style="color:#999">' + (i+1) + '</td>'
      + '<td style="font-size:12px;font-family:monospace">' + (r.fileName||'-') + '</td>'
      + '<td class="num-sub">'  + (r.submitCount||0).toLocaleString() + '</td>'
      + '<td class="num-send">' + (r.sendCount||0).toLocaleString()   + '</td>'
      + '<td style="color:#999">' + (r.repeatCount||0).toLocaleString()  + '</td>'
      + '<td style="color:' + ((r.errorCount||0)>0?'#dc2626':'#999') + '">' + (r.errorCount||0).toLocaleString() + '</td>'
      + '<td style="font-size:12px">' + (r.status||'-') + '</td>'
      + '<td style="font-size:12px;color:#999">' + (r.sendTime||r.createTime||'-') + '</td>'
      + '</tr>';
  });
  html += '</tbody></table></div>';
  document.getElementById('content').innerHTML = html;
}

async function syncSheets() {
  const records  = window._lastRecords || [];
  if (!records.length) return;

  const sheetTab = document.getElementById('sheetTab').value.trim();
  const date     = document.getElementById('queryDate').value;

  if (!sheetTab) { alert('请先填写标签页名称'); document.getElementById('sheetConfig').open = true; return; }

  const btn = document.getElementById('syncBtn');
  btn.disabled = true; btn.textContent = '同步中...';
  document.getElementById('syncMsg').innerHTML = '';

  try {
    const resp = await fetch('/api/sync-sheets', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ records, sheetTab, date })
    });
    const data = await resp.json();
    if (data.ok) {
      document.getElementById('syncMsg').innerHTML =
        '<div class="sync-ok">✅ 同步完成！写入 ' + data.written + ' 行' +
        (data.skipped ? '，未匹配 ' + data.skipped + ' 行' : '') + '</div>';
    } else {
      document.getElementById('syncMsg').innerHTML =
        '<div class="sync-err">❌ ' + (data.error||'同步失败') + '</div>';
    }
  } catch(e) {
    document.getElementById('syncMsg').innerHTML = '<div class="sync-err">❌ ' + e.message + '</div>';
  } finally {
    btn.disabled = false; btn.textContent = '📊 同步到 Google Sheet';
  }
}

function exportExcel() {
  const records = window._lastRecords || [];
  if (!records.length) return;
  const rows = [['文件名称','提交数','发送数','重复数','错误数','状态','时间']];
  records.forEach(r => rows.push([
    r.fileName||'', r.submitCount||0, r.sendCount||0,
    r.repeatCount||0, r.errorCount||0, r.status||'', r.sendTime||r.createTime||''
  ]));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:40},{wch:10},{wch:10},{wch:8},{wch:8},{wch:12},{wch:22}];
  XLSX.utils.book_append_sheet(wb, ws, 'RCS任务');
  const p = n => String(n).padStart(2,'0');
  XLSX.writeFile(wb, 'nodesms_rcs_'+now.getFullYear()+p(now.getMonth()+1)+p(now.getDate())+'.xlsx');
}

function showError(msg) {
  document.getElementById('error').innerHTML = '<div class="error-box">❌ ' + msg + '</div>';
  document.getElementById('result').style.display = 'none';
}
</script>
</body>
</html>
"""


def _parse_records(data: dict, task_name_filter: str) -> list[dict]:
    """
    从 API 响应中提取任务列表，兼容多种字段名约定。
    NodeSMS 字段映射（根据实际响应调整）:
      文件名 → fileName / file_name / taskName / name
      提交   → submitCount / submit / submitNum
      发送   → sendCount / send / sendNum
      重复   → repeatCount / repeat / repeatNum
      错误   → errorCount / error / errorNum
      状态   → status / taskStatus / statusName
      时间   → sendTime / createTime / startTime
    """
    raw = (
        data.get("data", {}).get("records") or
        data.get("data", {}).get("list")    or
        data.get("data", {}).get("rows")    or
        data.get("data", {}).get("data")    or
        (data.get("data") if isinstance(data.get("data"), list) else None) or
        data.get("records") or data.get("list") or []
    )

    STATUS_MAP = {1: '等待', 2: '群发中', 3: '完成', 4: '失败', 5: '已暂停'}

    results = []
    for r in raw:
        file_name = str(r.get("name") or "").strip()
        if task_name_filter and task_name_filter.lower() not in file_name.lower():
            continue
        status_code = r.get("status", 0)
        results.append({
            "fileName":    file_name,
            "submitCount": int(r.get("submitNum") or 0),
            "sendCount":   int(r.get("sendNum")   or 0),
            "repeatCount": int(r.get("repeatNum") or 0),
            "errorCount":  int(r.get("errorNum")  or 0),
            "status":      STATUS_MAP.get(status_code, str(status_code)),
            "sendTime":    str(r.get("sendTimeStr") or r.get("createdAt") or ""),
        })
    return results


@app.route('/')
def index():
    return render_template_string(HTML_PAGE)


BASE_URL = 'https://apip.nodesms.com/client/mass/campaignUserTask/taskListRcs'

# 自动探测 productId 的候选端点
PRODUCT_DETECT_URLS = [
    'https://apip.nodesms.com/client/mass/product/list',
    'https://apip.nodesms.com/client/product/list',
    'https://apip.nodesms.com/client/user/product',
]

_product_id_cache = {}   # token -> productId


def _get_headers(token: str) -> dict:
    return {
        'Authorization': 'Bearer ' + token,
        'Accept':        'application/json, text/plain, */*',
        'Origin':        'https://clientcdn.nodesms.com',
        'Referer':       'https://clientcdn.nodesms.com/',
        'User-Agent':    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }


def _detect_product_id(token: str) -> str | None:
    """尝试从已知端点自动获取 productId，失败则返回 None"""
    if token in _product_id_cache:
        return _product_id_cache[token]

    headers = _get_headers(token)
    for url in PRODUCT_DETECT_URLS:
        try:
            r = requests.get(url, headers=headers, timeout=10)
            if r.status_code != 200:
                continue
            d = r.json()
            # 尝试常见结构
            items = (
                d.get('data', {}).get('records') or
                d.get('data', {}).get('list') or
                d.get('data') if isinstance(d.get('data'), list) else None or
                d.get('list') or []
            )
            for item in (items or []):
                pid = item.get('productId') or item.get('id') or item.get('product_id')
                if pid:
                    _product_id_cache[token] = str(pid)
                    return str(pid)
        except Exception:
            continue
    return None


@app.route('/api/query', methods=['POST'])
def query_tasks():
    from urllib.parse import urlencode

    body       = request.get_json()
    token      = body.pop('__token__', '')
    task_name  = body.pop('__taskName__', '')
    start_date = body.pop('__startDate__', '')
    end_date   = body.pop('__endDate__', '')

    if not token:
        return jsonify({'ok': False, 'msg': '请先填写 Bearer Token'})

    headers = _get_headers(token)

    # 自动探测 productId；探测失败用已知默认值
    product_id = _detect_product_id(token) or '100083'

    static_params = {
        'productId': product_id,
        'plat':      '4',
        'business':  'RCS',
    }

    all_records = []
    page = 1
    while True:
        params = {
            **static_params,
            'sendTime[]': [start_date, end_date],
            'pageNum':    str(page),
            'pageSize':   '200',
        }
        url = BASE_URL + '?' + urlencode(params, doseq=True)

        try:
            resp = requests.get(url, headers=headers, timeout=30)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            return jsonify({'ok': False, 'msg': str(e)})

        code = data.get('code', data.get('status', 0))
        if code not in (0, 200, '0', '200', None):
            msg = data.get('msg') or data.get('message') or data.get('error') or '接口返回错误'
            return jsonify({'ok': False, 'msg': f'code={code}: {msg}'})

        page_records = _parse_records(data, task_name)
        all_records.extend(page_records)

        d     = data.get('data', {})
        total = d.get('total', 0) if isinstance(d, dict) else 0
        if len(page_records) < 200 or page * 200 >= total:
            break
        page += 1

    return jsonify({'ok': True, 'records': all_records, 'total': len(all_records)})


@app.route('/api/sync-sheets', methods=['POST'])
def sync_sheets():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        return jsonify({'ok': False, 'error': '请先安装: pip install gspread google-auth'})

    try:
        body      = request.get_json()
        records   = body.get('records', [])
        sheet_tab = body.get('sheetTab', '').strip()
        date      = body.get('date', '')

        if not records:   return jsonify({'ok': False, 'error': '没有数据'})
        if not sheet_tab: return jsonify({'ok': False, 'error': '请填写标签页名称'})

        # 读取 Google 服务账号
        creds_json = os.environ.get('GOOGLE_CREDENTIALS')
        if creds_json:
            creds_info = json.loads(creds_json)
        else:
            with open('credentials.json', 'r') as f:
                creds_info = json.load(f)

        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds  = Credentials.from_service_account_info(creds_info, scopes=scopes)
        gc     = gspread.authorize(creds)
        sh     = gc.open_by_key(SHEET_ID)
        ws     = sh.worksheet(sheet_tab)

        # 自动识别「数据来源」列和「条数」列
        header_row  = ws.row_values(1)
        file_col    = None
        tiaoshu_col = None
        for i, h in enumerate(header_row):
            h = h.strip()
            if h == '数据来源':
                file_col = i + 1
            elif h == '条数':
                tiaoshu_col = i + 1

        if file_col is None:
            return jsonify({'ok': False, 'error': '未找到「数据来源」列，请确认第一行表头'})
        if tiaoshu_col is None:
            return jsonify({'ok': False, 'error': '未找到「条数」列，请确认第一行表头'})

        # 按文件名匹配，只更新条数列
        all_vals    = ws.col_values(file_col)
        name_to_row = {v.strip(): i+1 for i, v in enumerate(all_vals) if v.strip()}

        batch   = []
        updated = 0
        skipped = 0
        col_letter = _col_letter(tiaoshu_col)

        for r in records:
            # 去掉 .txt 后缀再匹配
            key = r['fileName']
            if key.lower().endswith('.txt'):
                key = key[:-4]
            row = name_to_row.get(key)
            if row:
                batch.append({'range': f'{col_letter}{row}', 'values': [[r['sendCount']]]})
                updated += 1
            else:
                skipped += 1

        if batch:
            ws.batch_update(batch)
        return jsonify({'ok': True, 'written': updated, 'skipped': skipped})

    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


def _col_letter(n: int) -> str:
    result = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5051))
    print('=' * 50)
    print('  NodeSMS RCS 查询工具已启动')
    print(f'  请在浏览器打开: http://localhost:{port}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False)
