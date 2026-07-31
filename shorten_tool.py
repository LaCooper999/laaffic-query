"""
ShortenWorld 短链抓取 & 同步工具
- 从 Google Sheet Y列读取长链接
- 去 ShortenWorld 查询对应的短链接
- 写回 T列
运行后在浏览器打开: http://localhost:5052
"""

from flask import Flask, request, jsonify, render_template_string
import requests, os, json, time

SHEET_ID  = "1Am5AiKtbjMii0K0nYwNAUXPsNTtDbfaDUER_CY2SRCA"
TEAM_ID   = "6a381c9b2c686733f2d70508"
DOMAIN_ID = "63281bec3d2b0000ee0018c1"

SW_LIST_URL = "https://shortenworld.com/secure/link/list-ajax"

app = Flask(__name__)

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ShortenWorld 短链同步工具</title>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif; background: #f0f2f5; min-height: 100vh; padding: 24px; }
.container { max-width: 860px; margin: 0 auto; }
h1 { font-size: 22px; color: #1a1a2e; margin-bottom: 20px; font-weight: 700; }
.card { background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
label { display: block; font-size: 13px; color: #666; margin-bottom: 6px; font-weight: 500; }
input, textarea, select { width: 100%; border: 1.5px solid #e0e0e0; border-radius: 8px; padding: 10px 14px; font-size: 14px; outline: none; transition: border-color 0.2s; color: #333; background: white; }
input:focus, textarea:focus { border-color: #7c3aed; }
textarea { resize: vertical; min-height: 70px; font-family: monospace; font-size: 12px; }
.row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }
.full { margin-bottom: 16px; }
.btn { border: none; border-radius: 8px; padding: 11px 24px; font-size: 14px; cursor: pointer; font-weight: 600; transition: all 0.2s; }
.btn-primary { background: #7c3aed; color: white; }
.btn-primary:hover { background: #6d28d9; }
.btn-primary:disabled { background: #c4b5fd; cursor: not-allowed; }
.btn-success { background: #059669; color: white; }
.btn-success:hover { background: #047857; }
.btn-warn { background: #f59e0b; color: white; }
.btn-warn:hover { background: #d97706; }
.btn-row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin-bottom: 20px; }
.tip { font-size: 12px; color: #888; margin-top: 5px; line-height: 1.6; }
.section-title { font-size: 13px; font-weight: 600; color: #7c3aed; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 1px solid #ede9fe; }
.log-box { background: #1e1e2e; border-radius: 8px; padding: 16px; font-family: monospace; font-size: 12px; color: #cdd6f4; min-height: 120px; max-height: 360px; overflow-y: auto; white-space: pre-wrap; word-break: break-all; }
.log-ok   { color: #a6e3a1; }
.log-err  { color: #f38ba8; }
.log-info { color: #89b4fa; }
.log-warn { color: #f9e2af; }
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px; }
.stat { border-radius: 8px; padding: 14px; text-align: center; }
.stat.total  { background: #ede9fe; }
.stat.done   { background: #d1fae5; }
.stat.skip   { background: #fef3c7; }
.stat.err    { background: #fee2e2; }
.stat-label  { font-size: 11px; color: #555; margin-bottom: 4px; }
.stat-num    { font-size: 24px; font-weight: 700; }
.stat.total .stat-num { color: #6d28d9; }
.stat.done  .stat-num { color: #065f46; }
.stat.skip  .stat-num { color: #92400e; }
.stat.err   .stat-num { color: #991b1b; }
.progress-bar { width: 100%; height: 6px; background: #e5e7eb; border-radius: 3px; overflow: hidden; margin-bottom: 12px; }
.progress-fill { height: 100%; background: #7c3aed; border-radius: 3px; transition: width 0.3s; }
</style>
</head>
<body>
<div class="container">
  <h1>🔗 ShortenWorld 短链同步工具</h1>

  <div class="card">
    <p class="section-title">🍪 ShortenWorld 认证</p>
    <div class="full">
      <label>Cookie（F12 → Network → 任意请求 → Request headers → Cookie）</label>
      <textarea id="cookie" placeholder="粘贴完整 Cookie 字符串..."></textarea>
      <div class="tip">⚠️ Cookie 会话级有效，重新登录后需重新粘贴</div>
    </div>
  </div>

  <div class="card">
    <p class="section-title">📊 Google Sheet 配置</p>
    <div class="row2">
      <div>
        <label>子表名称</label>
        <input type="text" id="sheetTab" placeholder="例如：730">
      </div>
      <div>
        <label>处理模式</label>
        <select id="mode">
          <option value="skip">跳过 T列已有短链的行</option>
          <option value="overwrite">覆盖所有行</option>
        </select>
      </div>
    </div>
    <div class="tip">💡 读取 Y列长链接 → 在 ShortenWorld 查询对应短链 → 写入 T列</div>
  </div>

  <div class="btn-row">
    <button class="btn btn-primary" onclick="startProcess()" id="startBtn">🚀 开始抓取并同步</button>
    <button class="btn btn-success" onclick="previewSheet()" id="previewBtn">👁 预览 Sheet 数据</button>
    <button class="btn btn-warn" onclick="testCookie()">🔧 测试 Cookie</button>
  </div>

  <div class="card" id="resultCard" style="display:none">
    <div class="stats">
      <div class="stat total"><div class="stat-label">总行数</div><div class="stat-num" id="sTotal">0</div></div>
      <div class="stat done"> <div class="stat-label">已同步</div><div class="stat-num" id="sDone">0</div></div>
      <div class="stat skip"> <div class="stat-label">已跳过</div><div class="stat-num" id="sSkip">0</div></div>
      <div class="stat err">  <div class="stat-label">未找到/失败</div><div class="stat-num" id="sErr">0</div></div>
    </div>
    <div class="progress-bar"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
    <div class="log-box" id="logBox"></div>
  </div>
</div>

<script>
let isRunning = false;

function getCookie() {
  return document.getElementById('cookie').value.trim();
}

function log(msg, type='') {
  const box = document.getElementById('logBox');
  const cls = type ? `log-${type}` : '';
  const ts  = new Date().toTimeString().slice(0,8);
  box.innerHTML += `<span class="${cls}">[${ts}] ${msg}</span>\n`;
  box.scrollTop = box.scrollHeight;
}

function setStats(total, done, skip, err) {
  document.getElementById('sTotal').textContent = total;
  document.getElementById('sDone').textContent  = done;
  document.getElementById('sSkip').textContent  = skip;
  document.getElementById('sErr').textContent   = err;
  const pct = total > 0 ? Math.round((done + skip + err) / total * 100) : 0;
  document.getElementById('progressFill').style.width = pct + '%';
}

async function testCookie() {
  const cookie = getCookie();
  if (!cookie) { alert('请填写 Cookie'); return; }
  document.getElementById('resultCard').style.display = 'block';
  document.getElementById('logBox').innerHTML = '';
  log('正在测试 Cookie...', 'info');
  try {
    const resp = await fetch('/api/test-cookie', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ cookie })
    });
    const data = await resp.json();
    if (data.ok) {
      log(`✅ Cookie 有效！共找到 ${data.total} 条短链记录`, 'ok');
    } else {
      log('❌ Cookie 无效或已过期: ' + data.error, 'err');
    }
  } catch(e) {
    log('❌ 请求失败: ' + e.message, 'err');
  }
}

async function previewSheet() {
  const cookie   = getCookie();
  const sheetTab = document.getElementById('sheetTab').value.trim();
  if (!sheetTab) { alert('请填写子表名称'); return; }
  document.getElementById('resultCard').style.display = 'block';
  document.getElementById('logBox').innerHTML = '';
  log('正在读取 Google Sheet...', 'info');
  try {
    const resp = await fetch('/api/read-sheet', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sheetTab, urlCol: 'Y', shortCol: 'T', startRow: 3 })
    });
    const data = await resp.json();
    if (!data.ok) { log('❌ 读取失败: ' + data.error, 'err'); return; }
    log(`✅ 读取成功，共 ${data.rows.length} 行`, 'ok');
    const hasUrl   = data.rows.filter(r => r.url).length;
    const hasShort = data.rows.filter(r => r.short).length;
    log(`   有长链接: ${hasUrl} 行，已有短链: ${hasShort} 行`, 'info');
    data.rows.slice(0, 8).forEach(r => {
      const s = r.short ? `✅ ${r.short}` : '⬜ 无短链';
      log(`   行${r.row}: ${(r.url||'').slice(0,55)}... [${s}]`);
    });
    if (data.rows.length > 8) log(`   ... 还有 ${data.rows.length - 8} 行`);
    setStats(data.rows.length, 0, 0, 0);
  } catch(e) {
    log('❌ ' + e.message, 'err');
  }
}

async function startProcess() {
  if (isRunning) return;
  const cookie   = getCookie();
  const sheetTab = document.getElementById('sheetTab').value.trim();
  const mode     = document.getElementById('mode').value;

  if (!cookie)   { alert('请填写 Cookie'); return; }
  if (!sheetTab) { alert('请填写子表名称'); return; }

  isRunning = true;
  document.getElementById('startBtn').disabled = true;
  document.getElementById('startBtn').textContent = '处理中...';
  document.getElementById('resultCard').style.display = 'block';
  document.getElementById('logBox').innerHTML = '';

  log('开始处理...', 'info');

  // 第一步：读取 Sheet
  let rows = [];
  try {
    const resp = await fetch('/api/read-sheet', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sheetTab, urlCol: 'Y', shortCol: 'T', startRow: 3 })
    });
    const data = await resp.json();
    if (!data.ok) { log('❌ 读取 Sheet 失败: ' + data.error, 'err'); return; }
    rows = data.rows;
    log(`✅ 读取 Sheet 完成，共 ${rows.length} 行`, 'ok');
  } catch(e) {
    log('❌ ' + e.message, 'err'); return;
  } finally {
    if (!rows.length) {
      isRunning = false;
      document.getElementById('startBtn').disabled = false;
      document.getElementById('startBtn').textContent = '🚀 开始抓取并同步';
      return;
    }
  }

  // 第二步：从 ShortenWorld 获取所有短链（分页）
  log('正在从 ShortenWorld 获取短链列表...', 'info');
  let linkMap = {};
  try {
    const resp = await fetch('/api/fetch-links', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ cookie })
    });
    const data = await resp.json();
    if (!data.ok) { log('❌ 获取短链失败: ' + data.error, 'err'); return; }
    linkMap = data.linkMap;
    log(`✅ 获取到 ${Object.keys(linkMap).length} 条短链记录`, 'ok');
  } catch(e) {
    log('❌ ' + e.message, 'err'); return;
  }

  // 第三步：逐行匹配 + 写入
  setStats(rows.length, 0, 0, 0);
  let done = 0, skip = 0, err = 0;
  const writeQueue = [];

  for (const r of rows) {
    if (!r.url) { skip++; setStats(rows.length, done, skip, err); continue; }
    if (mode === 'skip' && r.short) {
      log(`  行${r.row}: 已有短链，跳过`, 'warn');
      skip++; setStats(rows.length, done, skip, err); continue;
    }

    // 匹配长链接
    const shortUrl = linkMap[r.url] || linkMap[r.url.replace(/\/$/, '')] || null;
    if (shortUrl) {
      writeQueue.push({ row: r.row, shortUrl });
      log(`  行${r.row}: ✅ ${shortUrl}`, 'ok');
      done++;
    } else {
      log(`  行${r.row}: ⚠️ 未找到对应短链`, 'warn');
      err++;
    }
    setStats(rows.length, done, skip, err);

    // 每10条写入一次
    if (writeQueue.length >= 10) {
      await flushWrite(writeQueue, sheetTab);
    }
  }

  if (writeQueue.length > 0) await flushWrite(writeQueue, sheetTab);

  log(`\n🎉 完成！已同步 ${done} 条，跳过 ${skip} 条，未找到 ${err} 条`, 'ok');
  isRunning = false;
  document.getElementById('startBtn').disabled = false;
  document.getElementById('startBtn').textContent = '🚀 开始抓取并同步';
}

async function flushWrite(queue, sheetTab) {
  const batch = queue.splice(0, queue.length);
  try {
    const resp = await fetch('/api/write-sheet', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sheetTab, shortCol: 'T', batch })
    });
    const data = await resp.json();
    if (data.ok) {
      log(`  📝 已写入 ${batch.length} 条到 Sheet`, 'info');
    } else {
      log(`  ❌ 写入失败: ${data.error}`, 'err');
      batch.forEach(b => queue.push(b));
    }
  } catch(e) {
    log(`  ❌ 写入异常: ${e.message}`, 'err');
  }
}
</script>
</body>
</html>
"""


def _get_sheet(sheet_tab: str):
    import gspread
    from google.oauth2.service_account import Credentials
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if creds_json:
        creds_info = json.loads(creds_json)
    else:
        with open('credentials.json', 'r') as f:
            creds_info = json.load(f)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds  = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc     = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID).worksheet(sheet_tab)


def col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def _get_csrf(cookie: str) -> str:
    import re
    headers = {
        'Cookie':     cookie,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Referer':    'https://shortenworld.com/',
    }
    resp = requests.get('https://shortenworld.com/secure/link/list', headers=headers, timeout=15)
    pat = re.compile(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}')
    idx = resp.text.find('_csrf')
    if idx != -1:
        m = pat.search(resp.text[idx:idx+200])
        if m:
            return m.group(0)
    m = pat.search(resp.text)
    if m:
        return m.group(0)
    raise Exception('无法获取 _csrf，请检查 Cookie 是否有效')


def _sw_headers(cookie: str) -> dict:
    return {
        'Cookie':           cookie,
        'Accept':           'application/json, text/javascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer':          'https://shortenworld.com/secure/link/list',
        'User-Agent':       'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }


@app.route('/')
def index():
    return render_template_string(HTML_PAGE)


@app.route('/api/test-cookie', methods=['POST'])
def test_cookie():
    body   = request.get_json()
    cookie = body.get('cookie', '')
    try:
        headers = _sw_headers(cookie)
        csrf = _get_csrf(cookie)
        params = {
            'draw': 1, 'start': 0, 'length': 1,
            'search[value]': '', 'search[regex]': 'false',
            '_csrf': csrf,
        }
        resp = requests.post(SW_LIST_URL, data=params,
            headers={**headers, 'Content-Type': 'application/x-www-form-urlencoded'},
            timeout=15)
        data = resp.json()
        total = int(data.get('recordsTotal') or data.get('total') or len(data.get('data', [])))
        return jsonify({'ok': True, 'total': total})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/api/fetch-links', methods=['POST'])
def fetch_links():
    """从 ShortenWorld 拉取所有短链，返回 {长链接: 短链接} 映射"""
    body   = request.get_json()
    cookie = body.get('cookie', '')
    headers = _sw_headers(cookie)

    link_map = {}
    page = 0
    page_size = 100

    try:
        while True:
            csrf = _get_csrf(cookie)
            params = {
                'draw':          page + 1,
                'start':         page * page_size,
                'length':        page_size,
                'search[value]': '',
                'search[regex]': 'false',
                '_csrf':         csrf,
            }
            for i in range(8):
                params[f'columns[{i}][searchable]']    = 'true'
                params[f'columns[{i}][orderable]']     = 'false'
                params[f'columns[{i}][search][value]'] = ''
                params[f'columns[{i}][search][regex]'] = 'false'

            resp = requests.post(
                SW_LIST_URL,
                data=params,
                headers={**headers, 'Content-Type': 'application/x-www-form-urlencoded'},
                timeout=20
            )
            data = resp.json()

            # 兼容列表或字典响应
            if isinstance(data, list):
                records = data
                total   = len(data) if page == 0 else 0
            else:
                records = (
                    data.get('data') or
                    data.get('links') or
                    data.get('records') or []
                )
                total = int(data.get('recordsTotal') or data.get('total') or 0)

            if not records:
                break

            for lk in records:
                # data 是数组：[序号, 短链, 长链接, ...]
                if isinstance(lk, list):
                    short = str(lk[1] if len(lk) > 1 else '').strip()
                    dest  = str(lk[2] if len(lk) > 2 else '').strip()
                else:
                    dest  = (lk.get('destination') or lk.get('dest') or '').strip()
                    short = (lk.get('url') or lk.get('shortUrl') or lk.get('short') or '').strip()
                short = short.replace('https://', '').replace('http://', '')
                if dest and short:
                    link_map[dest] = short
                    link_map[dest.rstrip('/')] = short

            page += 1
            # recordsTotal=100000000 是占位值，以实际返回条数判断是否还有更多
            if not records or len(records) < page_size:
                break

        return jsonify({'ok': True, 'linkMap': link_map, 'count': len(link_map)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/api/read-sheet', methods=['POST'])
def read_sheet():
    try:
        body      = request.get_json()
        sheet_tab = body.get('sheetTab', '').strip()
        url_col   = body.get('urlCol', 'Y').strip().upper()
        short_col = body.get('shortCol', 'T').strip().upper()
        start_row = int(body.get('startRow', 3))

        ws        = _get_sheet(sheet_tab)
        url_idx   = col_letter_to_index(url_col)
        short_idx = col_letter_to_index(short_col)

        all_url   = ws.col_values(url_idx)
        all_short = ws.col_values(short_idx)

        rows = []
        for i in range(start_row - 1, len(all_url)):
            url   = (all_url[i]   if i < len(all_url)   else '').strip()
            short = (all_short[i] if i < len(all_short) else '').strip()
            if not url:
                continue
            rows.append({'row': i + 1, 'url': url, 'short': short})

        return jsonify({'ok': True, 'rows': rows})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


@app.route('/api/write-sheet', methods=['POST'])
def write_sheet():
    try:
        body      = request.get_json()
        sheet_tab = body.get('sheetTab', '').strip()
        short_col = body.get('shortCol', 'T').strip().upper()
        batch     = body.get('batch', [])

        if not batch:
            return jsonify({'ok': True, 'written': 0})

        ws      = _get_sheet(sheet_tab)
        updates = [
            {'range': f'{short_col}{item["row"]}', 'values': [[item['shortUrl']]]}
            for item in batch
        ]
        ws.batch_update(updates)
        return jsonify({'ok': True, 'written': len(batch)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5052))
    print('=' * 50)
    print('  ShortenWorld 短链同步工具已启动')
    print(f'  请在浏览器打开: http://localhost:{port}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False)
