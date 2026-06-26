"""
Laaffic 任务查询工具
运行后在浏览器打开: http://localhost:5050
"""
from flask import Flask, request, jsonify, render_template_string
import requests

app = Flask(__name__)

HTML_PAGE = r"""
<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Laaffic 任务查询</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif; background: #f0f2f5; min-height: 100vh; padding: 24px; }
.container { max-width: 1000px; margin: 0 auto; }
h1 { font-size: 22px; color: #1a1a2e; margin-bottom: 20px; font-weight: 700; }
.card { background: white; border-radius: 12px; padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
label { display: block; font-size: 13px; color: #666; margin-bottom: 6px; font-weight: 500; }
input, textarea { width: 100%; border: 1.5px solid #e0e0e0; border-radius: 8px; padding: 10px 14px; font-size: 14px; outline: none; transition: border-color 0.2s; color: #333; }
input:focus, textarea:focus { border-color: #4f46e5; }
textarea { resize: vertical; min-height: 90px; font-family: monospace; font-size: 12px; }
.row3 { display: grid; grid-template-columns: 2fr 1fr 1fr; gap: 16px; margin-bottom: 16px; }
.full { margin-bottom: 16px; }
button { background: #4f46e5; color: white; border: none; border-radius: 8px; padding: 12px 32px; font-size: 15px; cursor: pointer; font-weight: 600; transition: background 0.2s; }
button:hover { background: #4338ca; }
button:disabled { background: #a5b4fc; cursor: not-allowed; }
.stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 20px; }
.stat-card { border-radius: 10px; padding: 18px; text-align: center; }
.stat-card.sms { background: #ede9fe; }
.stat-card.connected { background: #d1fae5; }
.stat-card.tasks { background: #fef3c7; }
.stat-label { font-size: 12px; color: #555; margin-bottom: 8px; }
.stat-num { font-size: 30px; font-weight: 700; }
.stat-card.sms .stat-num { color: #5b21b6; }
.stat-card.connected .stat-num { color: #065f46; }
.stat-card.tasks .stat-num { color: #92400e; }
table { width: 100%; border-collapse: collapse; }
th { background: #f8f9ff; color: #4f46e5; font-size: 13px; padding: 12px 16px; text-align: left; border-bottom: 2px solid #e8e8f0; white-space: nowrap; }
td { padding: 11px 16px; border-bottom: 1px solid #f0f0f0; font-size: 14px; color: #333; }
tr:hover td { background: #fafafe; }
.num-sms { font-weight: 700; color: #5b21b6; }
.num-conn { font-weight: 700; color: #059669; }
.empty { text-align: center; color: #999; padding: 48px; font-size: 15px; }
.loading { text-align: center; color: #4f46e5; padding: 24px; font-size: 15px; }
.error-box { color: #dc2626; padding: 12px 16px; background: #fef2f2; border-radius: 8px; margin-top: 14px; border: 1px solid #fecaca; }
.tip { font-size: 12px; color: #888; margin-top: 6px; }
.table-wrap { overflow-x: auto; }
.btn-export { background: #059669; color: white; border: none; border-radius: 8px; padding: 10px 24px; font-size: 14px; cursor: pointer; font-weight: 600; margin-left: 12px; transition: background 0.2s; }
.btn-export:hover { background: #047857; }
.btn-row { display: flex; align-items: center; margin-bottom: 16px; }
</style>
</head>
<body>
<div class="container">
  <h1>📞 Laaffic 任务查询</h1>

  <div class="card">
    <div class="full">
      <label>🔑 Bearer Token（从浏览器开发者工具 → 网络 → 任意请求 → 标头 → Authorization 复制）</label>
      <textarea id="token" placeholder="粘贴 Bearer Token（带不带 Bearer 前缀都行）..."></textarea>
      <div class="tip">Token 会话级有效，重新登录后需重新粘贴</div>
    </div>
    <div class="row3">
      <div>
        <label>任务名称（模糊搜索，留空显示全部）</label>
        <input type="text" id="taskName" placeholder="输入任务名称关键词...">
      </div>
      <div>
        <label>开始时间</label>
        <input type="datetime-local" id="strTime">
      </div>
      <div>
        <label>结束时间</label>
        <input type="datetime-local" id="endTime">
      </div>
    </div>
    <button onclick="fetchData()" id="btn">🔍 查 询</button>
    <div id="error"></div>
  </div>

  <div class="card" id="result" style="display:none">
    <div id="content"></div>
  </div>
</div>

<script>
const now = new Date();
const ago = new Date(now); ago.setDate(ago.getDate() - 90);
document.getElementById('endTime').value = fmt(now);
document.getElementById('strTime').value = fmt(ago);

function fmt(d) {
  const p = n => String(n).padStart(2,'0');
  return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())+'T'+p(d.getHours())+':'+p(d.getMinutes());
}
function toApiTime(v) {
  if (!v) return '';
  return v.replace('T', ' ') + ':00';
}

async function fetchData() {
  const rawToken = document.getElementById('token').value.trim();
  if (!rawToken) { showError('请先填写 Bearer Token'); return; }
  const token = rawToken.startsWith('Bearer ') ? rawToken.slice(7) : rawToken;

  const taskNameFilter = document.getElementById('taskName').value.trim();
  const strTimeVal = document.getElementById('strTime').value;
  const endTimeVal = document.getElementById('endTime').value;

  const btn = document.getElementById('btn');
  btn.disabled = true; btn.textContent = '查询中...';
  document.getElementById('error').innerHTML = '';
  document.getElementById('result').style.display = 'block';
  document.getElementById('content').innerHTML = '<div class="loading">⏳ 正在加载数据...</div>';

  try {
    const body = {
      __token__: token,
      __taskName__: taskNameFilter,
      current: 1,
      size: 200,
      params: {
        strTime: toApiTime(strTimeVal),
        endTime: toApiTime(endTimeVal),
        statusList: [],
        taskType: [],
        timezone: "UTC + 4",
        jobIdList: [],
        taskReCallQuerySwitch: false
      }
    };

    const resp = await fetch('/api/query', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const data = await resp.json();

    if (data.code !== 0) { showError('接口错误: ' + (data.msg || '未知错误')); return; }
    window._lastRecords = data.data && data.data.records ? data.data.records : [];
    renderTable(window._lastRecords);
  } catch(e) {
    showError('请求失败: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '🔍 查 询';
  }
}

function renderTable(records) {
  if (!records.length) {
    document.getElementById('content').innerHTML = '<div class="empty">🔍 没有找到匹配的任务</div>';
    return;
  }

  let totalSms = 0, totalConn = 0;
  records.forEach(function(r) { totalSms += r.sendSmsNum || 0; totalConn += r.success || 0; });

  let html = '<div class="btn-row">'
    + '<span style="font-size:14px;color:#333;font-weight:600">共 ' + records.length + ' 条任务</span>'
    + '<button class="btn-export" onclick="exportExcel()">📥 导出 Excel</button>'
    + '</div>'
    + '<div class="stats">'
    + '<div class="stat-card sms"><div class="stat-label">挂机短信发送数（合计）</div><div class="stat-num">' + totalSms.toLocaleString() + '</div></div>'
    + '<div class="stat-card connected"><div class="stat-label">接通数（合计）</div><div class="stat-num">' + totalConn.toLocaleString() + '</div></div>'
    + '<div class="stat-card tasks"><div class="stat-label">任务数量</div><div class="stat-num">' + records.length + '</div></div>'
    + '</div>'
    + '<div class="table-wrap"><table><thead><tr>'
    + '<th>#</th><th>任务名称</th><th>接通数</th><th>挂机短信发送数</th><th>总拨打数</th><th>接通率</th><th>发送时间</th>'
    + '</tr></thead><tbody>';

  records.forEach(function(r, i) {
    html += '<tr>'
      + '<td style="color:#999">' + (i+1) + '</td>'
      + '<td>' + (r.taskName || '-') + '</td>'
      + '<td class="num-conn">' + (r.success || 0).toLocaleString() + '</td>'
      + '<td class="num-sms">' + (r.sendSmsNum || 0).toLocaleString() + '</td>'
      + '<td>' + (r.callNum || 0).toLocaleString() + '</td>'
      + '<td>' + (r.successRate ? r.successRate + '%' : '-') + '</td>'
      + '<td style="font-size:12px;color:#999">' + (r.sendTime || '-') + '</td>'
      + '</tr>';
  });

  html += '</tbody></table></div>';
  document.getElementById('content').innerHTML = html;
}

function exportExcel() {
  const records = window._lastRecords || [];
  if (!records.length) return;

  const rows = [['任务名称', '接通数', '挂机短信发送数', '总拨打数', '接通率(%)', '发送时间']];
  records.forEach(function(r) {
    rows.push([
      r.taskName || '',
      r.success || 0,
      r.sendSmsNum || 0,
      r.callNum || 0,
      r.successRate || '',
      r.sendTime || ''
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);

  // 列宽
  ws['!cols'] = [
    {wch: 40}, {wch: 10}, {wch: 16}, {wch: 10}, {wch: 10}, {wch: 20}
  ];

  XLSX.utils.book_append_sheet(wb, ws, '任务数据');

  const now = new Date();
  const p = n => String(n).padStart(2,'0');
  const filename = 'laaffic_' + now.getFullYear() + p(now.getMonth()+1) + p(now.getDate()) + '_' + p(now.getHours()) + p(now.getMinutes()) + '.xlsx';
  XLSX.writeFile(wb, filename);
}

function showError(msg) {
  document.getElementById('error').innerHTML = '<div class="error-box">❌ ' + msg + '</div>';
  document.getElementById('result').style.display = 'none';
}
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/api/query', methods=['POST'])
def query_tasks():
    body = request.get_json()
    token = body.pop('__token__', '')
    task_name = body.pop('__taskName__', '')

    headers = {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
        'Origin': 'https://my.laaffic.com',
        'Referer': 'https://my.laaffic.com/',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    }

    try:
        resp = requests.post(
            'https://gw.onbuka.com/voice/voice/group/call/query',
            json=body,
            headers=headers,
            timeout=30
        )
        data = resp.json()
    except Exception as e:
        return jsonify({'code': -1, 'msg': str(e)}), 500

    if task_name and data.get('code') == 0:
        records = data.get('data', {}).get('records', [])
        data['data']['records'] = [
            r for r in records
            if task_name.lower() in (r.get('taskName') or '').lower()
        ]

    return jsonify(data)

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5050))
    print('=' * 50)
    print('  Laaffic 任务查询工具已启动')
    print(f'  请在浏览器打开: http://localhost:{port}')
    print('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False)
