"""
天天基金涨跌幅前100排名查询 - Web 应用
部署方式（三选一）：
  1. Render.com（免费，海外）: 直接推送 Git 仓库即可
  2. PythonAnywhere（免费）: 上传文件后在 Web 面板配置
  3. 自建服务器: python app.py 启动
"""

import re
import json
import requests
from flask import Flask, request, jsonify, send_from_directory
from datetime import datetime, timedelta

app = Flask(__name__)

API_URL = "https://fund.eastmoney.com/data/rankhandler.aspx"

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>天天基金涨跌幅前100排名查询</title>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;background:#f0f2f5;color:#333;min-height:100vh}
.header{background:linear-gradient(135deg,#1a73e8 0%,#1557b0 100%);color:#fff;padding:24px 0;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.15)}
.header h1{font-size:26px;font-weight:600;letter-spacing:1px}
.subtitle{font-size:13px;opacity:.85;margin-top:6px}
.controls{display:flex;align-items:center;justify-content:center;gap:14px;padding:20px;background:#fff;margin:20px auto 0;max-width:900px;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);flex-wrap:wrap}
.controls label{font-weight:600;font-size:14px}
.controls input[type="date"]{padding:8px 14px;border:1px solid #d0d5dd;border-radius:6px;font-size:14px;outline:none;transition:border-color .2s}
.controls input[type="date"]:focus{border-color:#1a73e8}
.btn{padding:9px 22px;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;transition:all .2s;letter-spacing:.5px}
.btn-query{background:#1a73e8;color:#fff}
.btn-query:hover{background:#1557b0}
.btn-query:disabled{background:#93b8f0;cursor:not-allowed}
.btn-excel{background:#fff;color:#217346;border:2px solid #217346}
.btn-excel:hover{background:#e8f5e9}
.btn-excel:disabled{color:#999;border-color:#ccc;cursor:not-allowed}
.status{text-align:center;margin-top:12px;font-size:13px;color:#888;min-height:20px}
.status.error{color:#d93025}
.status.success{color:#1e8e3e}
.content{display:flex;gap:20px;max-width:1400px;margin:20px auto;padding:0 20px;flex-wrap:wrap}
.table-section{flex:1;min-width:500px;background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.table-section h2{font-size:17px;padding:14px 18px;border-bottom:1px solid #eee;display:flex;align-items:center;gap:8px}
.badge{font-size:12px;padding:2px 10px;border-radius:12px;color:#fff}
.badge-up{background:#e53935}
.badge-down{background:#43a047}
.table-wrap{overflow-x:auto;max-height:600px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead{position:sticky;top:0;z-index:1}
thead th{background:#f8f9fa;padding:10px 12px;text-align:left;border-bottom:2px solid #e0e0e0;font-weight:600;white-space:nowrap}
tbody td{padding:8px 12px;border-bottom:1px solid #f0f0f0;white-space:nowrap}
tbody tr:hover{background:#f5f8ff}
.growth-up{color:#e53935;font-weight:600}
.growth-down{color:#43a047;font-weight:600}
.empty-state{text-align:center;padding:60px 20px;color:#999}
.footer{text-align:center;padding:20px;color:#999;font-size:12px}
@media(max-width:1100px){.content{flex-direction:column}.table-section{min-width:auto}}
</style>
</head>
<body>

<div class="header">
  <h1>天天基金涨跌幅前100排名查询</h1>
  <div class="subtitle">数据来源：天天基金网（fund.eastmoney.com）</div>
</div>

<div class="controls">
  <label for="queryDate">查询日期：</label>
  <input type="date" id="queryDate" />
  <button class="btn btn-query" id="btnQuery" onclick="queryData()">查询</button>
  <button class="btn btn-excel" id="btnExcel" onclick="downloadExcel()" disabled>下载 Excel</button>
</div>
<div class="status" id="status"></div>

<div class="content">
  <div class="table-section">
    <h2><span class="badge badge-up">涨幅</span> 涨幅前100名</h2>
    <div class="table-wrap">
      <table id="gainersTable">
        <thead><tr><th>排名</th><th>基金代码</th><th>基金名称</th><th>净值日期</th><th>单位净值</th><th>累计净值</th><th>日增长率</th></tr></thead>
        <tbody><tr class="empty-state"><td colspan="7">点击"查询"获取数据</td></tr></tbody>
      </table>
    </div>
  </div>
  <div class="table-section">
    <h2><span class="badge badge-down">跌幅</span> 跌幅前100名</h2>
    <div class="table-wrap">
      <table id="losersTable">
        <thead><tr><th>排名</th><th>基金代码</th><th>基金名称</th><th>净值日期</th><th>单位净值</th><th>累计净值</th><th>日增长率</th></tr></thead>
        <tbody><tr class="empty-state"><td colspan="7">点击"查询"获取数据</td></tr></tbody>
      </table>
    </div>
  </div>
</div>

<div class="footer">数据仅供参考，不构成投资建议 | 数据来源：天天基金网</div>

<script>
var lastGainers=null,lastLosers=null,lastQueryDate='';

function $(id){return document.getElementById(id)}

function setStatus(msg,cls){
  var el=$('status');
  el.textContent=msg;
  el.className='status '+(cls||'');
}

function parseFunds(rawDatas){
  if(!rawDatas||!Array.isArray(rawDatas))return[];
  return rawDatas.map(function(item){
    var f=item.split(',');
    if(f.length<7||isNaN(parseFloat(f[6])))return null;
    return{code:f[0],name:f[1],date:f[3],unitNav:f[4],cumNav:f[5],dayGrowth:parseFloat(f[6])};
  }).filter(Boolean);
}

function renderTable(tableId,funds,sortDesc){
  var tbody=document.querySelector('#'+tableId+' tbody');
  if(!funds.length){tbody.innerHTML='<tr class="empty-state"><td colspan="7">暂无数据</td></tr>';return}
  var sorted=funds.slice().sort(function(a,b){return sortDesc?b.dayGrowth-a.dayGrowth:a.dayGrowth-b.dayGrowth});
  tbody.innerHTML=sorted.map(function(f,i){
    var cls=f.dayGrowth>=0?'growth-up':'growth-down';
    var sign=f.dayGrowth>0?'+':'';
    return '<tr><td>'+(i+1)+'</td><td>'+f.code+'</td><td>'+f.name+'</td><td>'+f.date+'</td><td>'+f.unitNav+'</td><td>'+f.cumNav+'</td><td class="'+cls+'">'+sign+f.dayGrowth.toFixed(2)+'%</td></tr>';
  }).join('');
}

async function fetchRanking(sortOrder,date){
  var d=new Date(date);d.setFullYear(d.getFullYear()-1);
  var sd=d.toISOString().split('T')[0];
  var params='op=ph&dt=kf&ft=all&sc=srzdf&st='+sortOrder+'&sd='+sd+'&ed='+date+'&pi=1&pn=100&dx=1';
  var resp=await fetch('/api?'+params);
  if(!resp.ok)throw new Error('服务器返回错误: '+resp.status);
  var data=await resp.json();
  return data.datas||[];
}

async function queryData(){
  var date=$('queryDate').value;
  if(!date){setStatus('请选择查询日期','error');return}
  lastQueryDate=date;lastGainers=null;lastLosers=null;
  $('btnQuery').disabled=true;$('btnExcel').disabled=true;
  setStatus('正在查询中...','');
  try{
    setStatus('正在抓取涨幅前100名...','');
    lastGainers=parseFunds(await fetchRanking('desc',date));
    renderTable('gainersTable',lastGainers,true);
    setStatus('正在抓取跌幅前100名...','');
    lastLosers=parseFunds(await fetchRanking('asc',date));
    renderTable('losersTable',lastLosers,false);
    if(lastGainers.length||lastLosers.length){
      var d=(lastGainers[0]||lastLosers[0]).date;
      setStatus('查询完成！涨幅前'+lastGainers.length+'名 / 跌幅前'+lastLosers.length+'名（净值日期：'+d+'）','success');
      $('btnExcel').disabled=false;
    }else{setStatus('该日期暂无数据，请尝试其他日期','error')}
  }catch(e){setStatus('查询失败：'+e.message,'error')}
  $('btnQuery').disabled=false;
}

function downloadExcel(){
  if(!lastGainers&&!lastLosers)return;
  var dateStr=lastQueryDate.replace(/-/g,'');
  var wb=XLSX.utils.book_new();
  var headers=['排名','基金代码','基金名称','净值日期','单位净值','累计净值','日增长率(%)'];
  function toSheet(funds,sortDesc){
    var sorted=funds.slice().sort(function(a,b){return sortDesc?b.dayGrowth-a.dayGrowth:a.dayGrowth-b.dayGrowth});
    var rows=[headers];
    sorted.forEach(function(f,i){
      rows.push([i+1,f.code,f.name,f.date,parseFloat(f.unitNav),parseFloat(f.cumNav),parseFloat(f.dayGrowth.toFixed(2))]);
    });
    return XLSX.utils.aoa_to_sheet(rows);
  }
  if(lastGainers&&lastGainers.length)XLSX.utils.book_append_sheet(wb,toSheet(lastGainers,true),'涨幅前100');
  if(lastLosers&&lastLosers.length)XLSX.utils.book_append_sheet(wb,toSheet(lastLosers,false),'跌幅前100');
  ['涨幅前100','跌幅前100'].forEach(function(name){
    if(!wb.Sheets[name])return;
    var sheet=wb.Sheets[name];
    var cols=[];
    for(var c=0;c<7;c++){
      var maxW=8;
      for(var r=0;r<101;r++){
        var cell=sheet[XLSX.utils.encode_cell({r:r,c:c})];
        if(cell&&cell.v!=null)maxW=Math.max(maxW,String(cell.v).length*2);
      }
      cols.push({wch:Math.min(maxW+4,40)});
    }
    sheet['!cols']=cols;
  });
  XLSX.writeFile(wb,'天天基金涨跌幅前100排名查询-'+dateStr+'.xlsx');
}

(function(){
  var t=new Date();
  $('queryDate').value=t.getFullYear()+'-'+String(t.getMonth()+1).padStart(2,'0')+'-'+String(t.getDate()).padStart(2,'0');
})();
</script>
</body>
</html>"""


@app.route("/")
def index():
    return HTML_PAGE


@app.route("/api")
def api_proxy():
    try:
        query_string = request.query_string.decode("utf-8")
        url = f"{API_URL}?{query_string}"
        resp = requests.get(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Referer": "https://fund.eastmoney.com/data/fundranking.html",
        }, timeout=15)
        resp.encoding = "utf-8"

        match = re.search(r'(\bdatas)\s*:\s*(\[.*?\])\s*,\s*(\ballRecords)', resp.text, re.DOTALL)
        if not match:
            return jsonify({"error": "解析数据失败"}), 500

        datas = json.loads(match.group(2))
        return jsonify({"datas": datas})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8765, debug=False)
