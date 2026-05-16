import json
import re
from datetime import datetime

import requests
from flask import Flask, jsonify, request


app = Flask(__name__)

EASTMONEY_API = "https://fund.eastmoney.com/data/rankhandler.aspx"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer": "https://fund.eastmoney.com/data/fundranking.html",
    "Accept": "*/*",
}


HTML_PAGE = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>天天基金涨跌幅前100排名查询</title>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
  <style>
    *{box-sizing:border-box}
    body{margin:0;background:#f4f6f8;color:#1f2937;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",Arial,sans-serif}
    header{background:#1455a8;color:white;padding:24px 20px;text-align:center}
    h1{margin:0;font-size:26px;font-weight:650}
    .sub{margin-top:8px;font-size:13px;opacity:.86}
    .controls{max-width:980px;margin:18px auto 0;background:white;border:1px solid #e5e7eb;border-radius:8px;padding:16px;display:flex;gap:12px;align-items:center;justify-content:center;flex-wrap:wrap}
    label{font-weight:650}
    input{height:38px;border:1px solid #cbd5e1;border-radius:6px;padding:0 12px;font-size:14px}
    button{height:38px;border:0;border-radius:6px;padding:0 18px;font-size:14px;font-weight:650;cursor:pointer}
    button:disabled{cursor:not-allowed;opacity:.55}
    .primary{background:#1769d1;color:white}
    .excel{background:white;color:#16733c;border:1px solid #16733c}
    #status{text-align:center;min-height:22px;margin:12px auto;color:#64748b;font-size:13px}
    #status.ok{color:#16733c}
    #status.err{color:#c62828}
    main{max-width:1440px;margin:0 auto;padding:0 18px 24px;display:grid;grid-template-columns:1fr 1fr;gap:18px}
    section{background:white;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden}
    h2{font-size:17px;margin:0;padding:13px 16px;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;gap:8px}
    .badge{display:inline-block;color:white;border-radius:999px;padding:2px 10px;font-size:12px}
    .up-badge{background:#d93025}
    .down-badge{background:#188038}
    .table-wrap{overflow:auto;max-height:650px}
    table{width:100%;border-collapse:collapse;font-size:13px}
    th,td{padding:9px 10px;border-bottom:1px solid #eef2f7;text-align:left;white-space:nowrap}
    th{position:sticky;top:0;background:#f8fafc;z-index:1;font-weight:700}
    tbody tr:hover{background:#f8fbff}
    .up{color:#d93025;font-weight:700}
    .down{color:#188038;font-weight:700}
    .empty{text-align:center;color:#94a3b8;padding:42px 10px}
    footer{text-align:center;color:#94a3b8;font-size:12px;padding:18px}
    @media(max-width:1100px){main{grid-template-columns:1fr}}
  </style>
</head>
<body>
  <header>
    <h1>天天基金涨跌幅前100排名查询</h1>
    <div class="sub">数据来源：天天基金网。非交易日会自动取该日期前最近可用净值。</div>
  </header>

  <div class="controls">
    <label for="queryDate">查询日期</label>
    <input id="queryDate" type="date">
    <button id="queryBtn" class="primary" onclick="queryData()">查询</button>
    <button id="excelBtn" class="excel" onclick="downloadExcel()" disabled>下载 Excel</button>
  </div>
  <div id="status"></div>

  <main>
    <section>
      <h2><span class="badge up-badge">涨幅</span>涨幅前100</h2>
      <div class="table-wrap">
        <table id="gainers">
          <thead><tr><th>排名</th><th>基金代码</th><th>基金名称</th><th>净值日期</th><th>单位净值</th><th>累计净值</th><th>日增长率</th></tr></thead>
          <tbody><tr><td class="empty" colspan="7">请选择日期后查询</td></tr></tbody>
        </table>
      </div>
    </section>

    <section>
      <h2><span class="badge down-badge">跌幅</span>跌幅前100</h2>
      <div class="table-wrap">
        <table id="losers">
          <thead><tr><th>排名</th><th>基金代码</th><th>基金名称</th><th>净值日期</th><th>单位净值</th><th>累计净值</th><th>日增长率</th></tr></thead>
          <tbody><tr><td class="empty" colspan="7">请选择日期后查询</td></tr></tbody>
        </table>
      </div>
    </section>
  </main>

  <footer>数据仅供参考，不构成投资建议。</footer>

<script>
let gainers = [];
let losers = [];
let currentDate = "";

const $ = (id) => document.getElementById(id);

function setStatus(text, type="") {
  const el = $("status");
  el.textContent = text;
  el.className = type;
}

function fmtPct(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "";
  return (n > 0 ? "+" : "") + n.toFixed(2) + "%";
}

function renderTable(id, rows) {
  const tbody = document.querySelector(`#${id} tbody`);
  if (!rows.length) {
    tbody.innerHTML = '<tr><td class="empty" colspan="7">暂无数据</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((f, i) => {
    const cls = f.day_growth_rate >= 0 ? "up" : "down";
    return `<tr>
      <td>${i + 1}</td>
      <td>${f.fund_code}</td>
      <td>${f.fund_name}</td>
      <td>${f.net_value_date}</td>
      <td>${f.unit_net}</td>
      <td>${f.cumulative_net}</td>
      <td class="${cls}">${fmtPct(f.day_growth_rate)}</td>
    </tr>`;
  }).join("");
}

async function loadRanking(sort) {
  const resp = await fetch(`/api/ranking?date=${encodeURIComponent(currentDate)}&sort=${sort}`);
  const data = await resp.json();
  if (!resp.ok) throw new Error(data.error || "查询失败");
  return data.funds || [];
}

async function queryData() {
  currentDate = $("queryDate").value;
  if (!currentDate) {
    setStatus("请选择查询日期", "err");
    return;
  }
  $("queryBtn").disabled = true;
  $("excelBtn").disabled = true;
  try {
    setStatus("正在查询涨幅榜...");
    gainers = await loadRanking("desc");
    renderTable("gainers", gainers);

    setStatus("正在查询跌幅榜...");
    losers = await loadRanking("asc");
    renderTable("losers", losers);

    if (!gainers.length && !losers.length) {
      setStatus("该日期暂无数据，请换一个交易日或稍早日期。", "err");
    } else {
      const actualDate = (gainers[0] || losers[0]).net_value_date;
      setStatus(`查询完成：涨幅 ${gainers.length} 条，跌幅 ${losers.length} 条；实际净值日期 ${actualDate}`, "ok");
      $("excelBtn").disabled = false;
    }
  } catch (err) {
    setStatus("查询失败：" + err.message, "err");
  } finally {
    $("queryBtn").disabled = false;
  }
}

function toSheetRows(rows) {
  const header = ["排名", "基金代码", "基金名称", "净值日期", "单位净值", "累计净值", "日增长率(%)"];
  return [header].concat(rows.map((f, i) => [
    i + 1,
    f.fund_code,
    f.fund_name,
    f.net_value_date,
    Number(f.unit_net),
    Number(f.cumulative_net),
    Number(f.day_growth_rate.toFixed(2)),
  ]));
}

function setSheetWidth(sheet) {
  sheet["!cols"] = [
    {wch: 8}, {wch: 12}, {wch: 34}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}
  ];
}

function downloadExcel() {
  const wb = XLSX.utils.book_new();
  const s1 = XLSX.utils.aoa_to_sheet(toSheetRows(gainers));
  const s2 = XLSX.utils.aoa_to_sheet(toSheetRows(losers));
  setSheetWidth(s1);
  setSheetWidth(s2);
  XLSX.utils.book_append_sheet(wb, s1, "涨幅前100");
  XLSX.utils.book_append_sheet(wb, s2, "跌幅前100");
  XLSX.writeFile(wb, `天天基金涨跌幅前100排名查询-${currentDate.replaceAll("-", "")}.xlsx`);
}

(function initDate() {
  const d = new Date();
  $("queryDate").value = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
})();
</script>
</body>
</html>
"""


def parse_rank_data(text):
    match = re.search(r"datas\s*:\s*(\[.*?\])\s*,\s*allRecords", text, re.S)
    if not match:
        raise ValueError(text[:300])
    return json.loads(match.group(1))


def normalize_fund(item, rank):
    fields = item.split(",")
    if len(fields) < 7:
        return None
    try:
        day_growth_rate = float(fields[6])
    except ValueError:
        return None
    return {
        "rank": rank,
        "fund_code": fields[0],
        "fund_name": fields[1],
        "net_value_date": fields[3],
        "unit_net": fields[4],
        "cumulative_net": fields[5],
        "day_growth_rate": day_growth_rate,
    }


def fetch_ranking(query_date, sort_order):
    # 关键修复：日增长率排序字段是 rzdf，不是 srzdf。
    params = {
        "op": "ph",
        "dt": "kf",
        "ft": "all",
        "sc": "rzdf",
        "st": sort_order,
        "sd": "1900-01-01",
        "ed": query_date,
        "pi": "1",
        "pn": "100",
        "dx": "1",
    }
    resp = requests.get(EASTMONEY_API, params=params, headers=HEADERS, timeout=15)
    resp.encoding = "utf-8"
    resp.raise_for_status()
    funds = []
    for i, item in enumerate(parse_rank_data(resp.text), start=1):
        fund = normalize_fund(item, i)
        if fund:
            funds.append(fund)
    funds.sort(key=lambda x: x["day_growth_rate"], reverse=(sort_order == "desc"))
    for i, fund in enumerate(funds, start=1):
        fund["rank"] = i
    return funds[:100]


@app.route("/")
def index():
    return HTML_PAGE


@app.route("/api/ranking")
def api_ranking():
    query_date = request.args.get("date", "")
    sort_order = request.args.get("sort", "desc")
    if sort_order not in {"asc", "desc"}:
        return jsonify({"error": "sort 参数只能是 asc 或 desc"}), 400
    try:
        datetime.strptime(query_date, "%Y-%m-%d")
    except ValueError:
        return jsonify({"error": "date 参数格式应为 YYYY-MM-DD"}), 400
    try:
        return jsonify({"funds": fetch_ranking(query_date, sort_order)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 502


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8765, debug=False)
