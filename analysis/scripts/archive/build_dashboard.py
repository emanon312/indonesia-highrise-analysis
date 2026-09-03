# -*- coding: utf-8 -*-
"""
面板构建脚本（双数据源版：OSM vs 官方）

功能：
  1. 扫描 data/cities/*.json，每个文件代表一个城市
  2. 把所有城市数据内联进单个自包含 HTML 文件（内嵌为 JS 数据对象 var CITIES=[...]）
  3. 生成 output/dashboard.html，每次运行全量覆盖（加新城市后重跑即可更新）

数据结构（双数据源）：
  city.sources.osm      —— OpenStreetMap，按楼层（levels）
  city.sources.official —— 政府 LOD2 三维库，按高度（米）
  其中 official 可能缺失（其他城市可能只有 osm），面板需兼容：
    无 official 时只展示 OSM，不报错，也不显示对比区与官方视图切换。

自包含铁律：HTML 不依赖任何外部 CDN / 网络资源，所有 CSS / JS 内联，图表用纯 CSS 画。
"""

import json
import os
import glob

# 目录定位：脚本位于 analysis/scripts/，数据在 analysis/data/cities/，输出在 analysis/output/
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SCRIPT_DIR)
CITIES_DIR = os.path.join(ROOT_DIR, "data", "cities")
OUTPUT_DIR = os.path.join(ROOT_DIR, "output")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "dashboard.html")


def load_cities():
    """扫描 data/cities/*.json，按文件名排序后返回城市数据列表。"""
    cities = []
    for path in sorted(glob.glob(os.path.join(CITIES_DIR, "*.json"))):
        with open(path, "r", encoding="utf-8") as f:
            cities.append(json.load(f))
    if not cities:
        raise SystemExit("未找到任何城市数据文件（data/cities/*.json）")
    return cities


# 页面模板：用 __DATA_JSON__ 占位符注入内联数据，其余 CSS/JS 全部内联，无任何外链。
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>印尼高层建筑分析（真实数据 · OSM vs 官方对比）</title>
<style>
  /* =========================================================
     设计方向：「测绘账本 / Survey Ledger」
     —— 把 OSM 众包 与 政府官方 的认知落差当作品牌核心。
     数字采用等宽（账本）字体，prose 用无衬线；
     黄铜色=官方权威，冷蓝=众包，机构级深海军底。
     全部字体走系统安全栈，无任何外链。
     ========================================================= */
  :root{
    --ink:#0A1322;          /* 测绘深海军底 */
    --ink-2:#0E1A2E;        /* 面板底 */
    --surface:#132439;      /* 抬升表面 */
    --surface-2:#172C45;    /* 次级表面 / 轨道槽 */
    --rule:#203552;         /* 实体分割线 */
    --paper:#E8EEF6;        /* 主文字 */
    --muted:#8398B5;        /* 次要文字 */
    --faint:#566F8F;        /* 更弱文字 */
    --brass:#C8A04E;        /* 黄铜 = 官方权威强调 */
    --brass-soft:#E0BE74;   /* 黄铜亮调 */
    --osm:#5BA3F0;          /* 众包蓝 */
    --osm-soft:#8FC3FB;
    --official:#58C9A0;     /* 官方机构青绿 */
    --official-soft:#86E0BE;
    --green:#58C9A0;
    /* 等宽账本字体栈（系统内置，无外链） */
    --mono:"SF Mono","Cascadia Mono","JetBrains Mono","Consolas","Menlo",ui-monospace,monospace;
    --sans:"Inter","Segoe UI","PingFang SC","Microsoft YaHei",system-ui,sans-serif;
  }
  *{box-sizing:border-box;}
  html{-webkit-text-size-adjust:100%;}
  body{
    margin:0;color:var(--paper);
    background:var(--ink);
    /* 极淡的测绘网格底纹，纯 CSS 渐变，强化"账本/坐标"质感 */
    background-image:
      linear-gradient(rgba(91,163,240,.035) 1px,transparent 1px),
      linear-gradient(90deg,rgba(91,163,240,.035) 1px,transparent 1px),
      radial-gradient(ellipse 80% 60% at 50% -10%,rgba(200,160,78,.10),transparent 70%);
    background-size:44px 44px,44px 44px,100% 100%;
    font-family:var(--sans);
    line-height:1.6;font-size:15px;
    -webkit-font-smoothing:antialiased;
  }
  /* 全部数字走等宽账本栈：投资数据靠对齐说话 */
  .mono,.kpi .num,.cmp-bar .n,.cmp-card .ratio b,.cols .cnum,
  td.lv,td.rank,.bars .val,.uc-kpi .num{
    font-family:var(--mono);font-feature-settings:"tnum" 1;
    font-variant-numeric:tabular-nums;letter-spacing:-.01em;
  }
  .wrap{max-width:1100px;margin:0 auto;padding:0 18px 72px;}

  /* 通用结构性 eyebrow（小标签，宽字距大写感） */
  .eyebrow{
    font-family:var(--mono);font-size:11px;letter-spacing:.22em;
    text-transform:uppercase;color:var(--brass);font-weight:600;
  }

  /* ===== 报头 ===== */
  header.top{
    padding:34px 0 26px;margin-bottom:6px;
    border-bottom:1px solid var(--rule);
    position:relative;
  }
  header.top .kicker{
    display:flex;align-items:center;gap:10px;margin-bottom:14px;
    font-family:var(--mono);font-size:11px;letter-spacing:.2em;
    text-transform:uppercase;color:var(--faint);
  }
  header.top .kicker .seal{
    color:var(--brass);letter-spacing:.16em;
  }
  header.top .kicker .rule{flex:1;height:1px;background:var(--rule);}
  header.top h1{
    font-size:30px;line-height:1.18;margin:0 0 12px;
    letter-spacing:-.015em;font-weight:700;max-width:22ch;
  }
  header.top h1 .em{color:var(--brass-soft);}
  header.top .sub{
    color:var(--muted);font-size:14.5px;max-width:62ch;
  }
  header.top .sub .div{color:var(--rule);margin:0 9px;}

  /* 城市切换器：分段控件风 */
  .city-switch{
    display:flex;flex-wrap:wrap;gap:6px;
    margin:22px 0 4px;
  }
  .city-tab{
    background:transparent;border:1px solid var(--rule);color:var(--muted);
    padding:7px 18px;border-radius:8px;cursor:pointer;font-size:14px;
    font-weight:500;transition:all .18s ease;
  }
  .city-tab:hover{color:var(--paper);border-color:var(--osm);}
  .city-tab.active{
    background:linear-gradient(180deg,#1B355A,#13243A);
    border-color:var(--osm);color:#fff;font-weight:600;
    box-shadow:0 1px 0 rgba(91,163,240,.4) inset,0 6px 18px -10px rgba(91,163,240,.7);
  }

  section{margin-top:46px;}
  h2.title{
    font-size:17px;margin:0 0 16px;font-weight:600;
    letter-spacing:-.01em;display:flex;align-items:baseline;gap:10px;
    color:var(--paper);
  }
  h2.title::before{
    content:"";width:3px;height:15px;flex:none;border-radius:2px;
    background:var(--brass);transform:translateY(2px);
  }
  h2.title small{color:var(--faint);font-size:12.5px;font-weight:normal;font-family:var(--mono);letter-spacing:.02em;}

  /* ===== 核心对比区：OSM vs 官方 ===== */
  .cmp-legend{
    display:flex;gap:22px;align-items:center;margin:-2px 0 18px;
    font-size:12.5px;color:var(--muted);font-family:var(--mono);letter-spacing:.02em;
  }
  .cmp-legend .dot{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:7px;vertical-align:middle;}
  .dot.osm{background:var(--osm);}
  .dot.official{background:var(--official);}

  .cmp-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
  .cmp-card{
    background:linear-gradient(180deg,var(--surface),var(--ink-2));
    border:1px solid var(--rule);border-radius:14px;
    padding:20px 18px 18px;position:relative;overflow:hidden;
  }
  .cmp-card.highlight{
    border-color:rgba(200,160,78,.55);
    box-shadow:0 0 0 1px rgba(200,160,78,.25) inset,0 18px 40px -28px rgba(200,160,78,.7);
    background:linear-gradient(180deg,#1C2A3E,#121F33);
  }
  /* 标准档卡片顶部黄铜光条 */
  .cmp-card.highlight::before{
    content:"";position:absolute;top:0;left:0;right:0;height:2px;
    background:linear-gradient(90deg,transparent,var(--brass),transparent);
  }
  .cmp-card .ttl{font-size:14px;font-weight:600;margin-bottom:1px;letter-spacing:.01em;}
  .cmp-card .thr{
    color:var(--faint);font-size:11.5px;margin-bottom:16px;
    font-family:var(--mono);letter-spacing:.03em;
  }
  .cmp-card.highlight .ttl{color:var(--brass-soft);}
  .cmp-card .badge{
    position:absolute;top:16px;right:16px;font-size:10px;color:var(--brass-soft);
    border:1px solid rgba(200,160,78,.45);background:rgba(200,160,78,.08);
    padding:2px 9px;border-radius:6px;font-family:var(--mono);letter-spacing:.08em;
  }
  /* 单源横向对比条 */
  .cmp-bar{margin:12px 0;}
  .cmp-bar .head{display:flex;justify-content:space-between;align-items:baseline;font-size:12px;margin-bottom:5px;}
  .cmp-bar .head .src{
    color:var(--muted);font-family:var(--mono);font-size:11px;
    letter-spacing:.06em;text-transform:uppercase;
  }
  .cmp-bar .head .n{font-weight:600;font-size:18px;}
  .cmp-bar .track{
    background:var(--surface-2);border-radius:3px;height:9px;overflow:hidden;
    box-shadow:0 1px 1px rgba(0,0,0,.3) inset;
  }
  .cmp-bar .fill{height:100%;border-radius:3px;min-width:3px;transition:width .6s cubic-bezier(.2,.7,.2,1);}
  .cmp-bar.osm .n{color:var(--osm-soft);}
  .cmp-bar.osm .fill{background:linear-gradient(90deg,#3F7FCB,var(--osm));}
  .cmp-bar.official .n{color:var(--official-soft);}
  .cmp-bar.official .fill{background:linear-gradient(90deg,#2E9B79,var(--official));}
  .cmp-card .ratio{
    margin-top:14px;padding-top:12px;border-top:1px dashed var(--rule);
    color:var(--muted);font-size:11.5px;letter-spacing:.02em;
  }
  .cmp-card .ratio b{color:var(--brass-soft);font-size:14px;}

  /* 仅 OSM（无官方）时的 KPI 卡片 */
  .kpi-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
  .kpi{
    background:linear-gradient(180deg,var(--surface),var(--ink-2));
    border:1px solid var(--rule);border-radius:14px;
    padding:22px 18px;
  }
  .kpi .label{color:var(--muted);font-size:13px;}
  .kpi .num{font-size:40px;font-weight:600;margin:8px 0 2px;line-height:1;}
  .kpi .unit{color:var(--faint);font-size:12px;font-family:var(--mono);letter-spacing:.03em;}
  .kpi.highlight{border-color:rgba(200,160,78,.5);box-shadow:0 0 0 1px rgba(200,160,78,.2) inset;}
  .kpi.highlight .num{color:var(--brass-soft);}

  /* ===== 数据源视图切换 ===== */
  .view-switch{
    display:inline-flex;gap:4px;margin-bottom:20px;padding:4px;
    background:var(--ink-2);border:1px solid var(--rule);border-radius:11px;
  }
  .view-tab{
    text-align:center;background:transparent;border:1px solid transparent;
    color:var(--muted);padding:8px 18px;border-radius:8px;cursor:pointer;font-size:13.5px;
    font-weight:500;transition:all .16s ease;white-space:nowrap;
  }
  .view-tab:hover{color:var(--paper);}
  .view-tab.active{color:#fff;font-weight:600;}
  .view-tab.osm.active{background:rgba(91,163,240,.16);border-color:rgba(91,163,240,.5);color:var(--osm-soft);}
  .view-tab.official.active{background:rgba(88,201,160,.16);border-color:rgba(88,201,160,.5);color:var(--official-soft);}

  .view-meta{color:var(--muted);font-size:13px;margin:-6px 0 16px;}

  /* 横向条形图（类型/用途分布） */
  .bars{margin-top:2px;}
  .bars .row{display:flex;align-items:center;margin:9px 0;gap:12px;}
  .bars .name{width:104px;flex:none;color:var(--muted);font-size:13px;text-align:right;}
  .bars .track{flex:1;background:var(--surface-2);border-radius:4px;height:10px;position:relative;overflow:hidden;box-shadow:0 1px 1px rgba(0,0,0,.3) inset;}
  .bars .fill{height:100%;border-radius:4px;min-width:2px;transition:width .5s ease;}
  .bars.osm .fill{background:linear-gradient(90deg,#3F7FCB,var(--osm));}
  .bars.official .fill{background:linear-gradient(90deg,#2E9B79,var(--official));}
  .bars .val{width:56px;flex:none;font-size:13.5px;text-align:right;color:var(--paper);}

  /* 分布柱状图 */
  .cols{display:flex;align-items:flex-end;gap:10px;height:220px;
    background:linear-gradient(180deg,var(--surface),var(--ink-2));
    border:1px solid var(--rule);border-radius:14px;padding:20px 16px 10px;}
  .cols .col{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%;}
  .cols .barv{width:64%;border-radius:3px 3px 0 0;min-height:2px;transition:height .5s ease;}
  .cols.osm .barv{background:linear-gradient(180deg,var(--osm),#2F5F97);}
  .cols.official .barv{background:linear-gradient(180deg,var(--official),#1F7558);}
  .cols .cnum{font-size:12px;color:var(--paper);margin-bottom:6px;}
  .cols .clab{font-size:11px;color:var(--faint);margin-top:8px;white-space:nowrap;font-family:var(--mono);letter-spacing:.01em;}

  /* 表格 */
  table{width:100%;border-collapse:collapse;background:var(--ink-2);border:1px solid var(--rule);border-radius:14px;overflow:hidden;}
  th,td{padding:11px 14px;text-align:left;font-size:13.5px;border-bottom:1px solid var(--rule);}
  th{
    background:var(--surface);color:var(--faint);font-weight:500;
    font-family:var(--mono);font-size:11px;letter-spacing:.08em;text-transform:uppercase;
  }
  tbody tr{transition:background .12s;}
  tbody tr:hover{background:rgba(91,163,240,.05);}
  td.rank,th.rank{width:46px;text-align:center;color:var(--brass);}
  td.lv,th.lv{text-align:right;}
  td.lv{color:var(--paper);font-weight:600;}
  tr:last-child td{border-bottom:none;}

  /* 类型标签 */
  .tag{
    font-size:11.5px;padding:2px 9px;border-radius:6px;
    background:var(--surface-2);color:var(--muted);
    border:1px solid var(--rule);
  }

  /* 在建栏 */
  .uc-warn{
    background:rgba(200,160,78,.08);border:1px solid rgba(200,160,78,.35);color:var(--brass-soft);
    padding:11px 16px;border-radius:10px;font-size:13.5px;margin-bottom:16px;
  }
  .uc-warn b{color:#F0D08A;}
  .uc-kpi{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px;}
  .uc-kpi .kpi .num{font-size:32px;color:var(--brass-soft);}

  /* 数据透明栏 */
  .src-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
  .source{
    background:linear-gradient(180deg,var(--surface),var(--ink-2));
    border:1px solid var(--rule);border-radius:14px;padding:20px;
    position:relative;
  }
  .source::before{
    content:"";position:absolute;top:0;left:20px;right:20px;height:2px;border-radius:2px;
  }
  .source.osm::before{background:var(--osm);}
  .source.official::before{background:var(--official);}
  .source h3{margin:2px 0 12px;font-size:14px;font-weight:600;display:flex;align-items:center;gap:8px;}
  .source h3::before{
    content:"";width:8px;height:8px;border-radius:2px;
  }
  .source.osm h3{color:var(--osm-soft);}
  .source.osm h3::before{background:var(--osm);}
  .source.official h3{color:var(--official-soft);}
  .source.official h3::before{background:var(--official);}
  .source .line{margin:6px 0;font-size:13px;color:var(--muted);line-height:1.65;}
  .source .line b{color:var(--paper);font-weight:600;}
  .claim{
    margin-top:16px;padding:15px;border-radius:12px;
    background:rgba(88,201,160,.10);border:1px solid rgba(88,201,160,.45);
    color:var(--official-soft);font-weight:600;text-align:center;font-size:14.5px;
    letter-spacing:.01em;
  }
  .claim::before{content:"✓ ";color:var(--official);font-weight:700;}
  .compare-note{
    margin-top:12px;padding:14px 16px;border-radius:12px;
    background:var(--ink-2);border:1px solid var(--rule);
    color:var(--muted);font-size:13px;line-height:1.75;
  }

  .hidden{display:none !important;}
  footer{
    text-align:center;color:var(--faint);font-size:11.5px;margin-top:54px;
    padding-top:22px;border-top:1px solid var(--rule);
    font-family:var(--mono);letter-spacing:.04em;
  }

  @media (prefers-reduced-motion:reduce){
    *{transition:none !important;}
  }

  /* 响应式：手机 */
  @media (max-width:680px){
    .wrap{padding:0 14px 56px;}
    header.top h1{font-size:23px;}
    .cmp-grid{grid-template-columns:1fr;}
    .kpi-grid,.uc-kpi{grid-template-columns:1fr;}
    .src-grid{grid-template-columns:1fr;}
    .kpi .num{font-size:34px;}
    .cols{height:184px;gap:5px;padding:16px 12px 8px;}
    .cols .clab{font-size:9.5px;}
    .cols .barv{width:72%;}
    .bars .name{width:78px;}
    .view-switch{display:flex;width:100%;}
    .view-tab{flex:1;}
  }
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <div class="kicker">
      <span class="seal">◆ 测绘双源核对</span>
      <span class="rule"></span>
      <span>真实数据 · 离线内嵌</span>
    </div>
    <h1>印尼高层建筑投资分析<br><span class="em">众包视角 与 官方测绘 的覆盖落差</span></h1>
    <div class="sub">同一片天际线，两种"看见"方式：OpenStreetMap 众包标注按楼层计，政府 LOD2 三维库按高度（米）计。<span class="div">|</span>本面板逐档并列对照，不做任何推断补全。</div>
  </header>

  <!-- 城市切换器（标签页，纯前端切换） -->
  <div class="city-switch" id="citySwitch"></div>

  <!-- ===== 核心对比区：OSM vs 官方 ===== -->
  <section id="compareSection">
    <h2 class="title">核心对比：OSM 众包 vs 政府官方<small id="compareCoverage"></small></h2>
    <div class="cmp-legend">
      <span><i class="dot osm"></i>OSM 众包（按楼层）</span>
      <span><i class="dot official"></i>政府官方（按高度米）</span>
    </div>
    <div class="cmp-grid" id="compareGrid"></div>
  </section>

  <!-- 仅 OSM（无官方数据）时显示的简单 KPI 区 -->
  <section id="osmOnlySection" class="hidden">
    <h2 class="title">已建成高层建筑数量（OSM）</h2>
    <div class="kpi-grid" id="osmOnlyGrid"></div>
  </section>

  <!-- ===== 数据源详情视图（OSM / 官方 切换） ===== -->
  <section>
    <h2 class="title">数据源详情</h2>
    <div class="view-switch" id="viewSwitch"></div>

    <!-- 分布：类型/用途 -->
    <h2 class="title" id="byTypeTitle">类型分布</h2>
    <div class="bars" id="typeBars"></div>

    <!-- 分布：楼层/高度 -->
    <h2 class="title" id="distTitle">分布</h2>
    <div class="cols" id="distCols"></div>

    <!-- 仅官方视图：各区分布 -->
    <div id="districtBlock" class="hidden">
      <h2 class="title">各区分布（≥40 米数量）</h2>
      <div class="cols official" id="districtCols"></div>
    </div>

    <!-- Top 建筑表格 -->
    <h2 class="title" id="topTitle">最高建筑 Top25</h2>
    <table>
      <thead><tr id="topHead"></tr></thead>
      <tbody id="topBody"></tbody>
    </table>

    <!-- 仅 OSM 视图：在建栏 -->
    <div id="ucBlock" class="hidden">
      <h2 class="title">在建建筑（未建成，仅供参考）</h2>
      <div class="uc-warn">⚠ 以下为<b>在建（未建成）</b>项目，仅供参考，不计入已建成统计。</div>
      <div class="uc-kpi">
        <div class="kpi"><div class="label">在建 ≥8 层</div><div class="num" id="uc8">0</div><div class="unit">栋</div></div>
        <div class="kpi"><div class="label">在建 ≥10 层</div><div class="num" id="uc10">0</div><div class="unit">栋</div></div>
        <div class="kpi"><div class="label">在建 ≥12 层</div><div class="num" id="uc12">0</div><div class="unit">栋</div></div>
      </div>
      <table>
        <thead><tr><th class="rank">#</th><th>名称</th><th class="lv">楼层</th><th>类型</th></tr></thead>
        <tbody id="ucBody"></tbody>
      </table>
    </div>
  </section>

  <!-- ===== 数据透明栏 ===== -->
  <section>
    <h2 class="title">数据来源与透明声明</h2>
    <div class="src-grid" id="srcGrid"></div>
    <div class="claim">两源均为真实数据，未做任何推断补全</div>
    <div class="compare-note" id="compareNote"></div>
  </section>

  <footer>印尼高层建筑分析面板 · 数据离线内嵌 · 无外部网络依赖</footer>
</div>

<script>
// ===== 内联城市数据（构建时由 Python 注入） =====
var CITIES = __DATA_JSON__;

// 当前选中城市与当前数据源视图
var curCity = null;
var curView = "osm"; // "osm" | "official"

function el(id){return document.getElementById(id);}
function num(n){return (n == null ? 0 : n).toLocaleString();}
function esc(s){
  return String(s == null ? "" : s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}
// 阈值双口径标签，如 "≥10层 / ≥40米"
function thrLabel(t){return "≥" + t.floors + "层 / ≥" + t.meters + "米";}

// ===== 核心对比区：三档并排，OSM vs 官方 =====
function renderCompare(city){
  var osm = city.sources.osm;
  var off = city.sources.official;
  var thresholds = city.thresholds || [];

  // 无官方数据：隐藏对比区，改用纯 OSM 的 KPI 卡片
  if (!off){
    el("compareSection").classList.add("hidden");
    el("osmOnlySection").classList.remove("hidden");
    var gh = "";
    thresholds.forEach(function(t){
      var hl = (t.key === "std") ? " highlight" : "";
      gh += '<div class="kpi' + hl + '">'
        + '<div class="label">' + esc(t.label) + ' ' + thrLabel(t) + '</div>'
        + '<div class="num">' + num(osm.counts[t.key]) + '</div>'
        + '<div class="unit">栋（OSM 已标注）</div>'
        + '</div>';
    });
    el("osmOnlyGrid").innerHTML = gh;
    return;
  }

  el("compareSection").classList.remove("hidden");
  el("osmOnlySection").classList.add("hidden");
  el("compareCoverage").textContent = "官方覆盖远超众包，下图等比对照";

  // 全局最大值（取所有档两源最大），保证三张卡片条形图可横向比较
  var gmax = 1;
  thresholds.forEach(function(t){
    gmax = Math.max(gmax, osm.counts[t.key] || 0, off.counts[t.key] || 0);
  });

  var html = "";
  thresholds.forEach(function(t){
    var ov = osm.counts[t.key] || 0;
    var fv = off.counts[t.key] || 0;
    var hl = (t.key === "std") ? " highlight" : "";
    var ratio = ov > 0 ? (fv / ov).toFixed(1) : "—";
    html += '<div class="cmp-card' + hl + '">'
      + '<div class="ttl">' + esc(t.label) + '</div>'
      + '<div class="thr">' + thrLabel(t) + '</div>'
      + (t.key === "std" ? '<span class="badge">标准档（视觉高亮）</span>' : '')
      + '<div class="cmp-bar osm">'
      +   '<div class="head"><span class="src">OSM 众包</span><span class="n">' + num(ov) + '</span></div>'
      +   '<div class="track"><div class="fill" style="width:' + (ov / gmax * 100).toFixed(1) + '%"></div></div>'
      + '</div>'
      + '<div class="cmp-bar official">'
      +   '<div class="head"><span class="src">政府官方</span><span class="n">' + num(fv) + '</span></div>'
      +   '<div class="track"><div class="fill" style="width:' + (fv / gmax * 100).toFixed(1) + '%"></div></div>'
      + '</div>'
      + '<div class="ratio">官方约为 OSM 的 <b>' + ratio + '×</b></div>'
      + '</div>';
  });
  el("compareGrid").innerHTML = html;
}

// ===== 数据源详情：根据 curView 渲染 OSM 或官方 =====
function renderDetail(city){
  var off = city.sources.official;
  // 无官方时强制 OSM 视图，并隐藏视图切换标签
  if (!off) curView = "osm";
  var src = city.sources[curView];

  // 视图切换标签（无官方则只显示 OSM 单标签）
  var vs = '<div class="view-tab osm' + (curView === "osm" ? " active" : "") + '" '
         + 'onclick="switchView(\'osm\')">OSM 视图（按楼层）</div>';
  if (off){
    vs += '<div class="view-tab official' + (curView === "official" ? " active" : "") + '" '
        + 'onclick="switchView(\'official\')">官方视图（按高度米）</div>';
  }
  el("viewSwitch").innerHTML = vs;

  // 类型/用途分布
  el("byTypeTitle").textContent = src.by_type_label || "类型分布";
  var byType = src.by_type || {};
  var keys = Object.keys(byType);
  var maxType = 1;
  keys.forEach(function(k){ maxType = Math.max(maxType, byType[k]); });
  var barsCls = (curView === "official") ? "bars official" : "bars osm";
  el("typeBars").className = barsCls;
  var bh = "";
  // 按数量从大到小排序展示
  keys.sort(function(a,b){ return byType[b] - byType[a]; }).forEach(function(k){
    var v = byType[k];
    bh += '<div class="row">'
      + '<div class="name">' + esc(k) + '</div>'
      + '<div class="track"><div class="fill" style="width:' + (v / maxType * 100).toFixed(1) + '%"></div></div>'
      + '<div class="val">' + num(v) + '</div>'
      + '</div>';
  });
  el("typeBars").innerHTML = bh;

  // 楼层/高度分布柱状图
  el("distTitle").textContent = src.distribution_label || "分布";
  var dist = src.distribution || [];
  var maxD = 1;
  dist.forEach(function(d){ maxD = Math.max(maxD, d.count); });
  el("distCols").className = (curView === "official") ? "cols official" : "cols osm";
  var dh = "";
  dist.forEach(function(d){
    dh += '<div class="col">'
      + '<div class="cnum">' + num(d.count) + '</div>'
      + '<div class="barv" style="height:' + (d.count / maxD * 100).toFixed(1) + '%"></div>'
      + '<div class="clab">' + esc(d.range) + '</div>'
      + '</div>';
  });
  el("distCols").innerHTML = dh;

  // 各区分布（仅官方视图，且存在 by_district 时）
  if (curView === "official" && src.by_district){
    el("districtBlock").classList.remove("hidden");
    var bd = src.by_district;
    var dk = Object.keys(bd);
    var maxBd = 1;
    dk.forEach(function(k){ maxBd = Math.max(maxBd, bd[k].ge40 || 0); });
    var bdh = "";
    // 按 ge40 从大到小
    dk.sort(function(a,b){ return (bd[b].ge40||0) - (bd[a].ge40||0); }).forEach(function(k){
      var v = bd[k].ge40 || 0;
      bdh += '<div class="col">'
        + '<div class="cnum">' + num(v) + '</div>'
        + '<div class="barv" style="height:' + (v / maxBd * 100).toFixed(1) + '%"></div>'
        + '<div class="clab">' + esc(k) + '</div>'
        + '</div>';
    });
    el("districtCols").innerHTML = bdh;
  } else {
    el("districtBlock").classList.add("hidden");
  }

  // Top 建筑表格：用 top_metric / top_metric_label 决定显示楼层还是高度
  var metric = src.top_metric;          // "levels" 或 "height_m"
  var metricLabel = src.top_metric_label; // "楼层" 或 "高度(米)"
  el("topTitle").textContent = "最高建筑 Top25（" + (curView === "official" ? "官方 · 按高度" : "OSM · 按楼层") + "）";
  // 表头：官方视图额外显示「区」，类型列名按数据源区分
  if (curView === "official"){
    el("topHead").innerHTML = '<th class="rank">#</th><th>名称</th><th class="lv">' + esc(metricLabel)
      + '</th><th>区</th><th>用途</th>';
  } else {
    el("topHead").innerHTML = '<th class="rank">#</th><th>名称</th><th class="lv">' + esc(metricLabel)
      + '</th><th>类型</th>';
  }
  var tb = src.top_buildings || [];
  var th = "";
  tb.forEach(function(b){
    var v = b[metric];
    if (curView === "official"){
      th += '<tr>'
        + '<td class="rank">' + b.rank + '</td>'
        + '<td>' + esc(b.name || "（无名称）") + '</td>'
        + '<td class="lv">' + v + '</td>'
        + '<td><span class="tag">' + esc(b.district || "—") + '</span></td>'
        + '<td><span class="tag">' + esc(b.fungsi || "未分类") + '</span></td>'
        + '</tr>';
    } else {
      th += '<tr>'
        + '<td class="rank">' + b.rank + '</td>'
        + '<td>' + esc(b.name || "（无名称）") + '</td>'
        + '<td class="lv">' + v + '</td>'
        + '<td><span class="tag">' + esc(b.type || "未分类") + '</span></td>'
        + '</tr>';
    }
  });
  el("topBody").innerHTML = th;

  // 在建栏（仅 OSM 视图，且存在 under_construction 时）
  var uc = (curView === "osm") ? city.sources.osm.under_construction : null;
  if (uc){
    el("ucBlock").classList.remove("hidden");
    el("uc8").textContent  = num(uc.ge8);
    el("uc10").textContent = num(uc.ge10);
    el("uc12").textContent = num(uc.ge12);
    var ucb = uc.top_buildings || [];
    var uh = "";
    ucb.forEach(function(b, i){
      uh += '<tr>'
        + '<td class="rank">' + (i + 1) + '</td>'
        + '<td>' + esc(b.name || "（无名称）") + '</td>'
        + '<td class="lv">' + b.levels + '</td>'
        + '<td><span class="tag">' + esc(b.type || "未分类") + '</span></td>'
        + '</tr>';
    });
    el("ucBody").innerHTML = uh;
  } else {
    el("ucBlock").classList.add("hidden");
  }
}

// ===== 数据透明栏：分别标注两源 =====
function renderSources(city){
  var osm = city.sources.osm;
  var off = city.sources.official;
  var html = "";
  html += '<div class="source osm">'
    + '<h3>OSM 众包数据</h3>'
    + '<div class="line">名称：<b>' + esc(osm.name) + '</b></div>'
    + '<div class="line">抓取日期：<b>' + esc(osm.fetch_date) + '</b></div>'
    + '<div class="line">说明：<b>' + esc(osm.disclaimer) + '</b></div>'
    + '<div class="line">覆盖：<b>' + esc(osm.coverage_note) + '</b></div>'
    + '</div>';
  if (off){
    html += '<div class="source official">'
      + '<h3>政府官方数据</h3>'
      + '<div class="line">名称：<b>' + esc(off.name) + '</b></div>'
      + '<div class="line">抓取日期：<b>' + esc(off.fetch_date) + '</b></div>'
      + '<div class="line">说明：<b>' + esc(off.disclaimer) + '</b></div>'
      + '<div class="line">覆盖：<b>' + esc(off.coverage_note)
      + (off.total_buildings ? '（全市约 ' + num(off.total_buildings) + ' 栋）' : '') + '</b></div>'
      + '</div>';
  }
  el("srcGrid").innerHTML = html;
  // 若无官方，对比说明仍展示城市自带 compare_note
  el("compareNote").textContent = city.compare_note
    || "层与米为近似对齐（约 4 米/层）。";
}

// 切换数据源视图
function switchView(v){
  curView = v;
  renderDetail(curCity);
}

// 渲染整座城市
function render(city){
  curCity = city;
  // 默认进 OSM 视图（每次切城市重置）
  curView = "osm";
  renderCompare(city);
  renderDetail(city);
  renderSources(city);
}

// 构建城市切换标签
function buildSwitch(){
  var box = el("citySwitch");
  box.innerHTML = "";
  CITIES.forEach(function(city, idx){
    var tab = document.createElement("div");
    tab.className = "city-tab" + (idx === 0 ? " active" : "");
    tab.textContent = city.city_name_cn + " " + city.city_name_en;
    tab.onclick = function(){
      var all = box.getElementsByClassName("city-tab");
      for (var i = 0; i < all.length; i++){ all[i].className = "city-tab"; }
      tab.className = "city-tab active";
      render(city);
    };
    box.appendChild(tab);
  });
}

// 初始化：默认展示第一个城市
buildSwitch();
render(CITIES[0]);
</script>
</body>
</html>
"""


def build():
    """读取城市数据，注入模板，覆盖写出 dashboard.html。"""
    cities = load_cities()
    # ensure_ascii=False 保留中文；注入位置在 <script> 内，作为合法 JS 字面量
    data_json = json.dumps(cities, ensure_ascii=False)
    html = HTML_TEMPLATE.replace("__DATA_JSON__", data_json)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print("已生成：" + OUTPUT_FILE)
    print("包含城市数量：" + str(len(cities)))
    for c in cities:
        srcs = c.get("sources", {})
        flags = []
        if "osm" in srcs:
            flags.append("OSM")
        if "official" in srcs:
            flags.append("官方")
        print("  - " + c.get("city_name_cn", "?") + " (" + c.get("city_id", "?")
              + ") 数据源：" + ("+".join(flags) if flags else "无"))


if __name__ == "__main__":
    build()
