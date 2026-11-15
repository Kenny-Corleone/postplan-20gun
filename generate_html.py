# -*- coding: utf-8 -*-
import json, os

exec(open(r"C:\Users\ANCL\Desktop\20 gun\_data.py", encoding="utf-8").read())

data1_json = json.dumps(posts1, ensure_ascii=False)
data2_json = json.dumps(char_posts, ensure_ascii=False)

import base64 as _b64
logo_b64 = _b64.b64encode(open(r"C:\Users\ANCL\Desktop\20 gun\20 gun cover.jpg","rb").read()).decode()

html_doc = r'''<!DOCTYPE html>
<html lang="az">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Postplan — 20 günlük dil öyrənmə (TikTok)</title>
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter','Segoe UI',Arial,sans-serif;background:#0b1120;color:#d1d5db;line-height:1.5;padding:24px;font-size:13px}
  .container{max-width:1700px;margin:0 auto}
  .floating-logo{position:fixed;top:16px;left:16px;z-index:1000;background:#111827;border:1px solid #1f2937;border-radius:10px;padding:6px;box-shadow:0 4px 16px rgba(0,0,0,.4)}
  .floating-logo img{height:90px;width:auto;border-radius:6px;display:block}
  .logo-bar{display:flex;align-items:center;justify-content:flex-start;gap:24px;padding:24px;background:linear-gradient(180deg,#111827 0%,#0b1120 100%);border:1px solid #1f2937;border-radius:8px;margin-bottom:20px;}
  .logo-bar img{height:150px;width:auto;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.3)}
  .logo-bar .title{font-size:26px;font-weight:700;color:#f3f4f6;letter-spacing:.3px}
  .logo-bar .sub{font-size:12px;color:#9ca3af;margin-top:2px}
  header{text-align:center;padding:24px 24px;background:linear-gradient(180deg,#111827 0%,#0b1120 100%);border:1px solid #1f2937;border-radius:8px;margin-bottom:20px;}
  header h1{font-size:20px;font-weight:700;margin-bottom:4px;color:#f3f4f6;letter-spacing:.3px}
  header p{color:#9ca3af;font-size:13px;max-width:760px;margin:0 auto}
  .meta{display:flex;flex-wrap:wrap;gap:8px;justify-content:center;margin-top:12px}
  .meta span{background:#111827;border:1px solid #1f2937;padding:5px 12px;border-radius:4px;font-size:12px;color:#9ca3af}
  .tabs{display:flex;gap:4px;margin-bottom:20px;border-bottom:1px solid #1f2937;flex-wrap:wrap}
  .tab{padding:10px 20px;cursor:pointer;color:#9ca3af;border:1px solid transparent;border-bottom:none;border-radius:6px 6px 0 0;font-size:13px;font-weight:500;background:transparent}
  .tab.active{color:#b8963e;border-color:#1f2937;background:#111827;border-bottom:1px solid #111827;margin-bottom:-1px}
  .tab:hover{color:#d1d5db}
  .tabpane{display:none}
  .tabpane.active{display:block}
  .plan-section{background:#111827;border:1px solid #1f2937;border-radius:8px;padding:20px;margin-bottom:24px;}
  .plan-section h2{font-size:17px;font-weight:600;margin-bottom:14px;color:#e5e7eb;border-bottom:1px solid #1f2937;padding-bottom:10px;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
  .toolbar{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .btn{padding:7px 12px;border-radius:6px;border:1px solid #d1d5db;background:#ffffff;color:#374151;font-size:12px;cursor:pointer;font-weight:500}
  .btn:hover{border-color:#92400e;color:#92400e}
  .btn-primary{border-color:#92400e;color:#92400e;background:#fffbeb}
  .btn-danger{border-color:#dc2626;color:#dc2626}
  .btn-danger:hover{background:#dc2626;color:#fff}
  .btn-success{border-color:#059669;color:#059669;background:#ecfdf5}
  .btn-success:hover{background:#059669;color:#fff}
  .table-wrap{overflow-x:auto}
  table{width:100%;border-collapse:collapse;font-size:12px}
  th,td{padding:6px 7px;border-bottom:1px solid #1f2937;vertical-align:top}
  th{background:#0b1120;color:#9ca3af;font-weight:600;position:sticky;top:0;z-index:2;white-space:nowrap;font-size:10px;text-transform:uppercase;letter-spacing:.5px}
  tbody tr{transition:background .15s}
  tbody tr:hover{background:#1a2332}
  td[contenteditable="true"]{outline:none;min-height:14px}
  td[contenteditable="true"]:focus{background:#1f2937;border-radius:2px;box-shadow:0 0 0 1px #b8963e}
  input.cell-input{border:1px solid #374151;border-radius:4px;padding:3px 5px;font-size:11px;background:#0b1120;color:#d1d5db;width:100%}
  input.cell-input:focus{outline:none;border-color:#b8963e}
  .hook{font-weight:500;color:#e5e7eb;min-width:160px}
  .cta{color:#9ca3af;min-width:120px;font-size:11px}
  .tags{font-size:10px;color:#6b7280;min-width:140px}
  .sc-meta{font-size:10px;color:#9ca3af;min-width:110px}
  .sc-meta div{margin:1px 0}
  .sc-meta b{color:#b8963e;font-weight:500}
  .scenario{min-width:240px}
  details{cursor:pointer}
  details summary{color:#b8963e;font-weight:500;list-style:none;font-size:11px}
  details summary::-webkit-details-marker{display:none}
  details summary::before{content:'▸ ';font-size:9px;color:#6b7280}
  details[open] summary::before{content:'▾ '}
  .sc-block{margin:4px 0;padding:4px 7px;background:#0b1120;border-left:2px solid #b8963e;border-radius:2px;font-size:11px;color:#374151}
  .sc-block b{color:#b8963e;font-weight:600}
  .badge{display:inline-block;padding:2px 7px;border-radius:3px;font-size:10px;font-weight:500;white-space:nowrap;border:1px solid}
  .badge-type{border-color:#374151;color:#9ca3af;background:transparent}
  .badge-topic{border-color:#374151;color:#d1d5db;background:transparent}
  .badge-goal{border-color:#374151;color:#9ca3af;background:transparent}
  .status-select{background:#0b1120;border:1px solid #374151;color:#d1d5db;padding:3px 6px;border-radius:4px;font-size:11px;cursor:pointer;width:100%}
  .status-select:focus{outline:none;border-color:#b8963e}
  .row-actions{display:flex;gap:4px;justify-content:center}
  .icon-btn{background:transparent;border:1px solid #374151;color:#9ca3af;padding:2px 6px;border-radius:4px;cursor:pointer;font-size:11px}
  .icon-btn:hover{border-color:#b8963e;color:#b8963e}
  .icon-btn.del:hover{border-color:#dc2626;color:#dc2626}
  .params{background:#111827;border:1px solid #1f2937;border-radius:8px;padding:18px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
  .params h2{font-size:17px;font-weight:600;margin-bottom:12px;color:#e5e7eb}
  .params-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:8px}
  .param-card{background:#0b1120;padding:10px 14px;border-radius:4px;border-left:2px solid #b8963e}
  .param-card b{color:#9ca3af;display:block;margin-bottom:3px;font-size:10px;text-transform:uppercase;letter-spacing:.5px;font-weight:500}
  .param-card span{font-size:13px;color:#e5e7eb}
  .char-profile{background:#111827;border:1px solid #1f2937;border-radius:8px;padding:18px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
  .char-profile h2{font-size:17px;font-weight:600;margin-bottom:12px;color:#e5e7eb}
  .footer{text-align:center;padding:16px;opacity:.5;font-size:11px;color:#6b7280}
  .filters{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .search-input,.filter-select{padding:7px 10px;border-radius:6px;border:1px solid #374151;background:#0b1120;color:#d1d5db;font-size:12px}
  .search-input:focus,.filter-select:focus{outline:none;border-color:#b8963e}
  .empty{padding:30px;text-align:center;color:#6b7280;font-size:13px}
  /* Post Maker */
  .pm-wrap{display:grid;grid-template-columns:1fr 420px;gap:24px}
  @media(max-width:900px){.pm-wrap{grid-template-columns:1fr}}
  .pm-form{background:#111827;border:1px solid #1f2937;border-radius:8px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
  .pm-form h3{font-size:16px;font-weight:600;margin-bottom:14px;color:#e5e7eb;border-bottom:1px solid #1f2937;padding-bottom:8px}
  .field{margin-bottom:14px}
  .field label{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#9ca3af;margin-bottom:5px;font-weight:500}
  .field input,.field textarea,.field select{width:100%;padding:9px 11px;border-radius:6px;border:1px solid #374151;background:#0b1120;color:#d1d5db;font-size:13px;font-family:inherit}
  .field textarea{min-height:90px;resize:vertical}
  .field input:focus,.field textarea:focus,.field select:focus{outline:none;border-color:#b8963e}
  .file-drop{border:2px dashed #374151;border-radius:6px;padding:24px;text-align:center;cursor:pointer;color:#9ca3af;font-size:12px;background:#0b1120}
  .file-drop:hover{border-color:#b8963e;color:#b8963e}
  .file-drop input{display:none}
  .pm-preview{position:sticky;top:20px}
  .phone{width:320px;height:640px;margin:0 auto;background:#000;border-radius:36px;border:8px solid #1f2937;position:relative;overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,.25)}
  .phone .screen{width:100%;height:100%;position:relative;background:#000}
  .phone video{width:100%;height:100%;object-fit:cover}
  .phone .placeholder{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#6b7280;font-size:12px;text-align:center;padding:20px}
  .phone .overlay{position:absolute;inset:0;pointer-events:none}
  .phone .right{position:absolute;right:8px;bottom:90px;display:flex;flex-direction:column;gap:18px;align-items:center;color:#fff;font-size:11px}
  .phone .right .ico{width:40px;height:40px;border-radius:50%;background:rgba(255,255,255,.18);display:flex;align-items:center;justify-content:center;font-size:18px;backdrop-filter:blur(4px)}
  .phone .right .lbl{font-size:10px;margin-top:2px;text-align:center}
  .phone .bottom{position:absolute;left:10px;right:50px;bottom:12px;color:#fff;font-size:12px}
  .phone .bottom .user{font-weight:600;margin-bottom:4px;font-size:13px}
  .phone .bottom .caption{margin-bottom:4px;line-height:1.3}
  .phone .bottom .tags{color:#93c5fd;font-size:11px}
  .phone .bottom .sound{display:flex;align-items:center;gap:6px;font-size:11px;margin-top:6px;opacity:.9}
  .phone .bottom .sound .disc{width:28px;height:28px;border-radius:50%;background:#374151;display:flex;align-items:center;justify-content:center;font-size:14px}
  .phone .topbar{position:absolute;top:0;left:0;right:0;padding:10px 14px;display:flex;justify-content:space-between;color:#fff;font-size:12px;z-index:5;background:linear-gradient(180deg,rgba(0,0,0,.4),transparent)}
  .pm-actions{display:flex;gap:8px;margin-top:16px;flex-wrap:wrap}
  @media(max-width:768px){header h1{font-size:18px}table{font-size:10px}.phone{width:280px;height:560px}}
</style>
</head>
<body>
<div class="container">
<div class="floating-logo">
  <img src="data:image/jpeg;base64,__LOGO__" alt="20 günə dil öyrənmə">
</div>
<header>
  <h1>Postplan — 20 günlük dil öyrənmə (TikTok)</h1>
  <p>Kitabın tanıtımı və satış üçün motivasiyalı postplan (Azərbaycan dilində). Bütün sahələr redaktə edilə bilər; dəyişikliklər brauzerdə saxlanılır.</p>
  <div class="meta">
    <span>15.11.2026 — 28.02.2027</span>
    <span>Satış məqsədi</span>
    <span>Yalnız orqanik</span>
  </div>
</header>

<div class="tabs">
  <div class="tab active" data-tab="t1" onclick="switchTab('t1')">Postplan 1 — General <span id="count1" style="opacity:.7;font-size:11px">(0)</span></div>
  <div class="tab" data-tab="t2" onclick="switchTab('t2')">Postplan 2 — Xarakter <span id="count2" style="opacity:.7;font-size:11px">(0)</span></div>
  <div class="tab" data-tab="t3" onclick="switchTab('t3')">Post Maker (TikTok)</div>
</div>

<div id="t1" class="tabpane active">
  <div class="params">
    <h2>Kampaniya parametrləri</h2>
    <div class="params-grid">
      <div class="param-card"><b>Platforma</b><span>TikTok</span></div>
      <div class="param-card"><b>Hədəf auditoriya</b><span>Geniş auditoriya</span></div>
      <div class="param-card"><b>Ton</b><span>Motivasiyalı</span></div>
      <div class="param-card"><b>Əsas məqsəd</b><span>Kitab satışı</span></div>
      <div class="param-card"><b>Tezlik</b><span>Həftədə 3-4 post</span></div>
      <div class="param-card"><b>Məzmun formatı</b><span>Qarışıq (qısa + orta + uzun)</span></div>
      <div class="param-card"><b>Yayım vaxtı</b><span>18:00–21:00</span></div>
      <div class="param-card"><b>Audio</b><span>Səs + fon musiqi</span></div>
      <div class="param-card"><b>Subtitr</b><span>Açar sözlər</span></div>
      <div class="param-card"><b>Satış kanalı</b><span>DM (mesaj)</span></div>
      <div class="param-card"><b>Reklam</b><span>Yalnız orqanik</span></div>
      <div class="param-card"><b>Bonus</b><span>Yoxdur</span></div>
      <div class="param-card"><b>Zəmanət</b><span>Yoxdur</span></div>
      <div class="param-card"><b>Vizual stil</b><span>Minimalist</span></div>
      <div class="param-card"><b>Dillər</b><span>İngilis + Rus</span></div>
    </div>
  </div>
  <div class="plan-section">
    <h2>Postplan 1 — General <span class="toolbar">
      <input type="text" class="search-input" id="search1" placeholder="Axtar..." oninput="render1()">
      <select class="filter-select" id="filter1" onchange="render1()"><option value="">Bütün mövzular</option></select>
      <button class="btn btn-primary" onclick="addPost(1)">+ Yeni post</button>
      <button class="btn btn-success" onclick="exportExcel(1)">Excel eksport</button>
      <button class="btn" onclick="exportData(1)">JSON eksport</button>
      <button class="btn" onclick="importData(1)">İdxal</button>
      <button class="btn btn-danger" onclick="resetData(1)">Sıfırla</button>
    </span></h2>
    <div class="table-wrap">
      <table class="post-table">
        <thead><tr>
          <th>#</th><th>Tarix</th><th>Həftə</th><th>Məzmun növü</th><th>Mövzu</th><th>Başlıq (Hook)</th>
          <th>Vaxt</th><th>Audio</th><th>Subtitr</th><th>Kanal</th>
          <th>Ssenariy (Vizual / Səs / Mətn)</th><th>Çağırış (CTA)</th><th>Hashtag-lar</th><th>Məqsəd</th><th>Status</th><th>Əməliyyat</th>
        </tr></thead>
        <tbody id="tbody1"></tbody>
      </table>
    </div>
  </div>
</div>

<div id="t2" class="tabpane">
  <div class="char-profile">
    <h2>Xarakter profili</h2>
    <div class="params-grid">
      <div class="param-card"><b>Ad</b><span>Aysel (nümunə)</span></div>
      <div class="param-card"><b>Yaş</b><span>22</span></div>
      <div class="param-card"><b>Status</b><span>Tələbə</span></div>
      <div class="param-card"><b>Məqsəd</b><span>Karyera üçün ingilis dili</span></div>
      <div class="param-card"><b>Başlanğıc səviyyə</b><span>A1 (başlanğıc)</span></div>
      <div class="param-card"><b>Hədəf səviyyə</b><span>B1 (orta)</span></div>
      <div class="param-card"><b>Gündəlik vaxt</b><span>20 dəqiqə</span></div>
      <div class="param-card"><b>Metod</b><span>20 günlük metodika</span></div>
      <div class="param-card"><b>Müddət</b><span>20 gün</span></div>
      <div class="param-card"><b>Tezlik</b><span>Hər gün 1 post</span></div>
    </div>
  </div>
  <div class="plan-section">
    <h2>Postplan 2 — Xarakter <span class="toolbar">
      <input type="text" class="search-input" id="search2" placeholder="Axtar..." oninput="render2()">
      <select class="filter-select" id="filter2" onchange="render2()"><option value="">Bütün mövzular</option></select>
      <button class="btn btn-primary" onclick="addPost(2)">+ Yeni post</button>
      <button class="btn btn-success" onclick="exportExcel(2)">Excel eksport</button>
      <button class="btn" onclick="exportData(2)">JSON eksport</button>
      <button class="btn" onclick="importData(2)">İdxal</button>
      <button class="btn btn-danger" onclick="resetData(2)">Sıfırla</button>
    </span></h2>
    <div class="table-wrap">
      <table class="post-table">
        <thead><tr>
          <th>#</th><th>Gün</th><th>Tarix</th><th>Məzmun növü</th><th>Mövzu</th><th>Başlıq (Hook)</th>
          <th>Vaxt</th><th>Audio</th><th>Subtitr</th><th>Kanal</th>
          <th>Ssenariy (Vizual / Səs / Mətn)</th><th>Çağırış (CTA)</th><th>Hashtag-lar</th><th>Məqsəd</th><th>Status</th><th>Əməliyyat</th>
        </tr></thead>
        <tbody id="tbody2"></tbody>
      </table>
    </div>
  </div>
</div>

<div id="t3" class="tabpane">
  <div class="pm-wrap">
    <div class="pm-form">
      <h3>Post Maker — TikTok üçün</h3>
      <div class="field">
        <label>Mövcud postdan seç (avtomatik doldur)</label>
        <select id="pmPostSelect" onchange="fillFromPost()">
          <option value="">— Post seçin —</option>
          <optgroup label="Postplan 1 — General" id="pmOpt1"></optgroup>
          <optgroup label="Postplan 2 — Xarakter" id="pmOpt2"></optgroup>
        </select>
      </div>
      <div class="field">
        <label>Video yüklə</label>
        <div class="file-drop" onclick="document.getElementById('pmFile').click()">
          <input type="file" id="pmFile" accept="video/*" onchange="loadVideo(this)">
          <div id="pmFileLabel">Video seçmək üçün kliklə (MP4, MOV)</div>
        </div>
      </div>
      <div class="field">
        <label>İstifadəçi adı (@)</label>
        <input type="text" id="pmUser" value="@kitab20gun" oninput="updatePreview()">
      </div>
      <div class="field">
        <label>Açıqlama (caption)</label>
        <textarea id="pmCaption" oninput="updatePreview()"></textarea>
      </div>
      <div class="field">
        <label>Hashtag-lar</label>
        <input type="text" id="pmTags" oninput="updatePreview()">
      </div>
      <div class="field">
        <label>Səs / musiqi adı</label>
        <input type="text" id="pmSound" oninput="updatePreview()">
      </div>
      <div class="field">
        <label>Məxfilik</label>
        <select id="pmPrivacy" onchange="updatePreview()">
          <option value="public">Hər kəs (Public)</option>
          <option value="friends">Dostlar (Friends)</option>
          <option value="private">Yalnız mən (Private)</option>
        </select>
      </div>
      <div class="field">
        <label>Parametrlər</label>
        <label style="display:block;margin-bottom:4px"><input type="checkbox" id="pmComments" checked onchange="updatePreview()"> Şərhlərə icazə</label>
        <label style="display:block;margin-bottom:4px"><input type="checkbox" id="pmDuet" onchange="updatePreview()"> Duet icazəsi</label>
        <label style="display:block;margin-bottom:4px"><input type="checkbox" id="pmStitch" onchange="updatePreview()"> Stitch icazəsi</label>
      </div>
      <div class="pm-actions">
        <button class="btn btn-success" onclick="exportPost()">⬇ Eksport (yüklə)</button>
        <button class="btn btn-danger" onclick="clearPost()">Təmizlə</button>
      </div>
      <div style="margin-top:12px;font-size:11px;color:#6b7280">
        Eksport: video faylı (əgər seçilibsə) və post məlumatları (caption, hashtag, parametrlər) ayrıca .txt faylı kimi yüklənir.
      </div>
    </div>
    <div class="pm-preview">
      <div class="phone">
        <div class="screen">
          <div class="topbar"><span>Following</span><span>For You</span></div>
          <video id="pmVideo" style="display:none" muted loop playsinline></video>
          <div class="placeholder" id="pmPlaceholder">Video burada görünəcək.<br>Yükləmək üçün sol paneldən video seçin.</div>
          <div class="overlay">
            <div class="right">
              <div><div class="ico">♥</div><div class="lbl">12.4K</div></div>
              <div><div class="ico">💬</div><div class="lbl">342</div></div>
              <div><div class="ico">↗</div><div class="lbl">Share</div></div>
              <div><div class="ico">🔖</div><div class="lbl">Save</div></div>
            </div>
            <div class="bottom">
              <div class="user" id="pmPreviewUser">@kitab20gun</div>
              <div class="caption" id="pmPreviewCaption"></div>
              <div class="tags" id="pmPreviewTags"></div>
              <div class="sound"><div class="disc">♪</div><span id="pmPreviewSound">Orijinal səs</span></div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</div>


</div>

<script>
const STORAGE1 = 'postplan1_data_v4';
const STORAGE2 = 'postplan2_data_v4';
const DEFAULT1 = __DATA1__;
const DEFAULT2 = __DATA2__;

let data1 = loadOrInit(STORAGE1, DEFAULT1);
let data2 = loadOrInit(STORAGE2, DEFAULT2);

function loadOrInit(key, def){
  try {
    const raw = localStorage.getItem(key);
    if (raw) { const parsed = JSON.parse(raw); if (Array.isArray(parsed)) return parsed; }
  } catch(e){}
  return JSON.parse(JSON.stringify(def));
}
function save(key, data){ localStorage.setItem(key, JSON.stringify(data)); }

function switchTab(t){
  document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active', x.dataset.tab===t));
  document.querySelectorAll('.tabpane').forEach(x=>x.classList.toggle('active', x.id===t));
  updateCounts();
  if (t==='t3') populatePostSelect();
}

const STATUS_OPTS = ['Planlaşdırılıb','Hazırlanır','Yayımlandı','Ləğv edildi'];
function statusSelect(val, idx, plan){
  return `<select class="status-select" onchange="updateField(${plan},${idx},'status',this.value)">${STATUS_OPTS.map(s=>`<option value="${s}"${s===val?' selected':''}>${s}</option>`).join('')}</select>`;
}

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function toISO(s){ const m=(s||'').match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); if(!m) return ''; return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'); }
function toDMY(s){ const m=(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/); if(!m) return s; return m[3]+'.'+m[2]+'.'+m[1]; }

function populateFilter(id, data){
  const sel = document.getElementById(id);
  const cur = sel.value;
  const topics = [...new Set(data.map(p=>p.topic))].sort();
  sel.innerHTML = '<option value="">Bütün mövzular</option>' + topics.map(t=>`<option value="${esc(t)}">${esc(t)}</option>`).join('');
  sel.value = cur;
}

function render1(){
  const tbody = document.getElementById('tbody1');
  const q = (document.getElementById('search1').value||'').toLowerCase();
  const topic = document.getElementById('filter1').value||'';
  populateFilter('filter1', data1);
  let html = '';
  data1.forEach((p,i)=>{
    if (q && !JSON.stringify(p).toLowerCase().includes(q)) return;
    if (topic && p.topic !== topic) return;
    html += `<tr data-idx="${i}">
      <td>${i+1}</td>
      <td><input class="cell-input" type="date" value="${toISO(p.date)}" onchange="updateField(1,${i},'date',toDMY(this.value))"></td>
      <td><input class="cell-input" type="number" value="${esc(p.week)}" onchange="updateField(1,${i},'week',this.value)"></td>
      <td contenteditable="true" onblur="updateField(1,${i},'type',this.innerText)">${esc(p.type)}</td>
      <td contenteditable="true" onblur="updateField(1,${i},'topic',this.innerText)">${esc(p.topic)}</td>
      <td class="hook" contenteditable="true" onblur="updateField(1,${i},'hook',this.innerText)">${esc(p.hook)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(1,${i},'vaxt',this.innerText)"><b>Vaxt:</b> ${esc(p.vaxt)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(1,${i},'audio',this.innerText)"><b>Audio:</b> ${esc(p.audio)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(1,${i},'subtitr',this.innerText)"><b>Subtitr:</b> ${esc(p.subtitr)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(1,${i},'kanal',this.innerText)"><b>Kanal:</b> ${esc(p.kanal)}</td>
      <td class="scenario">
        <details>
          <summary>Ssenariyə bax</summary>
          <div class="sc-block"><b>Vizual:</b> <span contenteditable="true" onblur="updateField(1,${i},'visual',this.innerText)">${esc(p.visual)}</span></div>
          <div class="sc-block"><b>Səs (rəvayət):</b> <span contenteditable="true" onblur="updateField(1,${i},'speech',this.innerText)">${esc(p.speech)}</span></div>
          <div class="sc-block"><b>Ekranda mətn:</b> <span contenteditable="true" onblur="updateField(1,${i},'onscreen',this.innerText)">${esc(p.onscreen)}</span></div>
        </details>
      </td>
      <td class="cta" contenteditable="true" onblur="updateField(1,${i},'cta',this.innerText)">${esc(p.cta)}</td>
      <td class="tags" contenteditable="true" onblur="updateField(1,${i},'hashtags',this.innerText)">${esc(p.hashtags)}</td>
      <td contenteditable="true" onblur="updateField(1,${i},'goal',this.innerText)">${esc(p.goal)}</td>
      <td>${statusSelect(p.status,i,1)}</td>
      <td><div class="row-actions"><button class="icon-btn del" onclick="deletePost(1,${i})" title="Sil">✕</button></div></td>
    </tr>`;
  });
  tbody.innerHTML = html || '<tr><td colspan="16" class="empty">Post yoxdur. "+ Yeni post" ilə əlavə et.</td></tr>';
  updateCounts();
}

function render2(){
  const tbody = document.getElementById('tbody2');
  const q = (document.getElementById('search2').value||'').toLowerCase();
  const topic = document.getElementById('filter2').value||'';
  populateFilter('filter2', data2);
  let html = '';
  data2.forEach((p,i)=>{
    if (q && !JSON.stringify(p).toLowerCase().includes(q)) return;
    if (topic && p.topic !== topic) return;
    html += `<tr data-idx="${i}">
      <td>${i+1}</td>
      <td><input class="cell-input" type="number" value="${esc(p.day)}" onchange="updateField(2,${i},'day',this.value)"></td>
      <td><input class="cell-input" type="date" value="${toISO(p.date)}" onchange="updateField(2,${i},'date',toDMY(this.value))"></td>
      <td contenteditable="true" onblur="updateField(2,${i},'type',this.innerText)">${esc(p.type)}</td>
      <td contenteditable="true" onblur="updateField(2,${i},'topic',this.innerText)">${esc(p.topic)}</td>
      <td class="hook" contenteditable="true" onblur="updateField(2,${i},'hook',this.innerText)">${esc(p.hook)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(2,${i},'vaxt',this.innerText)"><b>Vaxt:</b> ${esc(p.vaxt)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(2,${i},'audio',this.innerText)"><b>Audio:</b> ${esc(p.audio)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(2,${i},'subtitr',this.innerText)"><b>Subtitr:</b> ${esc(p.subtitr)}</td>
      <td class="sc-meta" contenteditable="true" onblur="updateField(2,${i},'kanal',this.innerText)"><b>Kanal:</b> ${esc(p.kanal)}</td>
      <td class="scenario">
        <details>
          <summary>Ssenariyə bax</summary>
          <div class="sc-block"><b>Vizual:</b> <span contenteditable="true" onblur="updateField(2,${i},'visual',this.innerText)">${esc(p.visual)}</span></div>
          <div class="sc-block"><b>Səs (rəvayət):</b> <span contenteditable="true" onblur="updateField(2,${i},'speech',this.innerText)">${esc(p.speech)}</span></div>
          <div class="sc-block"><b>Ekranda mətn:</b> <span contenteditable="true" onblur="updateField(2,${i},'onscreen',this.innerText)">${esc(p.onscreen)}</span></div>
        </details>
      </td>
      <td class="cta" contenteditable="true" onblur="updateField(2,${i},'cta',this.innerText)">${esc(p.cta)}</td>
      <td class="tags" contenteditable="true" onblur="updateField(2,${i},'hashtags',this.innerText)">${esc(p.hashtags)}</td>
      <td contenteditable="true" onblur="updateField(2,${i},'goal',this.innerText)">${esc(p.goal)}</td>
      <td>${statusSelect(p.status,i,2)}</td>
      <td><div class="row-actions"><button class="icon-btn del" onclick="deletePost(2,${i})" title="Sil">✕</button></div></td>
    </tr>`;
  });
  tbody.innerHTML = html || '<tr><td colspan="16" class="empty">Post yoxdur. "+ Yeni post" ilə əlavə et.</td></tr>';
  updateCounts();
}

function updateCounts(){
  const c1 = document.getElementById('count1');
  const c2 = document.getElementById('count2');
  if (c1) c1.textContent = data1.length + ' post';
  if (c2) c2.textContent = data2.length + ' post';
}

function updateField(plan, idx, field, val){
  const arr = plan===1 ? data1 : data2;
  if (!arr[idx]) return;
  let v = val;
  if (field==='vaxt') v = val.replace(/^Vaxt:\s*/i,'').trim();
  if (field==='audio') v = val.replace(/^Audio:\s*/i,'').trim();
  if (field==='subtitr') v = val.replace(/^Subtitr:\s*/i,'').trim();
  if (field==='kanal') v = val.replace(/^Kanal:\s*/i,'').trim();
  arr[idx][field] = v.trim();
  save(plan===1 ? STORAGE1 : STORAGE2, arr);
  updateCounts();
}

function addPost(plan){
  const arr = plan===1 ? data1 : data2;
  const blank = plan===1
    ? {date:"",week:"",type:"Qısa video (15-30s)",topic:"",hook:"",visual:"",speech:"",onscreen:"",cta:"",hashtags:"",goal:"",status:"Planlaşdırılıb",vaxt:"18:00-21:00",audio:"Səs + fon musiqi",subtitr:"Açar sözlər",kanal:"DM (mesaj)"}
    : {day:arr.length+1,date:"",type:"Qısa video (15-30s)",topic:"",hook:"",visual:"",speech:"",onscreen:"",cta:"",hashtags:"",goal:"",status:"Planlaşdırılıb",vaxt:"18:00-21:00",audio:"Səs + fon musiqi",subtitr:"Açar sözlər",kanal:"DM (mesaj)"};
  arr.push(blank);
  save(plan===1 ? STORAGE1 : STORAGE2, arr);
  if (plan===1) render1(); else render2();
}

function deletePost(plan, idx){
  if (!confirm('Bu post silinsin?')) return;
  const arr = plan===1 ? data1 : data2;
  arr.splice(idx,1);
  save(plan===1 ? STORAGE1 : STORAGE2, arr);
  if (plan===1) render1(); else render2();
}

function resetData(plan){
  if (!confirm('Bütün dəyişikliklər sıfırlansın və ilkin məlumat bərpa olunsun?')) return;
  if (plan===1){ localStorage.removeItem(STORAGE1); data1 = JSON.parse(JSON.stringify(DEFAULT1)); }
  else { localStorage.removeItem(STORAGE2); data2 = JSON.parse(JSON.stringify(DEFAULT2)); }
  if (plan===1) render1(); else render2();
}

function exportData(plan){
  const arr = plan===1 ? data1 : data2;
  const blob = new Blob([JSON.stringify(arr,null,2)], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = (plan===1?'postplan1':'postplan2') + '.json';
  a.click(); URL.revokeObjectURL(url);
}

function importData(plan){
  const inp = document.createElement('input');
  inp.type='file'; inp.accept='application/json';
  inp.onchange = e=>{
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = ev=>{
      try {
        const parsed = JSON.parse(ev.target.result);
        if (!Array.isArray(parsed)) throw new Error('Array deyil');
        if (plan===1){ data1 = parsed; save(STORAGE1,data1); render1(); }
        else { data2 = parsed; save(STORAGE2,data2); render2(); }
      } catch(err){ alert('İdxal xətası: ' + err.message); }
    };
    r.readAsText(f);
  };
  inp.click();
}

// Excel export via ExcelJS (light professional theme)
async function exportExcel(plan){
  if (typeof ExcelJS === 'undefined'){ alert('ExcelJS yüklənmədi. İnternet bağlantısını yoxlayın.'); return; }
  const arr = plan===1 ? data1 : data2;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Postplan';
  wb.created = new Date();
  const sheetName = plan===1 ? 'Postplan General' : 'Postplan Xarakter';
  const ws = wb.addWorksheet(sheetName);
  const headers = plan===1
    ? ['Tarix','Həftə','Platforma','Vaxt','Audio','Subtitr','Satış kanalı','Məzmun növü','Mövzu','Başlıq (Hook)','Vizual','Səs (rəvayət)','Ekranda mətn','Çağırış (CTA)','Hashtag-lar','Məqsəd','Status']
    : ['Gün','Tarix','Platforma','Vaxt','Audio','Subtitr','Satış kanalı','Məzmun növü','Mövzu','Başlıq (Hook)','Vizual','Səs (rəvayət)','Ekranda mətn','Çağırış (CTA)','Hashtag-lar','Məqsəd','Status'];
  // Title row
  ws.addRow([plan===1 ? 'Postplan 1 — General (TikTok)' : 'Postplan 2 — Xarakter (TikTok)']);
  ws.mergeCells(1,1,1,headers.length);
  const titleCell = ws.getCell('A1');
  titleCell.font = {bold:true,size:14,color:{argb:'FF1F2937'}};
  titleCell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FFF3F4F6'}};
  titleCell.alignment = {horizontal:'center',vertical:'middle'};
  ws.getRow(1).height = 24;
  // Subtitle
  ws.addRow(['Yaradılma: ' + new Date().toLocaleString('az-AZ') + ' | Post sayı: ' + arr.length]);
  ws.mergeCells(2,1,2,headers.length);
  const subCell = ws.getCell('A2');
  subCell.font = {size:10,italic:true,color:{argb:'FF6B7280'}};
  subCell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FFF9FAFB'}};
  subCell.alignment = {horizontal:'center',vertical:'middle'};
  // Header row
  const headerRow = ws.addRow(headers);
  headerRow.height = 30;
  headerRow.eachCell(cell=>{
    cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FF1F2937'}};
    cell.font = {bold:true,color:{argb:'FFFBBF24'},size:10};
    cell.border = {top:{style:'thin',color:{argb:'FF374151'}},bottom:{style:'thin',color:{argb:'FF374151'}},left:{style:'thin',color:{argb:'FF374151'}},right:{style:'thin',color:{argb:'FF374151'}}};
    cell.alignment = {horizontal:'center',vertical:'middle',wrapText:true};
  });
  // Data rows
  arr.forEach(p=>{
    const rowVals = plan===1
      ? [p.date,p.week,'TikTok',p.vaxt,p.audio,p.subtitr,p.kanal,p.type,p.topic,p.hook,p.visual,p.speech,p.onscreen,p.cta,p.hashtags,p.goal,p.status]
      : [p.day,p.date,'TikTok',p.vaxt,p.audio,p.subtitr,p.kanal,p.type,p.topic,p.hook,p.visual,p.speech,p.onscreen,p.cta,p.hashtags,p.goal,p.status];
    const row = ws.addRow(rowVals);
    row.height = 58;
    row.eachCell((cell, colNum)=>{
      cell.border = {top:{style:'thin',color:{argb:'FFE5E7EB'}},bottom:{style:'thin',color:{argb:'FFE5E7EB'}},left:{style:'thin',color:{argb:'FFE5E7EB'}},right:{style:'thin',color:{argb:'FFE5E7EB'}}};
      cell.alignment = {vertical:'top',wrapText:true};
      cell.font = {size:10,color:{argb:'FF1F2937'}};
      cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}};
    });
    // status cell color
    const statusIdx = headers.length;
    const sc = row.getCell(statusIdx);
    const colors = {'Planlaşdırılıb':'E5E7EB','Hazırlanır':'FEF3C7','Yayımlandı':'D1FAE5','Ləğv edildi':'FEE2E2'};
    const fcolors = {'Planlaşdırılıb':'374151','Hazırlanır':'92400E','Yayımlandı':'065F46','Ləğv edildi':'991B1B'};
    if (colors[p.status]){
      sc.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FF'+colors[p.status]}};
      sc.font = {size:10,bold:true,color:{argb:'FF'+fcolors[p.status]}};
      sc.alignment = {vertical:'middle',horizontal:'center',wrapText:true};
    }
  });
  // Column widths
  const widths = [12,7,10,13,20,14,15,16,24,32,38,48,24,28,30,18,16];
  ws.columns.forEach((col,i)=>{ if (widths[i]) col.width = widths[i]; });
  // Freeze panes
  ws.views = [{state:'frozen', xSplit:0, ySplit:3}];
  // Auto filter
  ws.autoFilter = { from:{row:3,column:1}, to:{row:arr.length+3,column:headers.length} };
  // Conditional formatting
  const statusCol = String.fromCharCode(64 + headers.length);
  ws.addConditionalFormatting({
    ref: `${statusCol}4:${statusCol}${arr.length+3}`,
    rules: [
      {type:'cellIs',operator:'equal',formulae:['"Planlaşdırılıb"'],style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:'FFE5E7EB'}},font:{color:{argb:'FF374151'},bold:true}}},
      {type:'cellIs',operator:'equal',formulae:['"Hazırlanır"'],style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:'FFFEF3C7'}},font:{color:{argb:'FF92400E'},bold:true}}},
      {type:'cellIs',operator:'equal',formulae:['"Yayımlandı"'],style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:'FFD1FAE5'}},font:{color:{argb:'FF065F46'},bold:true}}},
      {type:'cellIs',operator:'equal',formulae:['"Ləğv edildi"'],style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:'FFFEE2E2'}},font:{color:{argb:'FF991B1B'},bold:true}}},
    ]
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = (plan===1?'Postplan_General':'Postplan_Xarakter') + '.xlsx';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Post Maker logic
let pmVideoUrl = null;
function loadVideo(inp){
  const f = inp.files[0];
  if (!f) return;
  if (pmVideoUrl) URL.revokeObjectURL(pmVideoUrl);
  pmVideoUrl = URL.createObjectURL(f);
  const v = document.getElementById('pmVideo');
  v.src = pmVideoUrl; v.style.display='block';
  document.getElementById('pmPlaceholder').style.display='none';
  document.getElementById('pmFileLabel').textContent = f.name + ' (' + (f.size/1024/1024).toFixed(2) + ' MB)';
  v.play().catch(()=>{});
}
function updatePreview(){
  const user = document.getElementById('pmUser').value || '@kitab20gun';
  const caption = document.getElementById('pmCaption').value || '';
  const tags = document.getElementById('pmTags').value || '';
  const sound = document.getElementById('pmSound').value || 'Orijinal səs';
  document.getElementById('pmPreviewUser').textContent = user;
  document.getElementById('pmPreviewCaption').textContent = caption;
  document.getElementById('pmPreviewTags').textContent = tags;
  document.getElementById('pmPreviewSound').textContent = sound;
}

function populatePostSelect(){
  const sel = document.getElementById('pmPostSelect');
  const cur = sel.value;
  let html = '<option value="">— Post seçin —</option>';
  html += '<optgroup label="Postplan 1 — General">';
  data1.forEach((p,i)=>{ html += `<option value="1_${i}">${esc(p.date)} — ${esc(p.topic)}: ${esc(p.hook).slice(0,40)}</option>`; });
  html += '</optgroup>';
  html += '<optgroup label="Postplan 2 — Xarakter">';
  data2.forEach((p,i)=>{ html += `<option value="2_${i}">${esc(p.date)} — ${esc(p.topic)}: ${esc(p.hook).slice(0,40)}</option>`; });
  html += '</optgroup>';
  sel.innerHTML = html;
  sel.value = cur;
}

function fillFromPost(){
  const val = document.getElementById('pmPostSelect').value;
  if (!val) return;
  const [plan, idx] = val.split('_');
  const arr = plan==='1' ? data1 : data2;
  const p = arr[parseInt(idx)];
  if (!p) return;
  document.getElementById('pmCaption').value = (p.hook||'') + (p.speech ? '\n\n' + p.speech : '');
  document.getElementById('pmTags').value = p.hashtags || '';
  document.getElementById('pmSound').value = p.audio || 'Orijinal səs';
  updatePreview();
}

function exportPost(){
  const fileInput = document.getElementById('pmFile');
  const f = fileInput.files[0];
  if (f){
    const url = URL.createObjectURL(f);
    const a = document.createElement('a');
    a.href = url; a.download = f.name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  const user = document.getElementById('pmUser').value;
  const caption = document.getElementById('pmCaption').value;
  const tags = document.getElementById('pmTags').value;
  const sound = document.getElementById('pmSound').value;
  const privacy = document.getElementById('pmPrivacy').value;
  const comments = document.getElementById('pmComments').checked;
  const duet = document.getElementById('pmDuet').checked;
  const stitch = document.getElementById('pmStitch').checked;
  const meta = `TikTok Post Metadata\n====================\nİstifadəçi: ${user}\nMəxfilik: ${privacy}\nŞərhlər: ${comments}\nDuet: ${duet}\nStitch: ${stitch}\n\nAçıqlama:\n${caption}\n\nHashtag-lar:\n${tags}\n\nSəs:\n${sound}\n`;
  const blob = new Blob([meta], {type:'text/plain;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'tiktok_post_meta.txt';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  alert('Eksport tamamlandı: video (əgər seçilibsə) və post məlumatları (txt) yükləndi.');
}
function clearPost(){
  document.getElementById('pmCaption').value = '';
  document.getElementById('pmTags').value = '';
  document.getElementById('pmSound').value = '';
  document.getElementById('pmFile').value = '';
  document.getElementById('pmFileLabel').textContent = 'Video seçmək üçün kliklə (MP4, MOV)';
  document.getElementById('pmPostSelect').value = '';
  if (pmVideoUrl){ URL.revokeObjectURL(pmVideoUrl); pmVideoUrl=null; }
  const v = document.getElementById('pmVideo');
  v.style.display='none'; v.src='';
  document.getElementById('pmPlaceholder').style.display='flex';
  updatePreview();
}

document.addEventListener('DOMContentLoaded', ()=>{
  render1(); render2(); updatePreview();
});
</script>
</body>
</html>'''

html_doc = html_doc.replace('__DATA1__', data1_json).replace('__DATA2__', data2_json).replace('__LOGO__', logo_b64)

out = r"C:\Users\ANCL\Desktop\20 gun\Postplan.html"
with open(out, "w", encoding="utf-8") as f:
    f.write(html_doc)
print("Saved:", out)
print("Posts1:", len(posts1), "Posts2:", len(char_posts))