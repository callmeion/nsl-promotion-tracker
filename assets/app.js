(() => {
const C = window.APP_CONFIG, $ = s => document.querySelector(s);
let rows = [];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const sha = async t => [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(t)))].map(b => b.toString(16).padStart(2,'0')).join('');
const t = v => (v ?? '').toString().trim();
const num = v => { const n = parseFloat(t(v).replace(/,/g,'')); return isNaN(n) ? 0 : n; };
const pct = (a, b) => b ? Math.round(a * 100 / b) : 0;

/* ---------- Auth ---------- */
const showApp = () => { $('#login').hidden = true; $('#app').hidden = false; if (!rows.length) load(); };
if (sessionStorage.getItem('nsl_auth') === '1') showApp();
$('#loginForm').addEventListener('submit', async e => {
  e.preventDefault();
  if (!crypto.subtle) { $('#loginErr').textContent = 'ต้องเปิดผ่าน https:// หรือ localhost เท่านั้น'; return; }
  if (await sha($('#pw').value) === C.passwordHash) { sessionStorage.setItem('nsl_auth','1'); $('#pw').value=''; showApp(); }
  else $('#loginErr').textContent = 'รหัสผ่านไม่ถูกต้อง';
});
$('#btnLogout').onclick = () => { sessionStorage.removeItem('nsl_auth'); location.reload(); };
$('#btnPrint').onclick = () => window.print();

/* ---------- Excel ---------- */
function toISO(v) {
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
  const s = t(v), m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) { let y = +m[3]; if (y > 2400) y -= 543; return `${y}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; }
  return s;
}
const fmtDate = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? `${m[3]}/${m[2]}/${m[1]}` : (iso || ''); };
const today = () => new Date().toISOString().slice(0,10);

function parse(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  return data.filter(r => /^\d+$/.test(t(r[0])) && t(r[1])).map(r => ({
    no: t(r[0]), id: t(r[1]), prefix: t(r[2]), first: t(r[3]), last: t(r[4]),
    div: t(r[5]) || '-', dept: t(r[6]), branch: t(r[7]), start: toISO(r[8]),
    ty: t(r[9]), tm: t(r[10]), td: t(r[11]), reqY: t(r[12]), sumY: t(r[13]),
    hrs: t(r[14]), reqH: t(r[15]), sumH: t(r[16]),
    oldPos: t(r[17]), oldLv: t(r[18]), newPos: t(r[19]), newLv: t(r[20]),
    status: t(r[21]) || C.statuses[0].key, statusDate: toISO(r[22]), intDate: toISO(r[23]), note: t(r[24])
  }));
}
function setRows(wb, label) {
  rows = parse(wb);
  $('#srcInfo').textContent = `${label} · ${rows.length} รายการ`;
  $('#fDiv').innerHTML = '<option value="">ทุกฝ่าย</option>' + [...new Set(rows.map(r => r.div))].sort().map(d => `<option>${esc(d)}</option>`).join('');
  $('#fStatus').innerHTML = '<option value="">ทุกสถานะ</option>' + C.statuses.map(s => `<option>${esc(s.key)}</option>`).join('');
  render();
}
async function load() {
  try {
    const res = await fetch(C.dataFile + '?t=' + Date.now());
    if (!res.ok) throw new Error(res.status);
    const lm = res.headers.get('last-modified');
    setRows(XLSX.read(await res.arrayBuffer(), { cellDates: true }), C.dataFile + (lm ? ` · อัปเดตไฟล์ ${new Date(lm).toLocaleDateString('th-TH')}` : ''));
    $('#banner').hidden = true;
  } catch (err) {
    $('#srcInfo').textContent = 'โหลดข้อมูลไม่สำเร็จ'; $('#banner').hidden = false;
    $('#banner').textContent = `โหลด ${C.dataFile} อัตโนมัติไม่ได้ (${err.message}) — ตรวจว่าไฟล์อยู่ใน repo และเปิดผ่าน GitHub Pages / local server (ไม่ใช่ file://) หรือกด "โหลดไฟล์ Excel"`;
  }
}
$('#fileIn').onchange = async e => { const f = e.target.files[0]; if (f) { setRows(XLSX.read(await f.arrayBuffer(), { cellDates: true }), f.name); $('#banner').hidden = true; } };

/* ---------- Helpers ---------- */
const col = s => (C.statuses.find(x => x.key === s) || { color: 'gray' }).color;
const crit = v => /ไม่/.test(v) ? 'red' : /ผ่าน|✓|ok|yes/i.test(v) ? 'green' : 'gray';
const okY = r => crit(r.sumY) === 'green', okH = r => crit(r.sumH) === 'green';
const bothOk = r => okY(r) && okH(r);
const years = r => num(r.ty) + num(r.tm) / 12;
const SCOL = { gray:'#7b8794', blue:'#2b6fd6', amber:'#d98e04', green:'#1f9d63', red:'#d64545' };
const barRow = (label, segs, total) => `<div class="bar-row"><span class="lb" title="${esc(label)}">${esc(label)}</span>
  <div class="bar">${segs.filter(s => s.n).map(s => `<span class="${s.cls}" style="width:${pct(s.n, total)}%" title="${esc(s.t||'')}: ${s.n}"></span>`).join('')}</div><b>${total}</b></div>`;

/* ---------- Render ---------- */
function render() {
  const all = rows, N = all.length, cnt = k => all.filter(r => r.status === k).length;
  const stKeys = C.statuses.map(s => s.key);
  const passed = cnt('ผ่าน'), failed = cnt('ไม่ผ่าน'), waiting = cnt('สัมภาษณ์แล้ว (รอผล)');
  const interviewed = passed + failed + waiting, scheduled = cnt('นัดสัมภาษณ์แล้ว'), notYet = cnt('ยังไม่สัมภาษณ์');
  const decided = passed + failed, both = all.filter(bothOk).length;
  const avgTen = N ? (all.reduce((a, r) => a + years(r), 0) / N).toFixed(1) : '0';
  const avgHrs = N ? Math.round(all.reduce((a, r) => a + num(r.hrs), 0) / N) : 0;

  $('#kpis').innerHTML = [
    ['var(--brand)', N, 'ผู้ได้รับการเสนอทั้งหมด', `${new Set(all.map(r=>r.div)).size} ฝ่าย`],
    ['var(--blue)', interviewed, 'สัมภาษณ์แล้ว', `${pct(interviewed, N)}% ของทั้งหมด`],
    ['var(--green)', passed, 'ผ่านการสัมภาษณ์', decided ? `อัตราผ่าน ${pct(passed, decided)}% ของที่ทราบผล` : 'ยังไม่มีผล'],
    ['var(--red)', failed, 'ไม่ผ่าน', ''],
    ['var(--amber)', waiting + scheduled + notYet, 'คงเหลือต้องดำเนินการ', `นัดแล้ว ${scheduled} · ยังไม่นัด ${notYet}`],
    ['var(--accent)', both, 'ผ่านเกณฑ์ปี+ชม.ครบ', `${pct(both, N)}% ของทั้งหมด`],
    ['var(--brand-2)', avgTen + ' ปี', 'อายุงานเฉลี่ย', `ชม.อบรมเฉลี่ย ${avgHrs} ชม.`]
  ].map(k => `<div class="kpi" style="--c:${k[0]}"><div class="n">${k[1]}</div><div class="l">${k[2]}</div><div class="s">${k[3]}&nbsp;</div></div>`).join('');

  // donut
  let acc = 0; const stops = C.statuses.map(s => { const a = acc; acc += pct(cnt(s.key), N) * 3.6 || 0; return `${SCOL[s.color]} ${a}deg ${(cnt(s.key) / (N || 1)) * 360 + a}deg`; });
  let a2 = 0; const stops2 = C.statuses.map(s => { const a = a2, b = a + (cnt(s.key) / (N || 1)) * 360; a2 = b; return `${SCOL[s.color]} ${a}deg ${b}deg`; });
  $('#cDonut').innerHTML = `<div class="donut" data-c="${pct(decided + waiting, N)}%\nสัมภาษณ์แล้ว" style="background:conic-gradient(${N ? stops2.join(',') : '#eef1f5 0 360deg'})"></div>
    <div class="legend">${C.statuses.map(s => `<div><i style="background:${SCOL[s.color]}"></i>${esc(s.key)} <b>${cnt(s.key)}</b></div>`).join('')}</div>`;

  // criteria
  const cY = all.filter(r => okY(r) && !okH(r)).length, cH = all.filter(r => !okY(r) && okH(r)).length, cN = N - both - cY - cH;
  $('#cCrit').innerHTML = [['ผ่านครบทั้ง 2 เกณฑ์', both, 's-green'], ['ผ่านเฉพาะเกณฑ์อายุงาน', cY, 's-brand'], ['ผ่านเฉพาะเกณฑ์ชั่วโมงอบรม', cH, 's-amber'], ['ยังไม่ผ่านทั้ง 2 เกณฑ์', cN, 's-red']]
    .map(c => barRow(c[0], [{ n: c[1], cls: c[2], t: c[0] }], N).replace(/<b>\d+<\/b>/, `<b>${c[1]}</b>`)).join('')
    + `<p class="none">เกณฑ์อายุงานผ่าน ${all.filter(okY).length}/${N} · เกณฑ์ชั่วโมงอบรมผ่าน ${all.filter(okH).length}/${N}</p>`;

  // by division
  const divs = [...new Set(all.map(r => r.div))].map(d => ({ d, rs: all.filter(r => r.div === d) })).sort((a, b) => b.rs.length - a.rs.length);
  $('#cDiv').innerHTML = divs.map(x => barRow(x.d, C.statuses.map(s => ({ n: x.rs.filter(r => r.status === s.key).length, cls: 's-' + s.color, t: s.key })), x.rs.length)).join('')
    + `<div class="legend" style="grid-auto-flow:column;justify-content:start;gap:14px;margin-top:8px">${C.statuses.map(s => `<span><i style="background:${SCOL[s.color]}"></i>${esc(s.key)}</span>`).join('')}</div>`;

  // upcoming interviews
  const td = today(), up = all.filter(r => r.status === 'นัดสัมภาษณ์แล้ว' && r.intDate).sort((a, b) => a.intDate.localeCompare(b.intDate));
  const next = up.filter(r => r.intDate >= td).slice(0, 6), overdue = up.filter(r => r.intDate < td);
  $('#cUpcoming').innerHTML = next.length ? `<div class="list">${next.map(r => `<div class="it"><div>${esc(r.prefix + r.first + ' ' + r.last)}<small>${esc(r.div)} → ${esc(r.newPos)}</small></div><span class="dt">${fmtDate(r.intDate)}</span></div>`).join('')}</div>` : '<p class="none">ไม่มีนัดสัมภาษณ์ที่กำลังจะถึง</p>';

  // new positions
  const pm = {}; all.forEach(r => { const k = r.newPos || '-'; pm[k] = pm[k] || { n: 0, p: 0 }; pm[k].n++; if (r.status === 'ผ่าน') pm[k].p++; });
  $('#cPos').innerHTML = Object.entries(pm).sort((a, b) => b[1].n - a[1].n).slice(0, 7)
    .map(([k, v]) => barRow(k, [{ n: v.p, cls: 's-green', t: 'ผ่านแล้ว' }, { n: v.n - v.p, cls: 's-brand', t: 'อยู่ระหว่างดำเนินการ' }], v.n)).join('') || '<p class="none">-</p>';

  // tenure
  const bk = [['< 3 ปี', r => years(r) < 3], ['3–5 ปี', r => years(r) >= 3 && years(r) < 5], ['5–10 ปี', r => years(r) >= 5 && years(r) < 10], ['≥ 10 ปี', r => years(r) >= 10]];
  $('#cTenure').innerHTML = bk.map(([l, f]) => { const n = all.filter(f).length; return barRow(l, [{ n, cls: 's-brand', t: l }], N).replace(/<b>\d+<\/b>/, `<b>${n}</b>`); }).join('');

  // attention
  const lst = (arr, f) => arr.length ? `<ul>${arr.slice(0, 6).map(f).map(x => `<li>${x}</li>`).join('')}${arr.length > 6 ? `<li class="more">… และอีก ${arr.length - 6} ราย</li>` : ''}</ul>` : '<p class="none">ไม่มี</p>';
  const nm = r => `${esc(r.prefix + r.first + ' ' + r.last)} <span class="sub">(${esc(r.div)})</span>`;
  const boxes = [
    ['var(--red)', `นัดสัมภาษณ์เกินกำหนดแต่ยังไม่อัปเดตผล (${overdue.length})`, lst(overdue, r => `${nm(r)} · ${fmtDate(r.intDate)}`)],
    ['var(--amber)', `สัมภาษณ์แล้ว รอประกาศผล (${waiting})`, lst(all.filter(r => r.status === 'สัมภาษณ์แล้ว (รอผล)'), nm)],
    ['var(--blue)', `ผ่านเกณฑ์ครบ แต่ยังไม่มีนัดสัมภาษณ์ (${all.filter(r => r.status === C.statuses[0].key && bothOk(r)).length})`, lst(all.filter(r => r.status === C.statuses[0].key && bothOk(r)), nm)],
    ['var(--gray)', `ยังไม่ผ่านเกณฑ์คุณสมบัติบางข้อ (${N - both})`, lst(all.filter(r => !bothOk(r)), r => `${nm(r)} · ${!okY(r) ? 'อายุงาน ' : ''}${!okH(r) ? 'ชม.อบรม' : ''}`)]
  ];
  $('#cAttn').innerHTML = boxes.map(b => `<div class="attn-box" style="--c:${b[0]}"><h4>${b[1]}</h4>${b[2]}</div>`).join('');

  renderTable();
  $('#foot').textContent = `NSL FOODS PCL · Promotion Tracker · ข้อมูล ณ วันที่แสดงผล ${new Date().toLocaleDateString('th-TH', { day: 'numeric', month: 'long', year: 'numeric' })}`;
}

function renderTable() {
  const q = $('#q').value.trim().toLowerCase(), fd = $('#fDiv').value, fs = $('#fStatus').value;
  const list = rows.filter(r => (!q || [r.id, r.prefix, r.first, r.last, r.div, r.dept, r.branch, r.newPos].join(' ').toLowerCase().includes(q)) && (!fd || r.div === fd) && (!fs || r.status === fs));
  $('#count').textContent = `แสดง ${list.length} / ${rows.length} รายการ`;
  $('#tbody').innerHTML = list.length ? list.map(r => `<tr>
    <td class="c">${esc(r.no)}</td><td>${esc(r.id)}</td><td class="name"><b>${esc(r.prefix + r.first + ' ' + r.last)}</b></td>
    <td><div>${esc(r.div)}</div><div class="sub">${esc(r.dept)} · ${esc(r.branch)}</div></td>
    <td class="c">${esc(r.ty)}ปี ${esc(r.tm)}ด ${esc(r.td)}ว</td>
    <td class="c"><span class="chip ${crit(r.sumY)}">${esc(r.sumY) || '-'}</span></td>
    <td class="c">${esc(r.hrs)}</td><td class="c"><span class="chip ${crit(r.sumH)}">${esc(r.sumH) || '-'}</span></td>
    <td>${esc(r.oldPos)} <span class="sub">${esc(r.oldLv)}</span> → <b style="font-weight:500">${esc(r.newPos)}</b> <span class="sub">${esc(r.newLv)}</span></td>
    <td class="c"><span class="chip ${col(r.status)}">${esc(r.status)}</span></td>
    <td class="c">${fmtDate(r.intDate) || '-'}</td><td class="c">${fmtDate(r.statusDate) || '-'}</td><td>${esc(r.note)}</td></tr>`).join('')
    : '<tr><td colspan="13" class="empty">ไม่พบข้อมูล</td></tr>';
}
['#q', '#fDiv', '#fStatus'].forEach(s => $(s).addEventListener('input', renderTable));
})();
