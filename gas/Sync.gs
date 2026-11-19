// ============================================================
// KPI กฟต.2 — Sync ข้อมูลจาก "แบบรายงานผล KPI กฟต.2 ประจำปี 2569"
// มาเป็น Landing Page / KPISummary / Results_K01–K17 ให้ Dashboard อ่าน
// (ไฟล์นี้แยกจากโค้ดระบบบันทึกเพิ่มเติม — อย่าลบไฟล์ใดไฟล์หนึ่ง)
// ทริกเกอร์เดิมเรียก syncAll และ syncGFJ_K01..K17 ทุก 6 ชม.
// ============================================================

const KS_SRC_ID = '1wNdue0PW9sx0VE2fSbAnRaxacarc6tQnFBQ-RpkI2MM';
const KS_DST_ID = '1VmYqWuoTuX_xo7mB2n6LzW6ZH1l_xE0CtzhvzASIkdI';
const KS_AOL_CODE = 'K04201'; // กฟส.อ่าวลึก

const KS_KPI_SHEETS = [
  '1.1','1.2.1','1.2.2','1.2.3','1.3.1','1.3.2','1.4.1','1.4.2',
  '1.5','1.6','1.7','1.8','1.9','1.10','1.11','1.12','1.13','1.14',
  '1.15','1.16','1.17','1.18','1.19.1','1.19.2','1.19.3','1.20.1',
  '1.20.2','1.20.3','1.21','1.22.1-1','1.22.1-2','1.22.2','1.22.3',
  '1.23.1-1','1.23.1-2','1.23.2','1.23.3','1.23.4','1.23.5-1',
  '1.23.5-2','1.23.5-3','1.24.1','1.25','1.26','1.27','1.28','1.29',
  '1.30','1.31','1.32','2.1','2.2','2.3','2.4.1','2.4.2','2.5.1',
  '2.5.2','2.5.3','2.5.4','2.5.5','2.5.6','2.5.7','2.6','2.7','2.8',
  '2.9','2.10.1','2.10.2','2.11','2.12.1','2.12.2','2.12.3','2.13.1',
  '2.13.2','2.14.1','2.14.2','2.15','2.16','2.17','2.18','2.19',
  '2.20','2.21','2.22','2.23','2.24','2.25','2.26','2.27','2.28.1',
  '2.28.2','2.29.1','2.29.2','2.30','2.31'
];

const KS_MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
const KS_MONTH_KEY = ['ม.ค','ก.พ','มี.ค','เม.ย','พ.ค','มิ.ย','ก.ค','ส.ค','ก.ย','ต.ค','พ.ย','ธ.ค'];

// รหัสหน่วยงานในแต่ละ Results_Kxx (ชุดเดียวกับที่ Dashboard ใช้อยู่)
const KS_GFJ = {
  K01: ['K01000','K01101','K01102','K01103','K01104','K01105','K01401','K01402','K01601'],
  K02: ['K02000','K02101','K02102','K02401'],
  K03: ['K03000','K03101','K03102','K03103','K03201','K03203','K03204','K03401','K03403','K03404','K03501'],
  K04: ['K04000','K04101','K04107','K04108','K04201','K04202','K04301','K04302','K04401','K04501','K04601'],
  K05: ['K05000','K05101','K05103','K05301','K05401'],
  K06: ['K06000','K06101','K06401'],
  K07: ['K07000','K07101','K07102','K07104','K07201','K07202','K07301','K07401','K07402'],
  K08: ['K08000','K08101','K08103','K08201'],
  K09: ['K09000','K09101','K09105','K09106','K09201','K09301','K09303','K09401','K09501','K09502'],
  K10: ['K10000','K10101','K10102','K10201'],
  K11: ['K11000','K11101','K11102'],
  K12: ['K12000','K12101','K12102','K12103','K12104'],
  K13: ['K13000','K13101'],
  K14: ['K14000','K14101','K14102'],
  K15: ['K15000','K15101','K15102','K15201'],
  K16: ['K16000','K16101','K16102','K16201','K16301'],
  K17: ['K17000','K17101','K17102','K17103']
};

// ============================================================
// จุดเรียกจากทริกเกอร์ (ชื่อต้องตรงกับทริกเกอร์เดิม)
// ============================================================
function syncAll() {
  const t0 = new Date();
  const src = SpreadsheetApp.openById(KS_SRC_ID);
  const dst = SpreadsheetApp.openById(KS_DST_ID);
  ksSyncLanding_(src, dst);
  ksSyncSummary_(src, dst);
  ksLog_(dst, 'Sync: Landing+KPISummary', t0);
}
function syncGFJ_K01(){ ksSyncGFJ_('K01'); }
function syncGFJ_K02(){ ksSyncGFJ_('K02'); }
function syncGFJ_K03(){ ksSyncGFJ_('K03'); }
function syncGFJ_K04(){ ksSyncGFJ_('K04'); }
function syncGFJ_K05(){ ksSyncGFJ_('K05'); }
function syncGFJ_K06(){ ksSyncGFJ_('K06'); }
function syncGFJ_K07(){ ksSyncGFJ_('K07'); }
function syncGFJ_K08(){ ksSyncGFJ_('K08'); }
function syncGFJ_K09(){ ksSyncGFJ_('K09'); }
function syncGFJ_K10(){ ksSyncGFJ_('K10'); }
function syncGFJ_K11(){ ksSyncGFJ_('K11'); }
function syncGFJ_K12(){ ksSyncGFJ_('K12'); }
function syncGFJ_K13(){ ksSyncGFJ_('K13'); }
function syncGFJ_K14(){ ksSyncGFJ_('K14'); }
function syncGFJ_K15(){ ksSyncGFJ_('K15'); }
function syncGFJ_K16(){ ksSyncGFJ_('K16'); }
function syncGFJ_K17(){ ksSyncGFJ_('K17'); }

// ============================================================
// อ่านชีต KPI 1 ชีต → โครงสร้างกลาง
// { name, unit, owner, scope, rows: {code6: {name, t[5], res[12], sc[12], last, lastSc}} }
// ============================================================
function ksReadKpi_(sheet) {
  const v = sheet.getDataRange().getValues();
  const cell = (r, c) => (v[r] && v[r][c] !== undefined && v[r][c] !== null) ? v[r][c] : '';
  const info = { name: String(cell(1,1)), unit: String(cell(2,1)), owner: String(cell(3,1)), scope: String(cell(4,1)), rows: {} };

  // หาแถวหัวตาราง (แถวที่มี "ม.ค." และ "ผลล่าสุด")
  let hdr = -1;
  for (let r = 0; r < Math.min(v.length, 30); r++) {
    const line = v[r].map(x => String(x)).join('|');
    if (line.indexOf('ผลล่าสุด') >= 0 && line.indexOf('ม.ค') >= 0) { hdr = r; break; }
  }
  let head = hdr >= 0 ? v[hdr].map(x => String(x).trim()) : [];
  // หัวคอลัมน์บางช่องอาจเป็นวันที่ → อ่านแบบข้อความที่แสดงผล
  if (hdr >= 0 && head.filter(h => KS_MONTH_KEY.some(m => h.indexOf(m) >= 0)).length < 12) {
    head = sheet.getRange(hdr + 1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(x => String(x).trim());
  }

  const resCol = Array(12).fill(-1), scCol = Array(12).fill(-1);
  let lastCol = -1, lastScCol = -1;
  head.forEach((h, j) => {
    if (!h || h.indexOf('ประมาณการ') >= 0) return;
    if (h === 'ผลล่าสุด') { lastCol = j; return; }
    if (h === 'คะแนนล่าสุด') { lastScCol = j; return; }
    for (let m = 0; m < 12; m++) {
      if (h.indexOf(KS_MONTH_KEY[m]) >= 0) {
        if (h.indexOf('คะแนน') >= 0) { if (scCol[m] < 0) scCol[m] = j; }
        else { if (resCol[m] < 0) resCol[m] = j; }
        break;
      }
    }
  });
  // ค่าตั้งต้นตามรูปแบบชีตปัจจุบัน ถ้าหาหัวคอลัมน์ไม่เจอ
  for (let m = 0; m < 12; m++) {
    if (resCol[m] < 0) resCol[m] = 8 + m;
    if (scCol[m] < 0) scCol[m] = 24 + m;
  }
  if (lastCol < 0) lastCol = 22;
  if (lastScCol < 0) lastScCol = 23;

  const clean = x => (x === '#N/A' || x === '#REF!' || x === '#DIV/0!' || x === '#VALUE!') ? '' : x;
  for (let r = (hdr >= 0 ? hdr + 1 : 10); r < v.length; r++) {
    const code6 = String(cell(r, 0)).trim();
    if (!/^K\d{5}$/.test(code6)) continue;
    info.rows[code6] = {
      name: String(cell(r, 2)),
      t:   [3,4,5,6,7].map(c => clean(cell(r, c))),
      res: resCol.map(c => clean(cell(r, c))),
      sc:  scCol.map(c => clean(cell(r, c))),
      last: clean(cell(r, lastCol)),
      lastSc: clean(cell(r, lastScCol))
    };
  }
  return info;
}

// ============================================================
// Landing Page — คัดลอกทั้งชีต
// ============================================================
function ksSyncLanding_(src, dst) {
  const s = src.getSheetByName('Landing Page');
  if (!s) return;
  const data = s.getDataRange().getValues();
  const d = ksSheet_(dst, 'Landing Page');
  d.clearContents();
  d.getRange(1, 1, data.length, 1).setNumberFormat('@');
  d.getRange(1, 1, data.length, data[0].length).setValues(data);
}

// ============================================================
// KPISummary — ภาพรวม กฟต.2 + กฟส.อ่าวลึก (65 คอลัมน์ ตามที่ Dashboard อ่าน)
// ============================================================
function ksSyncSummary_(src, dst) {
  const header = ['KPI Code','ชื่อ KPI','หน่วยวัด','ผู้รับผิดชอบ','หน่วยงานที่ประเมิน'];
  for (let i = 1; i <= 5; i++) header.push('เป้าหมาย Lv' + i);
  for (let i = 1; i <= 5; i++) header.push('เป้า อล. Lv' + i);
  KS_MONTHS.forEach(m => header.push('ผล กฟต.2 ' + m));
  KS_MONTHS.forEach(m => header.push('คะแนน กฟต.2 ' + m));
  KS_MONTHS.forEach(m => header.push('ผล อล. ' + m));
  header.push('ผลล่าสุด อล.', 'คะแนนล่าสุด อล.');
  KS_MONTHS.forEach(m => header.push('คะแนน อล. ' + m));

  const out = [];
  KS_KPI_SHEETS.forEach(sn => {
    const empty = Array(header.length - 2).fill('');
    try {
      const s = src.getSheetByName(sn);
      if (!s) { out.push([sn, '(ไม่พบชีต)'].concat(empty)); return; }
      const k = ksReadKpi_(s);
      const g = k.rows['K00000'] || { t: Array(5).fill(''), res: Array(12).fill(''), sc: Array(12).fill('') };
      const a = k.rows[KS_AOL_CODE] || { t: Array(5).fill(''), res: Array(12).fill(''), sc: Array(12).fill(''), last: '', lastSc: '' };
      out.push([sn, k.name, k.unit, k.owner, k.scope]
        .concat(g.t, a.t, g.res, g.sc, a.res, [a.last, a.lastSc], a.sc));
    } catch (e) {
      out.push([sn, 'Error: ' + e.message].concat(empty));
    }
  });

  const d = ksSheet_(dst, 'KPISummary');
  d.clearContents();
  d.getRange(1, 1, out.length + 1, 1).setNumberFormat('@'); // กัน 1.10 → 1.1
  d.getRange(1, 1, 1, header.length).setValues([header]);
  d.getRange(2, 1, out.length, header.length).setValues(out);
}

// ============================================================
// Results_Kxx — ทุกหน่วยงานในจุดรวมงาน (31 คอลัมน์/หน่วยงาน)
// ============================================================
function ksSyncGFJ_(key) {
  const codes = KS_GFJ[key];
  if (!codes) return;
  const t0 = new Date();
  const src = SpreadsheetApp.openById(KS_SRC_ID);
  const dst = SpreadsheetApp.openById(KS_DST_ID);

  const header = ['KPI Code', 'ชื่อ KPI'];
  codes.forEach(c => {
    for (let i = 1; i <= 5; i++) header.push(c + ' เป้า Lv' + i);
    header.push(c + ' ผลล่าสุด', c + ' คะแนนล่าสุด');
    KS_MONTHS.forEach(m => header.push(c + ' ผล ' + m));
    KS_MONTHS.forEach(m => header.push(c + ' คะแนน ' + m));
  });

  const out = [];
  KS_KPI_SHEETS.forEach(sn => {
    const empty = Array(header.length - 2).fill('');
    try {
      const s = src.getSheetByName(sn);
      if (!s) { out.push([sn, '(ไม่พบชีต)'].concat(empty)); return; }
      const k = ksReadKpi_(s);
      let row = [sn, k.name];
      codes.forEach(c => {
        const o = k.rows[c];
        if (!o) { row = row.concat(Array(31).fill('')); return; }
        row = row.concat(o.t, [o.last, o.lastSc], o.res, o.sc);
      });
      out.push(row);
    } catch (e) {
      out.push([sn, 'Error: ' + e.message].concat(empty));
    }
  });

  const d = ksSheet_(dst, 'Results_' + key);
  d.clearContents();
  d.getRange(1, 1, out.length + 1, 1).setNumberFormat('@');
  d.getRange(1, 1, 1, header.length).setValues([header]);
  d.getRange(2, 1, out.length, header.length).setValues(out);
  ksLog_(dst, 'Sync: Results_' + key, t0);
}

// ============================================================
// เครื่องมือ
// ============================================================
function ksSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function ksLog_(ss, label, t0) {
  const log = ksSheet_(ss, '_SyncLog');
  log.insertRowBefore(1);
  log.getRange(1, 1, 1, 3).setValues([[new Date(), label, Math.round((new Date() - t0) / 1000) + ' วิ']]);
  if (log.getMaxRows() > 3000) log.deleteRows(3001, log.getMaxRows() - 3000);
}

// ติดตั้งทริกเกอร์ใหม่ทั้งหมด (ใช้เฉพาะเมื่อทริกเกอร์หาย — ปกติไม่ต้องรัน)
function setupAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'syncAll' || f.indexOf('syncGFJ_') === 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(6).create();
  Object.keys(KS_GFJ).forEach(k => ScriptApp.newTrigger('syncGFJ_' + k).timeBased().everyHours(6).create());
}
