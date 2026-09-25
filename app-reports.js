/* ==========================================================================
   app-reports.js — เลือกเนื้อหารายงาน/หนังสือ แล้วส่งให้ app-docs.js จัดหน้า (PDF + Word .docx) · S-Curve · ส่งออก Excel
   รูปแบบบันทึกข้อความเลียนแบบรายงานจริงของ รย.23/2569
   ========================================================================== */
window.R = (function () {
  'use strict';
  const esc = U.esc;
  function A() { return window.App; }

  /* ---------------- S-Curve (SVG) ---------------- */
  function sCurveSvg(p, units, o) {
    o = o || {};
    const c = U.contract(p);
    if (!c.start || !c.days) return '<div class="empty">กรอกวันเริ่มและระยะเวลาสัญญาเพื่อแสดงกราฟ</div>';
    const W = o.w || 640, H = o.h || 280, ml = 38, mr = 12, mt = 12, mb = 34;
    const x0 = c.start;
    let x1 = c.endExt;
    if (p.completedDate && p.completedDate > x1) x1 = p.completedDate;
    const span = Math.max(1, U.diffDays(x0, x1));
    const X = function (d) { return ml + (W - ml - mr) * U.diffDays(x0, d) / span; };
    const Y = function (v) { return mt + (H - mt - mb) * (1 - v / 100); };
    const step = Math.max(1, Math.round(span / 120));
    let planPts = [], actPts = [];
    const until = o.until && o.until < x1 ? o.until : x1;
    for (let i = 0; i <= span; i += step) {
      const d = U.addDays(x0, i);
      planPts.push(X(d).toFixed(1) + ',' + Y(U.planCum(p, d)).toFixed(1));
      if (d <= until) actPts.push(X(d).toFixed(1) + ',' + Y(U.actualCum(p, units, d)).toFixed(1));
    }
    planPts.push(X(x1).toFixed(1) + ',' + Y(U.planCum(p, x1)).toFixed(1));
    if (until >= x0) actPts.push(X(until).toFixed(1) + ',' + Y(U.actualCum(p, units, until)).toFixed(1));
    let g = '';
    [0, 25, 50, 75, 100].forEach(function (v) {
      g += '<line x1="' + ml + '" x2="' + (W - mr) + '" y1="' + Y(v) + '" y2="' + Y(v) + '" stroke="#dde7f4" stroke-width="1"/>' +
        '<text x="' + (ml - 6) + '" y="' + (Y(v) + 4) + '" text-anchor="end" font-size="11" fill="#6f84a3">' + v + '%</text>';
    });
    let d = U.parse(x0); d = new Date(d.getFullYear(), d.getMonth(), 1);
    const monthsCount = Math.round(span / 30);
    while (U.iso(d) <= x1) {
      const s = U.iso(d);
      if (s >= x0) {
        g += '<line x1="' + X(s) + '" x2="' + X(s) + '" y1="' + mt + '" y2="' + (H - mb) + '" stroke="#edf2fa"/>';
      }
      const mid = s < x0 ? x0 : s;
      if (monthsCount <= 14 || d.getMonth() % 3 === 0)
        g += '<text x="' + (X(mid) + 3) + '" y="' + (H - mb + 15) + '" font-size="11" fill="#6f84a3">' + U.thMonthShort(s) + '</text>';
      d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }
    const marks = [[c.q1, '1/4'], [c.q2, '2/4'], [c.end, 'สิ้นสุด']];
    if (c.ext) marks.push([c.endExt, 'ขยาย']);
    marks.forEach(function (m) {
      g += '<line x1="' + X(m[0]) + '" x2="' + X(m[0]) + '" y1="' + mt + '" y2="' + (H - mb) + '" stroke="#1a56b0" stroke-width="1" stroke-dasharray="3 3" opacity=".7"/>' +
        '<text x="' + (X(m[0]) - 3) + '" y="' + (mt + 10) + '" text-anchor="end" font-size="10.5" fill="#1a56b0">' + m[1] + '</text>';
    });
    const endPlan = U.planCum(p, until), endAct = U.actualCum(p, units, until);
    return '<svg viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg" font-family="Sarabun, sans-serif" role="img" aria-label="กราฟแผนและผลงานสะสม">' + g +
      '<polyline points="' + planPts.join(' ') + '" fill="none" stroke="#9db5d8" stroke-width="2" stroke-dasharray="6 4"/>' +
      (actPts.length ? '<polyline points="' + actPts.join(' ') + '" fill="none" stroke="#e0620f" stroke-width="2.6" stroke-linejoin="round"/>' +
        '<circle cx="' + X(until) + '" cy="' + Y(endAct) + '" r="4" fill="#e0620f"/>' +
        '<text x="' + Math.min(X(until) + 6, W - 60) + '" y="' + (Y(endAct) + (endAct > endPlan ? -8 : 14)) + '" font-size="12" font-weight="700" fill="#c2540c">' + endAct.toFixed(2) + '%</text>' : '') +
      '</svg>';
  }

  /* ---------------- ข้อความตั้งต้นของหนังสือแต่ละชนิด (แก้ไขได้ในหน้าต่างก่อนสร้าง) ----------------
     ถ้อยคำคัดจากหนังสือจริงของ รย.23/2569 */
  function pct3(v) { return (isFinite(v) ? v : 0).toFixed(3) + ' %'; }
  function defaults(type, o) {
    const app = A(), p = app.S.p, pr = U.periods(p);
    const copies = DOCS.ccList(p).length + 1;
    if (type === '3day') {
      const t = pr.threeDay;
      const worked = app.active('daily').filter(function (d) { return d.date >= t.from && d.date <= U.addDays(t.to, 1) && d.works && d.works.length; })
        .sort(function (a, b) { return a.date.localeCompare(b.date); })[0];
      return {
        subject: worked ? 'รายงาน 3 วันทำการ และเข้าดำเนินการ' : 'รายงาน 3 วันทำการ',
        body: 'โครงการฯ ขอรายงานผลการปฏิบัติงานของผู้รับจ้าง ภายใน 3 วันทำการ ขณะนี้ ' +
          (worked ? 'ผู้รับจ้างได้เข้าดำเนินการ' + (worked.works[0].desc || '') + 'แล้วตั้งแต่วันที่ ' + U.thDate(worked.date) : 'ผู้รับจ้างยังไม่ได้เข้าดำเนินการ'),
        closing: 'จึงเรียนมาเพื่อโปรดทราบ', copies: copies, attach: { daily: true, photos: true }
      };
    }
    if (type === 'weekly') {
      const w = pr.weeks.find(function (x) { return x.key === o.period; });
      return {
        subject: 'รายงานสรุปผลการปฏิบัติงานประจำสัปดาห์ที่ ' + w.n,
        body: 'โครงการฯ ขอรายงานผลการปฏิบัติงานของผู้รับจ้าง ระหว่างวันที่ ' + U.thDate(w.from) + ' – ' + U.thDate(w.to) + ' (ตามเอกสารแนบ)',
        closing: 'จึงเรียนมาเพื่อโปรดทราบ', copies: copies, attach: { summary: true, daily: true, photos: true }
      };
    }
    if (type === 'monthly') {
      const m = pr.months.find(function (x) { return x.key === o.period; });
      return {
        subject: 'รายงานสรุปผลการปฏิบัติงานประจำเดือน ' + U.thMonth(m.ym),
        body: 'โครงการฯ ขอสรุปผลการปฏิบัติงานประจำเดือน ' + U.thMonth(m.ym) + ' ตามเอกสารแนบ 1 ฉบับ',
        closing: 'จึงเรียนมาเพื่อโปรดทราบ', copies: copies, attach: { plan: true }
      };
    }
    return { subject: '', body: '', closing: 'จึงเรียนมาเพื่อโปรดทราบ', copies: 0, attach: {} };
  }
  function monthSumLines(p, m) {
    const units = A().active('units');
    const planM = U.num(((p.plan || []).find(function (x) { return x.month === m.ym; }) || {}).pct);
    const prev = U.addDays(m.from, -1);
    const a1 = U.actualCum(p, units, m.to), a0 = prev >= p.startDate ? U.actualCum(p, units, prev) : 0;
    return [
      { label: 'แผนงานเดือน ' + U.thMonth(m.ym), value: pct3(planM) },
      { label: 'แผนงานรวม', value: pct3(U.planCum(p, m.to)) },
      { label: 'ผลงานเดือน ' + U.thMonth(m.ym), value: pct3(a1 - a0) },
      { label: 'ผลงานรวม', value: pct3(a1) }
    ];
  }

  /* ---------------- สร้างรายงานแต่ละชนิด → { pages, subject, reportKey, title } ----------------
     o: { no, date, subject, body (ขึ้นบรรทัดใหม่ = ย่อหน้าใหม่), contract (ใส่ย่อหน้าอ้างสัญญา), closing, copies, attach:{summary,daily,photos,plan} } */
  function memoOpts(p, o) {
    const paras = [];
    if (o.contract !== false) paras.push(DOCS.contractPara(p));
    String(o.body || '').replace('{สัญญา}', '').split(/\n+/).map(function (s) { return s.replace(/^\t/, '').trim(); }).filter(Boolean).forEach(function (s) { paras.push(s); });
    return { no: o.no, date: o.date, subject: o.subject, to: o.to, paras: paras, closing: o.closing, copies: U.num(o.copies) };
  }
  function photosIn(from, to) {
    return A().active('photos').filter(function (x) { return x.date >= from && x.date <= to; })
      .sort(function (a, b) { return String(a.date + (a.time || '')).localeCompare(String(b.date + (b.time || ''))); });
  }
  async function build(type, o) {
    const app = A(), p = app.S.p, pr = U.periods(p);
    const at = o.attach || {};
    let pages = [], r = {};
    if (type === '3day') {
      const t = pr.threeDay;
      pages = DOCS.memoPages(p, memoOpts(p, o));
      if (at.daily) pages = pages.concat(DOCS.dailyFormPages(p, t.from, t.to));
      if (at.photos) pages = pages.concat(DOCS.photoPages(p, photosIn(t.from, t.to)));
      r = { reportKey: '3day' };
    } else if (type === 'weekly') {
      const w = pr.weeks.find(function (x) { return x.key === o.period; });
      if (!w) throw new Error('เลือกสัปดาห์');
      pages = DOCS.memoPages(p, memoOpts(p, o));
      if (at.summary) pages.push(DOCS.weeklySummaryPage(p, w));
      if (at.daily) pages = pages.concat(DOCS.dailyFormPages(p, w.from, w.to));
      if (at.photos) pages = pages.concat(DOCS.photoPages(p, photosIn(w.from, w.to)));
      r = { reportKey: w.key };
    } else if (type === 'monthly') {
      const m = pr.months.find(function (x) { return x.key === o.period; });
      if (!m) throw new Error('เลือกเดือน');
      const mo = memoOpts(p, o); mo.sumLines = monthSumLines(p, m);
      pages = DOCS.memoPages(p, mo);
      if (at.plan) pages.push(DOCS.monthlyPage(p, m));
      r = { reportKey: m.key };
    } else if (type === 'letter') {
      pages = DOCS.memoPages(p, memoOpts(p, o));
    } else if (type === 'daily') {
      pages = DOCS.dailyFormPages(p, o.date2, o.date2).concat(DOCS.photoPages(p, photosIn(o.date2, o.date2)));
      r = { subject: 'แบบบันทึกการปฏิบัติงานประจำวัน ' + U.thDate(o.date2) };
    } else if (type === 'photos') {
      if (!(o.photos || []).length) throw new Error('ไม่มีรูปตามเงื่อนไขที่เลือก');
      pages = DOCS.photoPages(p, o.photos, o.title);
      r = { subject: o.title || 'รูปถ่ายประกอบรายงาน' };
    } else if (type === 'delivery') {
      pages = [DOCS.deliveryPage(p, o)];
      r = { subject: 'บัญชีแสดงปริมาณงานและเงินค่างานที่ขอส่ง' };
    } else if (type === 'tests') {
      pages = [DOCS.testsPage(p)];
      r = { subject: 'รายละเอียดการส่งตัวอย่าง' };
    } else throw new Error('ไม่รู้จักชนิดรายงาน');
    const subject = r.subject || o.subject || '';
    return { pages: pages, subject: subject, title: subject, reportKey: r.reportKey || '' };
  }

  /* ---------------- ตัวอย่างก่อนพิมพ์ + ดาวน์โหลด ---------------- */
  async function openPreview(r, meta) {
    meta = meta || {};
    const app = A();
    const html = await DOCS.toHtml(r.pages, r.title);
    const m = app.modal({
      title: 'ตัวอย่าง: ' + (r.title || '') + ' (' + r.pages.length + ' หน้า)', size: 'wide',
      body: '<iframe id="rpFrame" style="width:100%;height:70vh;border:1px solid var(--line);border-radius:8px;background:#dfe6f0"></iframe>' +
        '<p class="hint" style="margin:8px 0 0">พิมพ์เป็น PDF: เลือกเครื่องพิมพ์ "บันทึกเป็น PDF" · ขนาด A4 · ระยะขอบ "ค่าเริ่มต้น" · ปิด "ส่วนหัวและส่วนท้าย"</p>',
      foot: (meta.register ? '<button class="btn btn-outline left" id="rpReg">ลงทะเบียนหนังสือออก</button>' : '') +
        '<button class="btn btn-outline" id="rpWord">ดาวน์โหลด Word (.docx)</button><button class="btn btn-primary" id="rpPrint">พิมพ์ / บันทึกเป็น PDF</button>'
    });
    const fr = m.q('#rpFrame');
    fr.srcdoc = html;
    m.q('#rpPrint').onclick = function () { fr.contentWindow.focus(); fr.contentWindow.print(); };
    m.q('#rpWord').onclick = async function () {
      const done = app.busy(this, 'กำลังสร้างไฟล์...');
      try { U.download(U.safeName(meta.fileName || r.title || 'เอกสาร') + '.docx', await DOCS.toDocx(r.pages)); }
      catch (e) { app.toast(e.message, true); } finally { done(); }
    };
    if (meta.register) m.q('#rpReg').onclick = async function () {
      const rec = { type: 'out', no: meta.no || '', date: meta.date || U.today(), subject: r.subject || r.title, party: app.S.p.addressee || '', reportKey: r.reportKey || '', note: '' };
      const dup = r.reportKey && app.active('docs').find(function (d) { return d.reportKey === r.reportKey; });
      if (dup && !await app.confirm('รายงานรอบนี้ลงทะเบียนไว้แล้ว (เลขที่ ' + dup.no + ') ต้องการลงซ้ำ?', 'ลงทะเบียนซ้ำ')) return;
      try { await app.save('docs', rec, 'หนังสือออก ' + rec.no); app.toast('ลงทะเบียนหนังสือ ' + rec.no + ' แล้ว'); } catch (e) { app.toast(e.message, true); }
    };
    return m;
  }

  /* ---------------- ส่งออก Excel ทั้งโครงการ ---------------- */
  function exportExcel() {
    const app = A(), p = app.S.p;
    if (typeof XLSX === 'undefined') return app.toast('โหลดตัวสร้าง Excel ไม่สำเร็จ (ตรวจอินเทอร์เน็ต)', true);
    const c = U.contract(p), units = app.active('units');
    const wb = XLSX.utils.book_new();
    const add = function (name, rows) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name); };
    add('สัญญา', [['เลขที่สัญญา', p.code], ['ชื่องาน', p.name || p.shortName], ['ผู้รับจ้าง', p.contractor], ['ลงวันที่', U.thDate(p.contractDate)],
      ['เริ่มสัญญา', U.thDate(c.start)], ['สิ้นสุดสัญญา', U.thDate(c.end)], ['ระยะเวลา (วัน)', c.days], ['ขยายเวลา (วัน)', c.ext], ['สิ้นสุดหลังขยาย', U.thDate(c.endExt)],
      ['ค่างาน', c.value], ['ค่าปรับ/วัน', c.finePerDay], ['ส่งงาน', U.thDate(p.completedDate)], ['ตรวจรับ', U.thDate(p.acceptedDate)],
      ['แผนสะสมวันนี้ %', +U.planCum(p, U.today()).toFixed(3)], ['ผลงานสะสมวันนี้ %', +U.actualCum(p, units, U.today()).toFixed(3)]]);
    add('BOQ', [['ที่', 'รายการ', 'หน่วย', 'ปริมาณ', 'ราคา/หน่วย', 'รวมเงิน', 'คิดผลงานจาก', '% แล้วเสร็จวันนี้']].concat((p.boq || []).map(function (b, i) {
      return [i + 1, b.desc, b.unit, U.num(b.qty), U.num(b.unitPrice), U.num(b.qty) * U.num(b.unitPrice), b.track === 'units' ? 'ทะเบียนชิ้นงาน' : 'กรอกเอง', +(U.itemFraction(p, b, units, U.today()) * 100).toFixed(2)];
    })));
    add('แผนงาน', [['เดือน', 'แผนประจำเดือน %', 'แผนสะสม %']].concat((p.plan || []).map(function (x) {
      return [U.thMonth(x.month), U.num(x.pct), +U.planCum(p, U.iso(new Date(+x.month.slice(0, 4), +x.month.slice(5, 7), 0))).toFixed(3)];
    })));
    const machines = p.machines || [];
    const dRows = [['วันที่', 'วัน', 'รายละเอียดงาน', 'จาก Sta.', 'ถึง Sta.', 'Lt./Rt.', 'ผลงาน', 'เวลา', 'สภาพอากาศ'].concat(machines, ['โฟร์แมน', 'คนงาน', 'ปัญหาอุปสรรค', 'ข้อสั่งการ', 'หมายเหตุ'])];
    app.active('daily').sort(function (a, b) { return a.date.localeCompare(b.date); }).forEach(function (d) {
      const ws = d.works && d.works.length ? d.works : [{}];
      ws.forEach(function (w, i) {
        dRows.push([U.thDate(d.date, 'num'), U.thDow(d.date), w.desc || (d.noWork ? 'ไม่มีการทำงาน' : ''), w.staFrom || '', w.staTo || '', w.side || '', w.qty || '',
          i ? '' : (d.timeFrom || '') + (d.timeTo ? '-' + d.timeTo : ''), i ? '' : d.weather || ''].concat(machines.map(function (m) { return i ? '' : ((d.machines || {})[m] || ''); }),
          [i ? '' : d.foreman || '', i ? '' : d.workers || '', i ? '' : d.problems || '', i ? '' : d.orders || '', i ? '' : d.note || '']));
      });
    });
    add('บันทึกประจำวัน', dRows);
    const st = p.stages || [];
    add('ชิ้นงาน', [['รหัส', 'จุด', 'กม.', 'ฝั่ง'].concat(st.map(function (s) { return s.name; }), ['หมายเหตุ'])].concat(units.sort(function (a, b) { return String(a.code).localeCompare(String(b.code), 'th', { numeric: true }); }).map(function (u) {
      return [u.code, u.group, u.km, u.side].concat(st.map(function (s) { return u.stages && u.stages[s.key] ? U.thDate(u.stages[s.key], 'num') : ''; }), [u.note || '']);
    })));
    add('ผลทดสอบ', [['ประเภท', 'รายการ', 'ตำแหน่ง', 'เกณฑ์', 'ค่าที่ได้', 'หน่วย', 'เก็บ/เท', 'บริษัทส่ง', 'แขวงส่ง', 'หน่วยทดสอบ', 'ผลออก', 'เลขที่ผล', 'สถานะ', 'หมายเหตุ']].concat(app.active('tests').map(function (t) {
      return [t.type, t.item, t.location, t.spec, t.value, t.unit, U.thDate(t.sampleDate, 'num'), U.thDate(t.companyDate, 'num'), U.thDate(t.sentDate, 'num'), t.lab, U.thDate(t.resultDate, 'num'), t.resultNo, t.status, t.note];
    })));
    add('เอกสาร', [['ประเภท', 'เลขที่', 'วันที่', 'เรื่อง', 'จาก/ถึง', 'ฉบับแก้ไข', 'สถานะแบบ', 'ไฟล์/ที่เก็บ', 'หมายเหตุ']].concat(app.active('docs').sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); }).map(function (d) {
      return [d.type, d.no, U.thDate(d.date, 'num'), d.subject, d.party, d.rev, d.drawingStatus, d.fileName || d.link || '', d.note];
    })));
    add('รูปถ่าย', [['วันที่', 'เวลา', 'หมวด', 'กม.', 'ฝั่ง', 'ชิ้นงาน', 'คำอธิบาย', 'กล้อง', 'พิกัด', 'ข้อสังเกต', 'ไฟล์ต้นฉบับ']].concat(app.active('photos').sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); }).map(function (x) {
      return [U.thDate(x.date, 'num'), x.time || '', x.category, x.km, x.side, x.unitCode || '', x.desc, [x.exif && x.exif.make, x.exif && x.exif.model].filter(Boolean).join(' '),
        x.exif && x.exif.lat ? x.exif.lat + ',' + x.exif.lng : '', x.flag || (x.noExif ? 'ไม่มีข้อมูลกล้องในไฟล์' : ''), x.originalName || ''];
    })));
    add('ตรวจความปลอดภัย', [['วันที่', 'จุด', 'ผู้ตรวจ', 'ผล', 'ข้อบกพร่อง', 'การแก้ไข', 'แก้ไขเสร็จ']].concat(app.active('safety').map(function (s) {
      return [U.thDate(s.date, 'num'), s.location, s.inspector, s.result === 'fail' ? 'พบข้อบกพร่อง' : 'เรียบร้อย', s.issues, s.fixNote, U.thDate(s.fixedDate, 'num')];
    })));
    XLSX.writeFile(wb, U.safeName('ข้อมูลโครงการ_' + p.code + '_' + U.today()) + '.xlsx');
  }

  return { sCurveSvg: sCurveSvg, build: build, defaults: defaults, openPreview: openPreview, exportExcel: exportExcel };
})();
