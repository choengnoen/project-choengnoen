/* ==========================================================================
   app-reports.js — สร้างรายงาน/หนังสือ (พิมพ์เป็น PDF หรือดาวน์โหลด Word) + S-Curve + ส่งออก Excel
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
      g += '<line x1="' + ml + '" x2="' + (W - mr) + '" y1="' + Y(v) + '" y2="' + Y(v) + '" stroke="#e3e9e6" stroke-width="1"/>' +
        '<text x="' + (ml - 6) + '" y="' + (Y(v) + 4) + '" text-anchor="end" font-size="11" fill="#84958f">' + v + '%</text>';
    });
    let d = U.parse(x0); d = new Date(d.getFullYear(), d.getMonth(), 1);
    const monthsCount = Math.round(span / 30);
    while (U.iso(d) <= x1) {
      const s = U.iso(d);
      if (s >= x0) {
        g += '<line x1="' + X(s) + '" x2="' + X(s) + '" y1="' + mt + '" y2="' + (H - mb) + '" stroke="#eef2f0"/>';
      }
      const mid = s < x0 ? x0 : s;
      if (monthsCount <= 14 || d.getMonth() % 3 === 0)
        g += '<text x="' + (X(mid) + 3) + '" y="' + (H - mb + 15) + '" font-size="11" fill="#84958f">' + U.thMonthShort(s) + '</text>';
      d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }
    const marks = [[c.q1, '1/4'], [c.q2, '2/4'], [c.end, 'สิ้นสุด']];
    if (c.ext) marks.push([c.endExt, 'ขยาย']);
    marks.forEach(function (m) {
      g += '<line x1="' + X(m[0]) + '" x2="' + X(m[0]) + '" y1="' + mt + '" y2="' + (H - mb) + '" stroke="#c1402f" stroke-width="1" stroke-dasharray="3 3" opacity=".7"/>' +
        '<text x="' + (X(m[0]) - 3) + '" y="' + (mt + 10) + '" text-anchor="end" font-size="10.5" fill="#c1402f">' + m[1] + '</text>';
    });
    const endPlan = U.planCum(p, until), endAct = U.actualCum(p, units, until);
    return '<svg viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg" font-family="Sarabun, sans-serif" role="img" aria-label="กราฟแผนและผลงานสะสม">' + g +
      '<polyline points="' + planPts.join(' ') + '" fill="none" stroke="#9aa8a3" stroke-width="2" stroke-dasharray="6 4"/>' +
      (actPts.length ? '<polyline points="' + actPts.join(' ') + '" fill="none" stroke="#0b8f63" stroke-width="2.6" stroke-linejoin="round"/>' +
        '<circle cx="' + X(until) + '" cy="' + Y(endAct) + '" r="4" fill="#0b8f63"/>' +
        '<text x="' + Math.min(X(until) + 6, W - 60) + '" y="' + (Y(endAct) + (endAct > endPlan ? -8 : 14)) + '" font-size="12" font-weight="700" fill="#0a6f4d">' + endAct.toFixed(2) + '%</text>' : '') +
      '</svg>';
  }

  /* ---------------- ส่วนประกอบหนังสือ ---------------- */
  function contractPara(p) {
    const c = U.contract(p);
    let s = 'ตามสัญญาจ้างเลขที่ ' + (p.contractNo || p.code || '') + (p.contractDate ? ' ลงวันที่ ' + U.thDate(p.contractDate) : '') + ' ' + (p.name || p.shortName || '') +
      (p.contractor ? ' โดย ' + p.contractor + ' เป็นผู้รับจ้าง' : '') +
      (c.start ? ' เริ่มสัญญาวันที่ ' + U.thDate(c.start) + ' สิ้นสุดสัญญาวันที่ ' + U.thDate(c.end) + ' เวลาทำการ ' + c.days + ' วัน' : '');
    if (c.ext) s += ' ขยายเวลาทำการ ' + c.ext + ' วัน สิ้นสุดสัญญาวันที่ ' + U.thDate(c.endExt);
    s += ' ค่างาน ' + U.money(c.value) + ' บาท ค่าปรับวันละ ' + U.money(c.finePerDay) + ' บาท นั้น';
    return s;
  }
  function ccList(p) {
    return (p.committee || []).filter(function (m) { return m.role !== 'ประธานกรรมการ' && m.position; })
      .map(function (m) { return '- ' + m.position + ' (กรรมการฯ)'; });
  }
  function signBlock(p, role) {
    const sv = p.supervisor || {};
    return '<table class="sign"><tr><td></td><td class="c"><br><br>(' + esc(sv.name || '.......................................') + ')<br>' +
      esc(role || 'ผู้ควบคุมงาน') + (sv.position && !role ? '<br>' + esc(sv.position) : '') + '</td></tr></table>';
  }
  // o: { no, date, subject, to, paras:[text], closing, cc:bool }
  function memo(p, o) {
    const paras = (o.paras || []).map(function (t) { return '<p class="ind">' + esc(t).replace(/\n/g, '<br>') + '</p>'; }).join('');
    const cc = o.cc === false ? [] : ccList(p);
    return '<div class="memo">' +
      '<div class="memo-title">บันทึกข้อความ</div>' +
      '<p><b>ส่วนราชการ</b>&nbsp;&nbsp;' + esc(p.office || '') + (p.phone ? '&nbsp;&nbsp;โทรศัพท์ ' + esc(p.phone) : '') + '</p>' +
      '<table class="ln"><tr><td><b>ที่</b>&nbsp;&nbsp;' + esc(o.no || '') + '</td><td><b>วันที่</b>&nbsp;&nbsp;' + esc(o.date ? U.thDate(o.date) : '') + '</td></tr></table>' +
      '<p><b>เรื่อง</b>&nbsp;&nbsp;' + esc(o.subject || '') + '</p>' +
      '<p>เรียน&nbsp;&nbsp;' + esc(o.to || p.addressee || '') + '</p>' +
      paras + '<p class="ind">' + esc(o.closing || 'จึงเรียนมาเพื่อโปรดทราบ') + '</p>' +
      signBlock(p) +
      (cc.length ? '<p class="cc">สำเนาเรียน ' + cc.map(esc).join('<br><span class="cc-ind"></span>') + '<br>เพื่อโปรดทราบ</p>' : '') +
      '</div>';
  }

  /* ---------------- เนื้อหาแนบ ---------------- */
  function dailyTable(p, from, to) {
    const daily = A().active('daily');
    const machines = p.machines || [];
    const byDate = {}; daily.forEach(function (d) { byDate[d.date] = d; });
    let rows = '';
    for (let d = from; d <= to; d = U.addDays(d, 1)) {
      const r = byDate[d];
      const works = r && r.works && r.works.length ? r.works : [];
      const cell = function (fn) { return works.map(fn).join('<br>'); };
      rows += '<tr><td class="c">' + U.thDow(d) + '</td><td class="c nowrap">' + U.thDate(d, 'num') + '</td>' +
        '<td>' + (r ? (works.length ? cell(function (w) { return esc(w.desc || ''); }) : esc(r.noWork ? (r.note || 'ไม่มีการทำงาน') : '')) : '') + '</td>' +
        '<td class="c">' + cell(function (w) { return esc(w.staFrom || ''); }) + '</td><td class="c">' + cell(function (w) { return esc(w.staTo || ''); }) + '</td>' +
        '<td class="c">' + cell(function (w) { return esc(w.side || ''); }) + '</td><td class="c">' + cell(function (w) { return esc(w.qty || ''); }) + '</td>' +
        '<td class="c nowrap">' + esc(r && works.length ? (r.timeFrom || '') + (r.timeTo ? ' - ' + r.timeTo : '') : '') + '</td>' +
        '<td class="c">' + esc(r ? r.weather || '' : '') + '</td>' +
        machines.map(function (mc) { const v = r && r.machines ? r.machines[mc] : ''; return '<td class="c">' + (v ? esc(v) : '') + '</td>'; }).join('') +
        '<td class="c">' + (r && U.num(r.foreman) ? esc(r.foreman) : '') + '</td><td class="c">' + (r && U.num(r.workers) ? esc(r.workers) : '') + '</td>' +
        '<td>' + esc(r ? [r.problems, r.orders, !r.noWork ? r.note : ''].filter(Boolean).join(' / ') : '') + '</td></tr>';
    }
    return '<table class="grid-t"><thead><tr><th rowspan="2">วัน</th><th rowspan="2">วันที่</th><th rowspan="2">รายละเอียดงาน</th><th colspan="3">ช่วงดำเนินการ</th><th rowspan="2">ผลงานที่ได้</th>' +
      '<th rowspan="2">เวลา</th><th rowspan="2">สภาพอากาศ</th>' + (machines.length ? '<th colspan="' + machines.length + '">เครื่องจักร/เครื่องมือ (คัน/เครื่อง)</th>' : '') +
      '<th colspan="2">บุคลากร</th><th rowspan="2">ปัญหา/หมายเหตุ</th></tr><tr><th>จาก Sta.</th><th>ถึง Sta.</th><th>Lt./Rt.</th>' +
      machines.map(function (m) { return '<th class="vert">' + esc(m) + '</th>'; }).join('') + '<th class="vert">โฟร์แมน</th><th class="vert">คนงาน</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }
  function progressLine(p, d) {
    const units = A().active('units');
    const pl = U.planCum(p, d), ac = U.actualCum(p, units, d), diff = ac - pl;
    return 'ผลงานตามแผนสะสม ' + pl.toFixed(2) + '% · ผลงานจริงสะสม ' + ac.toFixed(2) + '% · ' +
      (Math.abs(diff) < 0.005 ? 'เป็นไปตามแผน' : (diff > 0 ? 'เร็วกว่าแผน ' : 'ช้ากว่าแผน ') + Math.abs(diff).toFixed(2) + '%') + ' (ณ วันที่ ' + U.thDate(d) + ')';
  }
  function attachHead(p, title, sub) {
    return '<div class="at-head"><div class="at-t">' + esc(title) + '</div>' + (sub ? '<div>' + esc(sub) + '</div>' : '') +
      '<div class="small">สัญญาจ้างเลขที่ ' + esc(p.contractNo || p.code || '') + ' · ' + esc(p.shortName || '') + (p.contractor ? ' · ผู้รับจ้าง ' + esc(p.contractor) : '') + '</div></div>';
  }
  function boqProgressTable(p, prevDate, toDate) {
    const units = A().active('units');
    const total = U.boqTotal(p) || 1;
    let sumPrev = 0, sumNow = 0;
    const rows = (p.boq || []).map(function (b, i) {
      const val = U.num(b.qty) * U.num(b.unitPrice);
      const w = val / total * 100;
      const fp = prevDate ? U.itemFraction(p, b, units, prevDate) : 0, fn = U.itemFraction(p, b, units, toDate);
      sumPrev += w * fp; sumNow += w * fn;
      return '<tr><td class="c">' + (i + 1) + '</td><td>' + esc(b.desc) + '</td><td class="r">' + U.money(b.qty, 0) + '</td><td class="c">' + esc(b.unit || '') + '</td>' +
        '<td class="r">' + U.money(b.unitPrice) + '</td><td class="r">' + U.money(val) + '</td><td class="r">' + w.toFixed(3) + '</td>' +
        '<td class="r">' + (w * fp).toFixed(3) + '</td><td class="r">' + (w * (fn - fp)).toFixed(3) + '</td><td class="r">' + (w * fn).toFixed(3) + '</td>' +
        '<td class="r">' + U.money(val * fn) + '</td></tr>';
    }).join('');
    return '<table class="grid-t"><thead><tr><th>ที่</th><th>รายการ</th><th>ปริมาณ</th><th>หน่วย</th><th>ราคา/หน่วย</th><th>รวมเงิน (บาท)</th><th>% งาน</th>' +
      '<th>สะสมยกมา %</th><th>ผลงานงวดนี้ %</th><th>สะสม %</th><th>มูลค่าสะสม (บาท)</th></tr></thead><tbody>' + rows +
      '</tbody><tfoot><tr><td></td><td>รวมทั้งสิ้น</td><td></td><td></td><td></td><td class="r">' + U.money(total) + '</td><td class="r">100.000</td>' +
      '<td class="r">' + sumPrev.toFixed(3) + '</td><td class="r">' + (sumNow - sumPrev).toFixed(3) + '</td><td class="r">' + sumNow.toFixed(3) + '</td><td class="r">' + U.money(sumNow / 100 * total) + '</td></tr></tfoot></table>';
  }
  function monthPlanTable(p, uptoDate) {
    const units = A().active('units');
    const plan = (p.plan || []).filter(function (x) { return x.month; });
    if (!plan.length) return '';
    let prevAct = 0;
    const cols = plan.map(function (x) {
      const endM = U.iso(new Date(+x.month.slice(0, 4), +x.month.slice(5, 7), 0));
      const pc = U.planCum(p, endM);
      const ref = endM > uptoDate ? uptoDate : endM;
      const has = x.month + '-01' <= uptoDate;
      const ac = has ? U.actualCum(p, units, ref) : null;
      const r = { m: x.month, pm: U.num(x.pct), pc: pc, am: ac === null ? null : ac - prevAct, ac: ac };
      if (ac !== null) prevAct = ac;
      return r;
    });
    const row = function (label, fn) { return '<tr><td>' + label + '</td>' + cols.map(function (c) { const v = fn(c); return '<td class="r">' + (v === null ? '' : v.toFixed(3)) + '</td>'; }).join('') + '</tr>'; };
    return '<table class="grid-t"><thead><tr><th>ร้อยละ</th>' + cols.map(function (c) { return '<th>' + U.thMonthShort(c.m) + '</th>'; }).join('') + '</tr></thead><tbody>' +
      row('แผนงานประจำเดือน', function (c) { return c.pm; }) + row('แผนงานสะสม', function (c) { return c.pc; }) +
      row('ผลงานประจำเดือน', function (c) { return c.am; }) + row('ผลงานสะสม', function (c) { return c.ac; }) + '</tbody></table>';
  }

  /* ---------------- สร้างรายงานแต่ละชนิด ---------------- */
  // คืน { parts:[{html, land}], subject, reportKey, title }
  async function build(type, o) {
    const app = A(), p = app.S.p;
    const c = U.contract(p);
    const parts = [];
    if (type === '3day') {
      const pr = U.periods(p).threeDay;
      const worked = app.active('daily').some(function (d) { return d.date >= pr.from && d.date <= pr.to && d.works && d.works.length; });
      const subject = 'รายงาน 3 วันทำการ' + (worked ? ' และเข้าดำเนินการ' : ' (ผู้รับจ้างยังไม่เข้าดำเนินการ)');
      parts.push({ html: memo(p, { no: o.no, date: o.date, subject: subject, paras: [contractPara(p),
        'โครงการฯ ขอรายงานผลการปฏิบัติงานของผู้รับจ้าง 3 วันทำการ ระหว่างวันที่ ' + U.thRange(pr.from, pr.to, true) +
        (worked ? ' ผู้รับจ้างได้เข้าดำเนินการแล้ว' : ' ผู้รับจ้างยังไม่ได้เข้าดำเนินการ') + ' (ตามเอกสารแนบ)'] }) });
      parts.push({ land: true, html: attachHead(p, 'รายงานผลการปฏิบัติงาน 3 วันทำการ', 'ระหว่างวันที่ ' + U.thRange(pr.from, pr.to, true)) + dailyTable(p, pr.from, pr.to) + signBlock(p) });
      return { parts: parts, subject: subject, reportKey: '3day', title: subject };
    }
    if (type === 'weekly') {
      const w = U.periods(p).weeks.find(function (x) { return x.key === o.period; });
      if (!w) throw new Error('เลือกสัปดาห์');
      const subject = 'รายงานสรุปผลการปฏิบัติงานประจำสัปดาห์ที่ ' + w.n;
      parts.push({ html: memo(p, { no: o.no, date: o.date, subject: subject, paras: [contractPara(p),
        'โครงการฯ ขอรายงานผลการปฏิบัติงานของผู้รับจ้าง ระหว่างวันที่ ' + U.thRange(w.from, w.to, true) + ' (ตามเอกสารแนบ)'] }) });
      parts.push({ land: true, html: attachHead(p, 'รายงานผลการปฏิบัติงานประจำสัปดาห์ที่ ' + w.n, 'ระหว่างวันที่ ' + U.thRange(w.from, w.to, true)) +
        dailyTable(p, w.from, w.to) + '<p class="small"><b>สรุป:</b> ' + esc(progressLine(p, w.to)) + '</p>' + signBlock(p) });
      return { parts: parts, subject: subject + ' (' + U.thRange(w.from, w.to) + ')', reportKey: w.key, title: subject };
    }
    if (type === 'monthly') {
      const m = U.periods(p).months.find(function (x) { return x.key === o.period; });
      if (!m) throw new Error('เลือกเดือน');
      const units = app.active('units');
      const pl = U.planCum(p, m.to), ac = U.actualCum(p, units, m.to), diff = ac - pl;
      const subject = 'รายงานผลการปฏิบัติงานประจำเดือน ' + U.thMonth(m.ym);
      parts.push({ html: memo(p, { no: o.no, date: o.date, subject: subject, paras: [contractPara(p),
        'โครงการฯ ขอรายงานผลการปฏิบัติงานของผู้รับจ้างประจำเดือน ' + U.thMonth(m.ym) + ' ระหว่างวันที่ ' + U.thRange(m.from, m.to, true) +
        ' ผลงานตามแผนสะสมร้อยละ ' + pl.toFixed(2) + ' ผลงานที่ทำได้จริงสะสมร้อยละ ' + ac.toFixed(2) + ' ' +
        (Math.abs(diff) < 0.005 ? 'เป็นไปตามแผนงาน' : (diff > 0 ? 'เร็วกว่าแผนงาน' : 'ช้ากว่าแผนงาน') + ' ร้อยละ ' + Math.abs(diff).toFixed(2)) + ' (ตามเอกสารแนบ)'] }) });
      const prev = U.addDays(m.from, -1);
      parts.push({ land: true, html: attachHead(p, 'แผนและผลงาน ประจำเดือน ' + U.thMonth(m.ym), 'ระหว่างวันที่ ' + U.thRange(m.from, m.to, true)) +
        boqProgressTable(p, prev < c.start ? '' : prev, m.to) + '<div style="height:8px"></div>' + monthPlanTable(p, m.to) +
        '<table class="two"><tr><td style="width:62%">' + sCurveSvg(p, units, { until: m.to, w: 620, h: 250 }) + '</td><td>' +
        '<p><b>สรุป ณ วันที่ ' + U.thDate(m.to) + '</b></p><p>แผนงานสะสม ' + pl.toFixed(3) + '%<br>ผลงานสะสม ' + ac.toFixed(3) + '%<br>' +
        (diff >= 0 ? 'เร็วกว่าแผน ' : 'ช้ากว่าแผน ') + Math.abs(diff).toFixed(3) + '%</p>' + signBlock(p) + '</td></tr></table>' });
      return { parts: parts, subject: subject, reportKey: m.key, title: subject };
    }
    if (type === 'daily') {
      const d = app.active('daily').find(function (x) { return x.date === o.date2; });
      const photos = app.active('photos').filter(function (x) { return x.date === o.date2; }).slice(0, 12);
      let html = attachHead(p, 'รายงานการปฏิบัติงานประจำวัน', U.thDow(o.date2) + ' ' + U.thDate(o.date2)) + dailyTable(p, o.date2, o.date2);
      if (d) html += '<p><b>ปัญหาอุปสรรค:</b> ' + esc(d.problems || '-') + '<br><b>ข้อสั่งการ/หมายเหตุ:</b> ' + esc([d.orders, d.note].filter(Boolean).join(' / ') || '-') + '</p>';
      html += await photoGrid(photos, 4);
      html += signBlock(p);
      return { parts: [{ land: true, html: html }], subject: 'รายงานประจำวัน ' + U.thDate(o.date2), title: 'รายงานประจำวัน' };
    }
    if (type === 'photos') {
      const list = o.photos || [];
      if (!list.length) throw new Error('ไม่มีรูปตามเงื่อนไขที่เลือก');
      const per = 6; const pages = [];
      for (let i = 0; i < list.length; i += per) pages.push(list.slice(i, i + per));
      for (let i = 0; i < pages.length; i++) {
        parts.push({ html: attachHead(p, o.title || 'รูปถ่ายประกอบรายงาน', o.subtitle || '') + await photoGrid(pages[i], 2, true) +
          '<div class="small c">หน้า ' + (i + 1) + '/' + pages.length + '</div>' });
      }
      return { parts: parts, subject: o.title || 'รูปถ่ายประกอบรายงาน', title: 'รูปถ่าย' };
    }
    if (type === 'delivery') {
      const units = app.active('units');
      const d = o.date || U.today();
      let tot = 0, del = 0;
      const rows = (p.boq || []).map(function (b, i) {
        const val = U.num(b.qty) * U.num(b.unitPrice);
        const f = U.itemFraction(p, b, units, d);
        const dq = U.num(b.qty) * f;
        const dv = b.deliveredAmount !== '' && b.deliveredAmount !== undefined && b.deliveredAmount !== null && f >= 0.999 ? U.num(b.deliveredAmount) : val * f;
        tot += val; del += dv;
        return '<tr><td class="c">' + (i + 1) + '</td><td>' + esc(b.desc) + '</td><td class="c">' + esc(b.unit || '') + '</td><td class="r">' + U.money(b.unitPrice) + '</td>' +
          '<td class="r">' + U.money(b.qty, 0) + '</td><td class="r">' + U.money(val) + '</td><td class="r">' + U.money(dq, dq % 1 ? 2 : 0) + '</td><td class="r">' + U.money(dv) + '</td>' +
          '<td class="r">' + (f * 100).toFixed(2) + '%</td><td>' + (Math.abs(dv - val * f) > 0.01 ? 'จ่ายจริงตามใบเสร็จรับเงิน' : '') + '</td></tr>';
      }).join('');
      const html = '<div class="at-head"><div class="at-t">บัญชีแสดงปริมาณงานและเงินค่างานที่ขอส่ง' + (o.round ? 'ครั้งที่ ' + esc(o.round) : '') + '</div>' +
        '<div>สัญญาจ้างเลขที่ ' + esc(p.contractNo || p.code) + (p.contractDate ? ' ลงวันที่ ' + U.thDate(p.contractDate) : '') + '</div><div>' + esc(p.name || p.shortName) + '</div></div>' +
        '<table class="grid-t"><thead><tr><th rowspan="2">ที่</th><th rowspan="2">รายการงาน</th><th rowspan="2">หน่วย</th><th rowspan="2">ราคาต่อหน่วย (บาท)</th><th colspan="2">ปริมาณงานตามสัญญา</th>' +
        '<th colspan="2">ปริมาณงานที่ขอส่ง</th><th rowspan="2">% แล้วเสร็จ</th><th rowspan="2">หมายเหตุ</th></tr><tr><th>ปริมาณ</th><th>เงิน (บาท)</th><th>ปริมาณ</th><th>เงิน (บาท)</th></tr></thead><tbody>' + rows +
        '</tbody><tfoot><tr><td></td><td>รวม</td><td></td><td></td><td></td><td class="r">' + U.money(tot) + '</td><td></td><td class="r">' + U.money(del) + '</td><td></td><td></td></tr>' +
        '<tr><td></td><td>เงินคงเหลือ</td><td></td><td></td><td></td><td></td><td></td><td class="r">' + U.money(tot - del) + '</td><td></td><td></td></tr></tfoot></table>' + signBlock(p);
      return { parts: [{ land: true, html: html }], subject: 'บัญชีแสดงปริมาณงานและเงินค่างานที่ขอส่ง', title: 'บัญชีปริมาณงาน' };
    }
    if (type === 'tests') {
      const tests = app.active('tests').sort(function (a, b) { return String(a.sampleDate || a.companyDate || '').localeCompare(String(b.sampleDate || b.companyDate || '')); });
      const stn = { pending: 'รอผล', pass: 'ผ่าน', fail: 'ไม่ผ่าน', na: '-' };
      const html = '<div class="at-head"><div class="at-t">รายละเอียดการส่งตัวอย่างและผลการทดสอบ ' + esc(p.code) + '</div><div>' + esc(p.name || p.shortName) + '</div></div>' +
        '<table class="grid-t"><thead><tr><th>ที่</th><th>ประเภท</th><th>รายการ</th><th>เก็บ/เท</th><th>บริษัทส่งหนังสือ</th><th>แขวงส่ง</th><th>หน่วยทดสอบ</th><th>ผลออก/อนุมัติ</th><th>เลขที่ผล</th><th>เกณฑ์</th><th>ผล</th></tr></thead><tbody>' +
        tests.map(function (t, i) {
          return '<tr><td class="c">' + (i + 1) + '</td><td>' + esc(t.type || '') + '</td><td>' + esc(t.item || '') + (t.location ? '<br><span class="small">' + esc(t.location) + '</span>' : '') + '</td>' +
            '<td class="c">' + U.thDate(t.sampleDate, 'short') + '</td><td class="c">' + U.thDate(t.companyDate, 'short') + '</td><td class="c">' + U.thDate(t.sentDate, 'short') + '</td><td>' + esc(t.lab || '') + '</td>' +
            '<td class="c">' + U.thDate(t.resultDate, 'short') + '</td><td>' + esc(t.resultNo || '') + '</td><td>' + esc(t.spec || '') + '</td><td>' + esc((t.value !== '' && t.value !== undefined && t.value !== null ? t.value + ' ' + (t.unit || '') + ' ' : '') + (stn[t.status] || 'รอผล')) + '</td></tr>';
        }).join('') + '</tbody></table>' + signBlock(p);
      return { parts: [{ land: true, html: html }], subject: 'ทะเบียนการส่งตัวอย่างและผลทดสอบ', title: 'ผลทดสอบ' };
    }
    if (type === 'letter') {
      const paras = String(o.body || '').replace('{สัญญา}', contractPara(p)).split(/\n+/).map(function (s) { return s.replace(/^\t/, '').trim(); }).filter(Boolean);
      parts.push({ html: memo(p, { no: o.no, date: o.date, subject: o.subject, to: o.to, paras: paras, closing: o.closing }) });
      return { parts: parts, subject: o.subject, title: o.subject };
    }
    throw new Error('ไม่รู้จักชนิดรายงาน');
  }

  async function photoGrid(list, cols, big) {
    if (!list.length) return '';
    const urls = await Promise.all(list.map(function (x) { return FBL.fileUrl(big ? x.path : (x.thumbPath || x.path)); }));
    let h = '<table class="ph-t">';
    for (let i = 0; i < list.length; i += cols) {
      h += '<tr>';
      for (let j = 0; j < cols; j++) {
        const x = list[i + j];
        h += '<td style="width:' + (100 / cols) + '%">' + (x ? '<img src="' + esc(urls[i + j]) + '" alt=""><div class="ph-cap">' +
          esc([U.thDate(x.date, 'short'), x.km ? 'กม.' + x.km : '', x.side, x.category].filter(Boolean).join(' · ')) + (x.desc ? '<br>' + esc(x.desc) : '') + '</div>' : '') + '</td>';
      }
      h += '</tr>';
    }
    return h + '</table>';
  }

  /* ---------------- ประกอบเป็นเอกสาร ---------------- */
  const PRINT_CSS =
    '@page{size:A4 portrait;margin:15mm 18mm 15mm 25mm}@page land{size:A4 landscape;margin:12mm 12mm 12mm 12mm}' +
    'body{margin:0;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
    '.page{page-break-after:always;break-after:page}.page:last-child{page-break-after:auto;break-after:auto}.land{page:land}' +
    'p{margin:0 0 4pt}.ind{text-indent:2.5cm;text-align:justify}.c{text-align:center}.r{text-align:right}.small{font-size:.82em}.nowrap{white-space:nowrap}' +
    '.memo-title{text-align:center;font-weight:bold;font-size:1.45em;margin:0 0 10pt}' +
    'table.ln{width:100%;border-collapse:collapse}table.ln td{padding:0 0 4pt;width:50%}' +
    'table.sign{width:100%;margin-top:16pt;page-break-inside:avoid}table.sign td{width:50%;vertical-align:top}' +
    '.cc{margin-top:18pt;font-size:.92em}.cc-ind{display:inline-block;width:5.2em}' +
    '.at-head{text-align:center;margin-bottom:8pt}.at-t{font-weight:bold;font-size:1.15em}' +
    'table.grid-t{width:100%;border-collapse:collapse;font-size:.8em}table.grid-t th,table.grid-t td{border:1px solid #000;padding:2pt 4pt;vertical-align:top}' +
    'table.grid-t th{background:#eef2f0;font-weight:bold;text-align:center;vertical-align:middle}table.grid-t tfoot td{font-weight:bold}' +
    'table.grid-t th.vert{writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-weight:normal;padding:4pt 1pt}' +
    'table.two{width:100%;border-collapse:collapse;margin-top:8pt}table.two td{vertical-align:top}table.two svg{width:100%;height:auto}' +
    'table.ph-t{width:100%;border-collapse:collapse;margin-top:6pt}table.ph-t td{padding:4pt;vertical-align:top;text-align:center}' +
    'table.ph-t img{width:100%;max-height:78mm;object-fit:contain;border:1px solid #999}.land table.ph-t img{max-height:60mm}.ph-cap{font-size:.8em;margin-top:2pt}';

  function printDoc(r) {
    return '<!doctype html><html lang="th"><head><meta charset="utf-8"><title>' + esc(r.title || 'รายงาน') + '</title>' +
      '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>' +
      '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap">' +
      '<style>' + PRINT_CSS + 'body{font-family:"TH SarabunPSK","TH Sarabun New","Sarabun",sans-serif;font-size:12.5pt;line-height:1.45}' +
      '@media screen{body{background:#dfe5e2;padding:12px}.page{background:#fff;max-width:210mm;margin:0 auto 12px;padding:15mm 18mm 15mm 25mm;box-shadow:0 2px 8px rgba(0,0,0,.2);box-sizing:border-box;min-height:297mm}.page.land{max-width:297mm;min-height:210mm;padding:12mm}}' +
      '</style></head><body>' + r.parts.map(function (x) { return '<div class="page' + (x.land ? ' land' : '') + '">' + x.html + '</div>'; }).join('') + '</body></html>';
  }
  // Word (.doc แบบ HTML) — เปิดด้วย Microsoft Word ได้ แก้ไขต่อได้ ใช้ฟอนต์ TH SarabunPSK 16
  function wordDoc(r) {
    let secs = '', css = '';
    r.parts.forEach(function (x, i) {
      const n = i + 1;
      css += '@page S' + n + '{size:' + (x.land ? '29.7cm 21cm;mso-page-orientation:landscape;margin:1.2cm' : '21cm 29.7cm;margin:1.5cm 2cm 1.5cm 3cm') + '}div.S' + n + '{page:S' + n + '}';
      secs += (i ? '<br clear=all style="page-break-before:always;mso-break-type:section-break">' : '') + '<div class="S' + n + '">' + x.html.replace(/<svg[\s\S]*?<\/svg>/g, '<p class="small">[กราฟ S-Curve ดูในไฟล์ PDF]</p>') + '</div>';
    });
    return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>' + esc(r.title || '') + '</title>' +
      '<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->' +
      '<style>' + css + PRINT_CSS.replace(/@page[^}]*}/g, '') + 'body{font-family:"TH SarabunPSK";font-size:16pt;line-height:1.1}table.grid-t{font-size:12pt}table.grid-t th.vert{writing-mode:tb-rl;mso-rotate:90}.ph-cap{font-size:12pt}</style></head><body>' + secs + '</body></html>';
  }

  function openPreview(r, meta) {
    meta = meta || {};
    const app = A();
    const html = printDoc(r);
    const m = app.modal({
      title: 'ตัวอย่าง: ' + (r.title || ''), size: 'wide',
      body: '<iframe id="rpFrame" style="width:100%;height:70vh;border:1px solid var(--line);border-radius:8px;background:#dfe5e2"></iframe>',
      foot: (meta.register ? '<button class="btn btn-outline left" id="rpReg">ลงทะเบียนหนังสือออก</button>' : '') +
        '<button class="btn btn-outline" id="rpWord">ดาวน์โหลด Word (.doc)</button><button class="btn btn-primary" id="rpPrint">พิมพ์ / บันทึกเป็น PDF</button>'
    });
    const fr = m.q('#rpFrame');
    fr.srcdoc = html;
    m.q('#rpPrint').onclick = function () { fr.contentWindow.focus(); fr.contentWindow.print(); };
    m.q('#rpWord').onclick = function () {
      U.download(U.safeName((meta.fileName || r.title || 'รายงาน')) + '.doc', new Blob(['﻿' + wordDoc(r)], { type: 'application/msword' }));
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

  return { sCurveSvg: sCurveSvg, build: build, openPreview: openPreview, exportExcel: exportExcel, contractPara: contractPara, printDoc: printDoc };
})();
