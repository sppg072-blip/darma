/* ============================================================
   DARMA — Generator Laporan Manajemen PowerPoint
   - PPT SPPG + Tenaga Kerja (Naker)
   - PPT KDMP
   - Narasi berbasis data, peta SVG lokal, tanpa perubahan schema D1
   ============================================================ */
(function (global) {
  'use strict';

  const C = {
    navy: '0B1F3A', navy2: '15345B', red: 'C8102E', blue: '2563EB', blue2: '60A5FA',
    green: '16A34A', green2: '86EFAC', amber: 'D97706', amber2: 'FDE68A',
    rose: 'B91C1C', slate: '475569', muted: '94A3B8', line: 'E2E8F0',
    pale: 'F8FAFC', white: 'FFFFFF', ink: '0F172A', cyan: '0891B2', purple: '7C3AED'
  };
  const REGION_ORDER = ['Kota Pekalongan', 'Kab. Pekalongan', 'Kab. Batang'];
  const FINDING_RULES = [
    ['Sanitasi & Kebersihan', /sanit|bersih|kotor|lalat|sampah|cuci|limbah|higien/i],
    ['Distribusi & Armada', /distrib|terlambat|kendaraan|armada|kirim|rute|jarak/i],
    ['Bahan Baku & Supplier', /bahan baku|supplier|stok|beras|telur|ayam|sayur|harga/i],
    ['Keamanan Pangan', /keracun|pencernaan|sampel|makanan|basi|segar|kualitas/i],
    ['Administrasi & Keuangan', /laporan|dokumen|administr|keuangan|saldo|top up|petty cash/i],
    ['SDM & SOP', /tenaga|pekerja|sop|beban kerja|juru masak|bpjs|koordinasi/i],
    ['Sarana & Fasilitas', /fasilitas|alat|dapur|gudang|cold storage|bangunan|lahan/i]
  ];

  let reportKind = 'sppg';
  let reportCache = null;
  let reportInitialized = false;

  function dbNow() {
    if (typeof global.getDarmaDB === 'function') return global.getDarmaDB();
    try { return DB; } catch (e) { return { units: [], monitoring: [] }; }
  }
  function n(v) {
    if (v === null || v === undefined || v === '') return 0;
    const x = Number(String(v).replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(x) ? x : 0;
  }
  function sum(a) { return a.reduce((x, y) => x + n(y), 0); }
  function avg(a) { const x = a.map(n).filter(v => Number.isFinite(v)); return x.length ? sum(x) / x.length : 0; }
  function pct(a, b) { return b ? (a / b) * 100 : 0; }
  function fmt(v, dec) { return Number(v || 0).toLocaleString('id-ID', { maximumFractionDigits: dec == null ? 0 : dec }); }
  function fmtPct(v) { return fmt(v, 1) + '%'; }
  function safe(v) { return String(v == null ? '' : v); }
  function html(v) { return safe(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function xml(v) { return safe(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c])); }
  function dateInRange(v, start, end) {
    if (!v) return false;
    return (!start || v >= start) && (!end || v <= end);
  }
  function latestBy(list, keyFn) {
    const m = new Map();
    list.forEach(x => { const k = keyFn(x); const old = m.get(k); if (!old || safe(x.tgl) > safe(old.tgl)) m.set(k, x); });
    return m;
  }
  function topEntries(obj, limit) {
    return Object.entries(obj || {}).sort((a, b) => b[1] - a[1]).slice(0, limit || 6);
  }
  function countVal(list, fn) {
    const out = {};
    list.forEach(x => { const k = safe(fn(x)).trim() || 'Tidak diisi'; out[k] = (out[k] || 0) + 1; });
    return out;
  }
  function arrayVals(v) { return Array.isArray(v) ? v : (v ? [v] : []); }
  function fieldsOf(m) { return (m && m.form && m.form.fields) || {}; }
  function monFormType(m) { return safe(m && m.formType || (m && m.jenis === 'KDMP' ? 'KDMP' : 'SPPG')).toUpperCase(); }
  function formHasData(m) { const f = fieldsOf(m); return Object.keys(f).some(k => f[k] !== '' && f[k] != null); }
  function monthName(iso) {
    if (!iso) return '';
    const d = new Date(iso + 'T00:00:00');
    return isNaN(d) ? iso : d.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
  }

  function getConfig() {
    const gv = id => { const e = document.getElementById(id); return e ? e.value.trim() : ''; };
    const type = reportKind;
    const defaultTitle = type === 'sppg'
      ? 'Laporan Manajemen Monitoring SPPG dan Tenaga Kerja'
      : 'Laporan Manajemen Monitoring dan Evaluasi KDMP';
    return {
      type, start: gv('rptStart'), end: gv('rptEnd'), region: gv('rptRegion'),
      mode: gv('rptMode') || 'lengkap', author: gv('rptAuthor') || 'Tim Monitoring DARMA',
      title: gv('rptTitle') || defaultTitle,
      organization: gv('rptOrg') || 'DARMA — Dashboard Akurat Reporting Manajemen Aplikasi'
    };
  }

  function baseAnalysis(cfg) {
    const db = dbNow();
    const wantedType = cfg.type === 'kdmp' ? 'KDMP' : 'SPPG';
    const units = (db.units || []).filter(u => u.jenis === wantedType && (!cfg.region || u.kab === cfg.region));
    const ids = new Set(units.map(u => u.id));
    const monitoring = (db.monitoring || []).filter(m => ids.has(m.unitId) && dateInRange(m.tgl, cfg.start, cfg.end));
    const visitsByUnit = {};
    monitoring.forEach(m => { visitsByUnit[m.unitId] = (visitsByUnit[m.unitId] || 0) + 1; });
    const visited = new Set(Object.keys(visitsByUnit));
    const latest = latestBy(monitoring, m => m.unitId);
    const status = countVal(units, u => u.status || 'tidak diketahui');
    const byRegion = {};
    REGION_ORDER.forEach(r => {
      const us = units.filter(u => u.kab === r); const ui = new Set(us.map(u => u.id));
      const vm = monitoring.filter(m => ui.has(m.unitId)); const vu = new Set(vm.map(m => m.unitId));
      byRegion[r] = { units: us.length, monitored: vu.size, visits: vm.length, coverage: pct(vu.size, us.length) };
    });
    const freq = { '0 kali': 0, '1 kali': 0, '2 kali': 0, '3 kali': 0, '4 kali': 0, '5+ kali': 0 };
    units.forEach(u => { const c = visitsByUnit[u.id] || 0; const k = c === 0 ? '0 kali' : c >= 5 ? '5+ kali' : c + ' kali'; freq[k]++; });
    const priority = [];
    units.forEach(u => {
      const m = latest.get(u.id); let score = 0; const reasons = [];
      if (!visited.has(u.id)) { score += u.status === 'aktif' ? 5 : 3; reasons.push('Belum dimonitor'); }
      if (m && m.hasil === 'kritis') { score += 8; reasons.push('Hasil kritis'); }
      else if (m && m.hasil === 'perbaikan') { score += 4; reasons.push('Perlu perbaikan'); }
      if (u.status === 'kendala') { score += 5; reasons.push('Status kendala'); }
      if (u.status === 'aktif' && !n(u.kapasitas) && u.jenis === 'SPPG') { score += 2; reasons.push('Kapasitas belum terisi'); }
      if (score) priority.push({ unit: u, mon: m, score, reasons });
    });
    priority.sort((a, b) => b.score - a.score || a.unit.nama.localeCompare(b.unit.nama));
    return {
      db, cfg, units, monitoring, visitsByUnit, visited, latest, status, byRegion, freq, priority,
      total: units.length, visits: monitoring.length, monitored: visited.size,
      unmonitored: units.length - visited.size, coverage: pct(visited.size, units.length),
      critical: [...latest.values()].filter(m => m.hasil === 'kritis').length,
      improvement: [...latest.values()].filter(m => m.hasil === 'perbaikan').length
    };
  }

  function findingSummary(monitoring, extraTexts) {
    const counts = {}; FINDING_RULES.forEach(([k]) => counts[k] = 0);
    const texts = [];
    monitoring.forEach(m => {
      const t = [m.temuan, m.rekom].filter(Boolean).join(' '); if (t) texts.push({ text: t, unitId: m.unitId });
    });
    (extraTexts || []).forEach(t => { if (t) texts.push({ text: safe(t), unitId: '' }); });
    texts.forEach(o => FINDING_RULES.forEach(([k, re]) => { if (re.test(o.text)) counts[k]++; }));
    return { counts, samples: texts.slice(0, 8) };
  }

  function analyzeSPPG(cfg) {
    const a = baseAnalysis(cfg);
    const sppgForms = a.monitoring.filter(m => monFormType(m) === 'SPPG' && m.form && formHasData(m));
    const nakerForms = a.monitoring.filter(m => monFormType(m) === 'NAKER' && m.form && formHasData(m));
    const latestSPPG = latestBy(sppgForms, m => m.unitId);
    const snapshots = [...latestSPPG.values()];
    const benefKeys = ['siswa', 'ibuhamil', 'balita', 'guru', 'posyandu'];
    const benefLabels = { siswa: 'Siswa', ibuhamil: 'Ibu Hamil & Menyusui', balita: 'Balita', guru: 'Guru & Tendik', posyandu: 'Kader Posyandu' };
    const benef = {}; benefKeys.forEach(k => benef[k] = 0);
    const schoolKeys = ['paud', 'tk', 'sd', 'smp', 'sma'];
    const schoolLabels = { paud: 'PAUD', tk: 'TK/RA', sd: 'SD/MI', smp: 'SMP/MTs', sma: 'SMA/SMK/MA' };
    const schools = {}; schoolKeys.forEach(k => schools[k] = 0);
    let benefFilled = 0, schoolFilled = 0;
    const workers = [], cooks = [], days = [], supplierCounts = [];
    const finance = {
      supplierSpend: { KDKMP: 0, 'BUMDes/Koperasi': 0, 'Agen/Pasar': 0, Distributor: 0, Produsen: 0, UMKM: 0 },
      originSpend: { 'Dalam kota': 0, 'Luar kota': 0 }, operational: {}, obstacles: {}, filled: 0
    };
    const dist = { '< 15 menit': 0, '< 30 menit': 0, '> 30 menit': 0, 'Tidak diisi': 0 };
    const bpjs = { '0-20%': 0, '21-50%': 0, '51-80%': 0, '81-100%': 0, 'Tidak diisi': 0 };
    const safety = { noExternal: 0, noSample: 0, digestive: 0, riskyWaste: 0, filled: 0 };
    const extraTexts = [];
    snapshots.forEach(m => {
      const f = fieldsOf(m); const b = f.sp201 || {}, sc = f.sp202 || {};
      if (benefKeys.some(k => n(b[k]) > 0)) benefFilled++;
      if (schoolKeys.some(k => n(sc[k]) > 0)) schoolFilled++;
      benefKeys.forEach(k => benef[k] += n(b[k])); schoolKeys.forEach(k => schools[k] += n(sc[k]));
      if (f.sp205 !== '' && f.sp205 != null) workers.push(n(f.sp205));
      if (f.sp206 !== '' && f.sp206 != null) cooks.push(n(f.sp206));
      if (f.sp203 !== '' && f.sp203 != null) days.push(n(f.sp203));
      if (f.sp409 !== '' && f.sp409 != null) supplierCounts.push(n(f.sp409));
      const ss=f.sp410||{};
      finance.supplierSpend.KDKMP+=n(ss.kdkmp);finance.supplierSpend['BUMDes/Koperasi']+=n(ss.bumdes);finance.supplierSpend['Agen/Pasar']+=n(ss.agen);finance.supplierSpend.Distributor+=n(ss.distributor);finance.supplierSpend.Produsen+=n(ss.produsen);finance.supplierSpend.UMKM+=n(ss.umkm);
      const os=f.sp411||{};Object.values(os).forEach(v=>{if(v&&typeof v==='object'){finance.originSpend['Dalam kota']+=n(v.dalam);finance.originSpend['Luar kota']+=n(v.luar);}});
      const op=f.sp413||{};Object.entries(op).forEach(([k,v])=>finance.operational[k]=(finance.operational[k]||0)+n(v));
      arrayVals(f.sp414).forEach(x=>finance.obstacles[x]=(finance.obstacles[x]||0)+1);
      if(Object.keys(ss).length||Object.keys(os).length||Object.keys(op).length||arrayVals(f.sp414).length)finance.filled++;
      const dv = safe(f.sp204); if (!dv) dist['Tidak diisi']++; else if (dv.includes('> 30')) dist['> 30 menit']++; else if (dv.includes('30')) dist['< 30 menit']++; else dist['< 15 menit']++;
      const bv = safe(f.sp208) || 'Tidak diisi'; bpjs[bv] = (bpjs[bv] || 0) + 1;
      if (f.sp301 || f.sp308 || f.sp310 || f.sp311) safety.filled++;
      if (safe(f.sp301).toLowerCase() === 'tidak') safety.noExternal++;
      if (safe(f.sp308).toLowerCase() === 'tidak ada') safety.noSample++;
      if (/^ya/i.test(safe(f.sp310))) safety.digestive++;
      if (/sungai|tpa/i.test(safe(f.sp311))) safety.riskyWaste++;
      extraTexts.push(f.sp312); arrayVals(f.sp414).forEach(x => extraTexts.push(x));
    });
    const nf = nakerForms.map(m => fieldsOf(m));
    const naker = {
      responses: nf.length,
      positions: countVal(nf, f => f.nk102), education: countVal(nf, f => f.nk103),
      priorWork: countVal(nf, f => f.nk201), sop: countVal(nf, f => f.nk301),
      impact: countVal(nf, f => f.nk308), overtime: countVal(nf, f => f.nk208),
      avgDays: avg(nf.map(f => f.nk205).filter(v => v !== '')),
      avgHours: avg(nf.map(f => f.nk206).filter(v => v !== '')),
      avgOldWage: avg(nf.map(f => f.nk203).filter(v => v !== '')),
      avgNewWage: avg(nf.map(f => f.nk207).filter(v => v !== '')),
      constraints: {}
    };
    nf.forEach(f => arrayVals(f.nk304).forEach(x => naker.constraints[x] = (naker.constraints[x] || 0) + 1));
    const findings = findingSummary(a.monitoring, extraTexts);
    const capTotal = sum(a.units.map(u => u.kapasitas));
    const benefTotal = sum(Object.values(benef)), schoolTotal = sum(Object.values(schools));
    const completeness = {
      sppgForms: snapshots.length, benef: benefFilled, schools: schoolFilled,
      safety: safety.filled, naker: naker.responses
    };
    const period = cfg.start && cfg.end ? `${cfg.start} s.d. ${cfg.end}` : 'seluruh periode';
    const parts = [
      `Pada ${period}, terdapat ${fmt(a.total)} SPPG dalam cakupan laporan. ${fmt(a.monitored)} unit telah memiliki kunjungan monitoring (${fmtPct(a.coverage)}), sedangkan ${fmt(a.unmonitored)} unit belum dimonitor.`,
      snapshots.length
        ? `Form operasional SPPG terbaru tersedia untuk ${fmt(snapshots.length)} unit (${fmtPct(pct(snapshots.length, a.total))} dari unit terdaftar).`
        : 'Belum tersedia form operasional SPPG terisi pada periode ini; analisis rinci ditampilkan sebagai kebutuhan kelengkapan data.'
    ];
    if (benefFilled) parts.push(`Dari ${fmt(benefFilled)} SPPG yang mengisi data penerima, tercatat ${fmt(benefTotal)} penerima/porsi per hari; kelompok siswa berjumlah ${fmt(benef.siswa)} (${fmtPct(pct(benef.siswa, benefTotal))}).`);
    if (schoolFilled) parts.push(`Jaringan layanan mencakup ${fmt(schoolTotal)} sekolah; jenjang terbanyak adalah ${topEntries(schools, 1).map(([k, v]) => schoolLabels[k] + ' (' + fmt(v) + ')')[0]}.`);
    if (safety.digestive || safety.noSample || safety.riskyWaste) parts.push(`Terdapat sinyal risiko keamanan pangan: ${fmt(safety.digestive)} laporan gangguan pencernaan, ${fmt(safety.noSample)} unit tanpa uji sampel, dan ${fmt(safety.riskyWaste)} unit dengan pengelolaan limbah berisiko.`);
    const recs = [];
    if (a.unmonitored) recs.push(`Prioritaskan monitoring pada ${fmt(a.unmonitored)} SPPG yang belum dikunjungi, terutama unit berstatus aktif.`);
    if (a.critical) recs.push(`Lakukan tindak lanjut maksimal 14 hari pada ${fmt(a.critical)} unit dengan hasil kritis.`);
    if (dist['> 30 menit']) recs.push(`Evaluasi rute dan armada pada ${fmt(dist['> 30 menit'])} SPPG dengan waktu distribusi lebih dari 30 menit.`);
    if (safety.digestive || safety.noSample) recs.push('Perkuat verifikasi keamanan pangan, uji sampel, dan pengawasan eksternal pada unit berisiko.');
    if (benefFilled < a.total * .7) recs.push(`Tingkatkan kelengkapan data penerima; cakupan isian saat ini ${fmtPct(pct(benefFilled, a.total))}.`);
    if (naker.responses) recs.push(`Gunakan ${fmt(naker.responses)} respons Naker untuk menindaklanjuti beban kerja, SOP, perlindungan, dan dampak ekonomi pekerja.`);
    if (!recs.length) recs.push('Pertahankan cakupan monitoring dan lakukan evaluasi berkala berbasis risiko.');
    return Object.assign(a, {
      kind: 'sppg', sppgForms, nakerForms, snapshots, benef, benefLabels, schools, schoolLabels,
      benefTotal, schoolTotal, workers, cooks, days, supplierCounts, finance, dist, bpjs, safety, naker,
      findings, capTotal, completeness, executive: parts.join(' '), recommendations: recs
    });
  }

  function analyzeKDMP(cfg) {
    const a = baseAnalysis(cfg);
    const forms = a.monitoring.filter(m => monFormType(m) === 'KDMP' && m.form);
    const latestForms = latestBy(forms, m => m.unitId); const snapshots = [...latestForms.values()];
    const dims = {};
    const dimAnswers = {};
    const compliance = { yes: 0, total: 0, units: 0 };
    const open = { strengths: [], constraints: [], suggestions: [], products: [] };
    const scoreCategories = {};
    const extraTexts = [];
    snapshots.forEach(m => {
      const f = m.form || {};
      (f.sections || []).forEach(s => {
        const vals = (s.scores || []).map(n).filter(v => v > 0);
        if (!vals.length) return;
        const k = s.judul || s.kode || 'Dimensi';
        dims[k] = (dims[k] || 0) + sum(vals); dimAnswers[k] = (dimAnswers[k] || 0) + vals.length;
      });
      const comp = f.compliance || []; if (comp.length) compliance.units++;
      compliance.yes += comp.filter(x => String(x).toLowerCase() === 'ya').length;
      compliance.total += comp.filter(Boolean).length;
      const cat=safe(f.kategori)||'Belum dikategorikan';scoreCategories[cat]=(scoreCategories[cat]||0)+1;
      if(f.open_kelebihan)open.strengths.push(f.open_kelebihan);if(f.open_kendala)open.constraints.push(f.open_kendala);if(f.open_saran)open.suggestions.push(f.open_saran);if(f.open_produk)open.products.push(f.open_produk);
      [f.open_kelebihan, f.open_kendala, f.open_saran, f.open_produk].forEach(x => extraTexts.push(x));
    });
    Object.keys(dims).forEach(k => dims[k] = dims[k] / (dimAnswers[k] || 1));
    const scores = snapshots.map(m => n(m.form && m.form.avg)).filter(v => v > 0);
    const avgScore = avg(scores);
    const findings = findingSummary(a.monitoring, extraTexts);
    const period = cfg.start && cfg.end ? `${cfg.start} s.d. ${cfg.end}` : 'seluruh periode';
    const parts = [
      `Pada ${period}, terdapat ${fmt(a.total)} KDMP dalam cakupan laporan. ${fmt(a.monitored)} unit telah dimonitor (${fmtPct(a.coverage)}), sementara ${fmt(a.unmonitored)} unit belum dimonitor.`,
      snapshots.length ? `Kuesioner KDMP lengkap tersedia untuk ${fmt(snapshots.length)} unit dengan skor rata-rata ${fmt(avgScore, 2)} dari 4,00.` : 'Belum tersedia kuesioner KDMP terisi untuk analisis dimensi.'
    ];
    if (Object.keys(dims).length) {
      const high = topEntries(dims, 1)[0], low = Object.entries(dims).sort((x, y) => x[1] - y[1])[0];
      parts.push(`Dimensi tertinggi adalah ${high[0]} (${fmt(high[1], 2)}), sedangkan perhatian terbesar diperlukan pada ${low[0]} (${fmt(low[1], 2)}).`);
    }
    if (compliance.total) parts.push(`Tingkat kepatuhan jawaban “Ya” mencapai ${fmtPct(pct(compliance.yes, compliance.total))}.`);
    const recs = [];
    if (a.unmonitored) recs.push(`Jadwalkan monitoring pada ${fmt(a.unmonitored)} KDMP yang belum dikunjungi.`);
    Object.entries(dims).filter(([, v]) => v < 2.51).sort((x, y) => x[1] - y[1]).slice(0, 3).forEach(([k, v]) => recs.push(`Prioritaskan penguatan ${k} yang memperoleh skor ${fmt(v, 2)}.`));
    if (compliance.total && pct(compliance.yes, compliance.total) < 80) recs.push('Susun tindak lanjut pemenuhan dokumen, BAST, sarana, dan kepatuhan regulasi yang belum lengkap.');
    if (!recs.length) recs.push('Pertahankan tata kelola dan dorong pengembangan usaha berbasis potensi lokal.');
    return Object.assign(a, {
      kind: 'kdmp', forms, snapshots, dims, compliance, avgScore, findings, open, scoreCategories,
      executive: parts.join(' '), recommendations: recs,
      completeness: { forms: snapshots.length, scored: scores.length, compliance: compliance.units }
    });
  }

  function analyze(cfg) { return cfg.type === 'kdmp' ? analyzeKDMP(cfg) : analyzeSPPG(cfg); }

  function selectReportKind(kind) {
    reportKind = kind === 'kdmp' ? 'kdmp' : 'sppg';
    document.querySelectorAll('.rpt-type-card').forEach(x => x.classList.toggle('active', x.dataset.kind === reportKind));
    const title = document.getElementById('rptTitle');
    if (title) title.value = reportKind === 'sppg'
      ? 'Laporan Manajemen Monitoring SPPG dan Tenaga Kerja'
      : 'Laporan Manajemen Monitoring dan Evaluasi KDMP';
    markReportStale();
  }

  function initReportPanel() {
    if (!reportInitialized) {
      const now = new Date();
      const end = now.toISOString().slice(0, 10);
      const start = end.slice(0, 8) + '01';
      const s = document.getElementById('rptStart'), e = document.getElementById('rptEnd');
      if (s && !s.value) s.value = start; if (e && !e.value) e.value = end;
      reportInitialized = true;
      selectReportKind(reportKind);
    } else {
      document.querySelectorAll('.rpt-type-card').forEach(x => x.classList.toggle('active', x.dataset.kind === reportKind));
    }
    if(!reportCache)previewManagementReport(false);
  }

  function markReportStale() {
    reportCache = null;
    const st = document.getElementById('rptStatus');
    if (st) st.innerHTML = '<i class="fas fa-circle" style="color:#f59e0b"></i> Konfigurasi berubah — klik Tinjau Ringkasan.';
  }

  function previewManagementReport(showToast) {
    try {
      const cfg = getConfig(); const data = analyze(cfg); reportCache = { cfg, data };
      const kpi = document.getElementById('rptKpis');
      if (kpi) {
        const extra = cfg.type === 'sppg'
          ? [{ v: fmt(data.benefTotal), l: 'Penerima/hari terisi' }, { v: fmt(data.naker.responses), l: 'Respons Naker' }]
          : [{ v: data.snapshots.length ? fmt(data.avgScore, 2) : '—', l: 'Skor KDMP / 4' }, { v: fmtPct(pct(data.compliance.yes, data.compliance.total)), l: 'Kepatuhan' }];
        const cards = [
          { v: fmt(data.total), l: cfg.type === 'sppg' ? 'SPPG' : 'KDMP' },
          { v: fmt(data.monitored), l: 'Sudah dimonitor' }, { v: fmt(data.unmonitored), l: 'Belum dimonitor' },
          { v: fmtPct(data.coverage), l: 'Cakupan' }, ...extra
        ];
        kpi.innerHTML = cards.map((x, i) => `<div class="rpt-kpi k${i + 1}"><b>${html(x.v)}</b><span>${html(x.l)}</span></div>`).join('');
      }
      const nar = document.getElementById('rptNarrative'); if (nar) nar.value = data.executive;
      const rec = document.getElementById('rptRecommendations'); if (rec) rec.value = data.recommendations.map((x, i) => `${i + 1}. ${x}`).join('\n');
      const note = document.getElementById('rptDataNote');
      if (note) note.innerHTML = buildDataNote(data);
      const st = document.getElementById('rptStatus');
      if (st) st.innerHTML = `<i class="fas fa-check-circle" style="color:#16a34a"></i> Ringkasan siap · ${fmt(data.total)} unit · ${fmt(data.visits)} kunjungan dalam periode.`;
      if (showToast && typeof toast === 'function') toast('📊 Ringkasan laporan berhasil dihitung');
      return data;
    } catch (e) {
      console.error(e); if (typeof toast === 'function') toast('Gagal menghitung laporan: ' + e.message, 'e'); return null;
    }
  }

  function buildDataNote(d) {
    if (d.kind === 'sppg') return `<b>Kelengkapan sumber:</b> Form SPPG ${fmt(d.completeness.sppgForms)}/${fmt(d.total)} · Penerima ${fmt(d.completeness.benef)}/${fmt(d.total)} · Sekolah ${fmt(d.completeness.schools)}/${fmt(d.total)} · Keamanan pangan ${fmt(d.completeness.safety)}/${fmt(d.total)} · Respons Naker ${fmt(d.completeness.naker)}.`;
    return `<b>Kelengkapan sumber:</b> Form KDMP ${fmt(d.completeness.forms)}/${fmt(d.total)} · Skor ${fmt(d.completeness.scored)}/${fmt(d.total)} · Kepatuhan ${fmt(d.completeness.compliance)}/${fmt(d.total)}.`;
  }

  async function getGeoJSON() {
    try { if (typeof currentGeoJSON !== 'undefined' && currentGeoJSON) return currentGeoJSON; } catch (e) {}
    const r = await fetch('./regencies.json'); return r.json();
  }
  function coordsFlat(x, out) {
    if (!Array.isArray(x)) return;
    if (typeof x[0] === 'number' && typeof x[1] === 'number') out.push(x);
    else x.forEach(v => coordsFlat(v, out));
  }
  function svgData(svg) {
    const enc = unescape(encodeURIComponent(svg));
    return 'data:image/svg+xml;base64,' + btoa(enc);
  }
  function mapSvg(geo, data) {
    const W = 1000, H = 560, pad = 28, pts = [];
    (geo.features || []).forEach(f => coordsFlat(f.geometry.coordinates, pts));
    if (!pts.length) return '';
    const minX = Math.min(...pts.map(p => p[0])), maxX = Math.max(...pts.map(p => p[0]));
    const minY = Math.min(...pts.map(p => p[1])), maxY = Math.max(...pts.map(p => p[1]));
    const project = p => [pad + ((p[0] - minX) / (maxX - minX || 1)) * (W - pad * 2), H - pad - ((p[1] - minY) / (maxY - minY || 1)) * (H - pad * 2)];
    function ringPath(r) { return r.map((p, i) => { const q = project(p); return (i ? 'L' : 'M') + q[0].toFixed(1) + ' ' + q[1].toFixed(1); }).join(' ') + ' Z'; }
    function pathsOf(g) {
      const polys = g.type === 'Polygon' ? [g.coordinates] : g.coordinates;
      return polys.map(poly => poly.map(ringPath).join(' '));
    }
    const fills = { 'Kota Pekalongan': '#DBEAFE', Pekalongan: '#E0E7FF', Batang: '#FEE2E2' };
    let shapes = '', labels = '';
    (geo.features || []).forEach(f => {
      const name = f.properties.NAME_2 || ''; const fill = fills[name] || '#E2E8F0';
      pathsOf(f.geometry).forEach(d => shapes += `<path d="${d}" fill="${fill}" stroke="#64748B" stroke-width="2" fill-rule="evenodd"/>`);
      const cp = []; coordsFlat(f.geometry.coordinates, cp); const cx = avg(cp.map(x => x[0])), cy = avg(cp.map(x => x[1])); const q = project([cx, cy]);
      labels += `<text x="${q[0]}" y="${q[1]}" text-anchor="middle" font-family="Arial" font-size="18" font-weight="700" fill="#334155">${xml(name)}</text>`;
    });
    let markers = '';
    data.units.forEach(u => {
      const q = project([n(u.lng), n(u.lat)]), c = data.visited.has(u.id) ? '#16A34A' : (u.jenis === 'KDMP' ? '#DC2626' : '#2563EB');
      const vc = data.visitsByUnit[u.id] || 0, critical = data.latest.get(u.id) && data.latest.get(u.id).hasil === 'kritis';
      if (vc) markers += `<circle cx="${q[0]}" cy="${q[1]}" r="${5 + Math.min(vc, 5) * 1.8}" fill="none" stroke="#16A34A" stroke-width="1.5" opacity=".42"/>`;
      if (critical) markers += `<circle cx="${q[0]}" cy="${q[1]}" r="11" fill="none" stroke="#B91C1C" stroke-width="3"/>`;
      markers += `<circle cx="${q[0]}" cy="${q[1]}" r="5" fill="${c}" stroke="#FFFFFF" stroke-width="2"/>`;
    });
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="100%" height="100%" rx="24" fill="#F8FAFC"/>${shapes}${labels}${markers}</svg>`;
  }

  function addText(slide, text, x, y, w, h, opts) {
    slide.addText(safe(text), Object.assign({ x, y, w, h, fontFace: 'Aptos', fontSize: 13, color: C.ink, margin: 0, breakLine: false, valign: 'mid', fit: 'shrink' }, opts || {}));
  }
  function addShape(slide, pptx, type, x, y, w, h, opts) {
    slide.addShape(type || pptx.ShapeType.rect, Object.assign({ x, y, w, h, line: { color: C.line, transparency: 100 }, fill: { color: C.white } }, opts || {}));
  }
  function addHeader(slide, pptx, title, subtitle, page) {
    addShape(slide, pptx, pptx.ShapeType.rect, 0, 0, 13.333, .18, { fill: { color: C.red }, line: { transparency: 100 } });
    addText(slide, title, .55, .36, 9.8, .38, { fontFace: 'Aptos Display', fontSize: 23, bold: true, color: C.navy });
    if (subtitle) addText(slide, subtitle, .57, .82, 10.6, .24, { fontSize: 9.5, color: C.slate });
    addText(slide, 'DARMA', 11.55, .35, 1.2, .3, { fontSize: 16, bold: true, color: C.red, align: 'right' });
    addShape(slide, pptx, pptx.ShapeType.line, .55, 7.16, 12.2, 0, { line: { color: C.line, width: 1 } });
    addText(slide, 'Sumber: Master Unit dan Monitoring DARMA', .58, 7.22, 6, .15, { fontSize: 7.5, color: C.muted });
    addText(slide, String(page || ''), 12.15, 7.2, .55, .16, { fontSize: 8, color: C.muted, align: 'right' });
  }
  function kpi(slide, pptx, x, y, w, h, value, label, color, note) {
    addShape(slide, pptx, pptx.ShapeType.roundRect, x, y, w, h, { rectRadius: .08, fill: { color: C.white }, line: { color: C.line, width: 1 }, shadow: { type: 'outer', color: '94A3B8', opacity: .15, blur: 1, angle: 45, distance: 1 } });
    addShape(slide, pptx, pptx.ShapeType.rect, x, y, .08, h, { fill: { color }, line: { transparency: 100 } });
    addText(slide, value, x + .2, y + .14, w - .35, .38, { fontSize: 22, bold: true, color });
    addText(slide, label, x + .2, y + .58, w - .35, .22, { fontSize: 9.5, bold: true, color: C.slate });
    if (note) addText(slide, note, x + .2, y + .84, w - .35, .18, { fontSize: 7.5, color: C.muted });
  }
  function panel(slide, pptx, x, y, w, h, title, fill) {
    addShape(slide, pptx, pptx.ShapeType.roundRect, x, y, w, h, { rectRadius: .06, fill: { color: fill || C.white }, line: { color: C.line, width: 1 } });
    addText(slide, title, x + .22, y + .14, w - .44, .26, { fontSize: 12, bold: true, color: C.navy });
  }
  function barList(slide, pptx, entries, x, y, w, h, color, maxValue, suffix) {
    const rows = entries.length || 1, rh = h / rows, max = maxValue || Math.max(1, ...entries.map(x => n(x[1])));
    entries.forEach((e, i) => {
      const yy = y + i * rh, val = n(e[1]), bw = Math.max(.03, (w - 2.25) * (val / max));
      addText(slide, e[0], x, yy, 2.05, rh - .03, { fontSize: 8.5, color: C.slate });
      addShape(slide, pptx, pptx.ShapeType.roundRect, x + 2.08, yy + .1, w - 2.7, Math.max(.11, rh - .22), { fill: { color: 'E2E8F0' }, line: { transparency: 100 }, rectRadius: .04 });
      addShape(slide, pptx, pptx.ShapeType.roundRect, x + 2.08, yy + .1, bw, Math.max(.11, rh - .22), { fill: { color }, line: { transparency: 100 }, rectRadius: .04 });
      addText(slide, fmt(val, suffix === ' skor' ? 2 : 0) + (suffix || ''), x + w - .56, yy, .55, rh - .03, { fontSize: 8.5, bold: true, color, align: 'right' });
    });
  }
  function narrative(slide, pptx, text, x, y, w, h, title) {
    panel(slide, pptx, x, y, w, h, title || 'Interpretasi Manajerial', 'F8FAFC');
    addText(slide, text, x + .25, y + .52, w - .5, h - .68, { fontSize: 10.5, color: C.slate, breakLine: true, valign: 'top', margin: .04, fit: 'shrink' });
  }
  function bullets(slide, items, x, y, w, h, color) {
    const runs = [];
    (items || []).forEach((t, i) => { runs.push({ text: safe(t), options: { bullet: { indent: 12 }, hanging: 3, breakLine: i < items.length - 1 } }); });
    slide.addText(runs.length ? runs : [{ text: 'Belum ada rekomendasi.' }], { x, y, w, h, fontFace: 'Aptos', fontSize: 11, color: color || C.slate, margin: .04, breakLine: true, valign: 'top', fit: 'shrink', paraSpaceAfterPt: 9 });
  }
  function coverSlide(pptx, d, cfg) {
    const s = pptx.addSlide(); s.background = { color: C.navy };
    addShape(s, pptx, pptx.ShapeType.rect, 0, 0, .22, 7.5, { fill: { color: C.red }, line: { transparency: 100 } });
    addShape(s, pptx, pptx.ShapeType.arc, 9.25, -.75, 4.8, 4.8, { adjustPoint: .35, rotate: 18, fill: { color: C.blue, transparency: 45 }, line: { color: C.blue2, transparency: 55, width: 2 } });
    addShape(s, pptx, pptx.ShapeType.ellipse, 10.65, 4.75, 2.2, 2.2, { fill: { color: C.red, transparency: 25 }, line: { transparency: 100 } });
    addText(s, 'DARMA', .72, .62, 2.6, .45, { fontSize: 27, bold: true, color: C.white, charSpacing: 2.5 });
    addText(s, cfg.title, .75, 1.58, 8.55, 1.55, { fontFace: 'Aptos Display', fontSize: 30, bold: true, color: C.white, valign: 'mid', fit: 'shrink' });
    addText(s, cfg.type === 'sppg' ? 'SPPG & TENAGA KERJA' : 'KOPERASI DESA MERAH PUTIH', .78, 3.22, 6.8, .3, { fontSize: 13, bold: true, color: C.blue2, charSpacing: 1.7 });
    addText(s, `${cfg.start || 'Awal data'} — ${cfg.end || 'Data terbaru'}${cfg.region ? '  |  ' + cfg.region : '  |  Semua Wilayah'}`, .78, 3.72, 7.7, .34, { fontSize: 13, color: 'CBD5E1' });
    kpi(s, pptx, .78, 4.48, 2.15, 1.12, fmt(d.total), cfg.type === 'sppg' ? 'SPPG tercakup' : 'KDMP tercakup', C.blue, 'Basis unit laporan');
    kpi(s, pptx, 3.08, 4.48, 2.15, 1.12, fmtPct(d.coverage), 'Cakupan monitoring', C.green, `${fmt(d.monitored)} unit`);
    kpi(s, pptx, 5.38, 4.48, 2.15, 1.12, fmt(d.visits), 'Kunjungan periode', C.amber, 'Aktivitas monitoring');
    addText(s, cfg.organization, .78, 6.48, 8.4, .26, { fontSize: 10, color: 'CBD5E1' });
    addText(s, `Disusun oleh: ${cfg.author}`, .78, 6.82, 8.4, .22, { fontSize: 9, color: C.muted });
    return s;
  }

  function executiveSlide(pptx, d, cfg, page, editedNarrative) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Ringkasan Eksekutif', 'Fakta utama, cakupan, dan sinyal prioritas', page);
    const items = [
      [fmt(d.total), cfg.type === 'sppg' ? 'Total SPPG' : 'Total KDMP', C.blue],
      [fmt(d.monitored), 'Sudah dimonitor', C.green], [fmt(d.unmonitored), 'Belum dimonitor', C.amber],
      [fmtPct(d.coverage), 'Cakupan', C.cyan], [fmt(d.critical), 'Hasil kritis', C.rose]
    ];
    items.forEach((x, i) => kpi(s, pptx, .55 + i * 2.5, 1.28, 2.22, 1.08, x[0], x[1], x[2]));
    narrative(s, pptx, editedNarrative || d.executive, .55, 2.68, 8.1, 3.92, 'Narasi Otomatis Berdasarkan Data Lapangan');
    panel(s, pptx, 8.9, 2.68, 3.88, 3.92, 'Fokus Keputusan', 'FFF7ED');
    bullets(s, d.recommendations.slice(0, 6), 9.16, 3.24, 3.35, 3.05, C.slate);
  }

  function mapSlide(pptx, d, cfg, geo, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Peta Sebaran & Progres Monitoring', 'Titik, status kunjungan, dan konsentrasi wilayah', page);
    const svg = mapSvg(geo, d);
    if (svg) s.addImage({ data: svgData(svg), x: .55, y: 1.22, w: 8.5, h: 5.63 });
    panel(s, pptx, 9.25, 1.22, 3.53, 2.12, 'Legenda & Cakupan', 'F8FAFC');
    const legend = [
      ['Sudah dimonitor', C.green], [cfg.type === 'sppg' ? 'SPPG belum' : 'KDMP belum', cfg.type === 'sppg' ? C.blue : C.red],
      ['Halo merah: kritis', C.rose], ['Cincin: frekuensi', C.green]
    ];
    legend.forEach((x, i) => { addShape(s, pptx, pptx.ShapeType.ellipse, 9.53, 1.79 + i * .35, .16, .16, { fill: { color: x[1] }, line: { color: C.white, width: 1 } }); addText(s, x[0], 9.8, 1.75 + i * .35, 2.55, .22, { fontSize: 9, color: C.slate }); });
    panel(s, pptx, 9.25, 3.58, 3.53, 3.27, 'Per Wilayah', 'FFFFFF');
    const entries = REGION_ORDER.filter(r => !cfg.region || r === cfg.region).map(r => [r.replace('Kab. ', ''), d.byRegion[r].units]);
    barList(s, pptx, entries, 9.52, 4.18, 2.95, 1.55, cfg.type === 'sppg' ? C.blue : C.red);
    const lowest = REGION_ORDER.map(r => [r, d.byRegion[r].coverage]).filter(x => d.byRegion[x[0]].units).sort((a, b) => a[1] - b[1])[0];
    if (lowest) addText(s, `Cakupan terendah: ${lowest[0]} (${fmtPct(lowest[1])}).`, 9.52, 5.95, 2.88, .48, { fontSize: 9.5, bold: true, color: C.amber, valign: 'top' });
  }

  function coverageSlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Cakupan & Pemerataan Monitoring', 'Frekuensi kunjungan dan kesenjangan pengawasan', page);
    kpi(s, pptx, .6, 1.25, 2.55, 1.15, fmtPct(d.coverage), 'Cakupan unit', C.green, `${fmt(d.monitored)} dari ${fmt(d.total)} unit`);
    kpi(s, pptx, 3.32, 1.25, 2.55, 1.15, fmt(d.visits), 'Total kunjungan', C.blue, 'Dalam periode laporan');
    kpi(s, pptx, 6.04, 1.25, 2.55, 1.15, fmt(d.unmonitored), 'Belum dimonitor', C.amber, 'Target kunjungan');
    kpi(s, pptx, 8.76, 1.25, 2.55, 1.15, fmt(d.critical), 'Unit kritis', C.rose, 'Tindak lanjut segera');
    panel(s, pptx, .6, 2.72, 6.1, 3.95, 'Distribusi Frekuensi Kunjungan');
    barList(s, pptx, Object.entries(d.freq), .9, 3.32, 5.45, 2.82, C.blue);
    panel(s, pptx, 6.95, 2.72, 5.82, 3.95, 'Cakupan per Wilayah', 'F8FAFC');
    const regionEntries = REGION_ORDER.filter(r => d.byRegion[r].units).map(r => [r.replace('Kab. ', ''), d.byRegion[r].coverage]);
    barList(s, pptx, regionEntries, 7.25, 3.35, 5.15, 1.85, C.green, 100, '%');
    const msg = d.unmonitored
      ? `${fmt(d.unmonitored)} unit belum pernah dimonitor pada periode ini. Pemerataan jadwal lebih penting daripada mengulang kunjungan pada unit yang sama.`
      : 'Seluruh unit telah memiliki kunjungan pada periode ini. Fokus berikutnya adalah kualitas tindak lanjut.';
    addText(s, msg, 7.28, 5.48, 5.0, .82, { fontSize: 10.5, color: C.slate, bold: true, valign: 'top' });
  }

  function sppgBeneficiarySlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Penerima Manfaat & Jaringan Sekolah', 'Rekap monitoring terakhir per SPPG — tidak menggandakan kunjungan', page);
    kpi(s, pptx, .6, 1.23, 2.55, 1.08, fmt(d.benefTotal), 'Penerima/porsi per hari', C.blue, `${fmt(d.completeness.benef)}/${fmt(d.total)} SPPG mengisi`);
    kpi(s, pptx, 3.32, 1.23, 2.55, 1.08, fmt(d.benef.siswa), 'Penerima siswa', C.cyan, `${fmtPct(pct(d.benef.siswa, d.benefTotal))} dari isian`);
    kpi(s, pptx, 6.04, 1.23, 2.55, 1.08, fmt(d.schoolTotal), 'Sekolah dilayani', C.green, `${fmt(d.completeness.schools)}/${fmt(d.total)} SPPG mengisi`);
    kpi(s, pptx, 8.76, 1.23, 2.55, 1.08, fmt(d.capTotal), 'Kapasitas Master Unit', C.amber, 'Pembanding data master');
    panel(s, pptx, .6, 2.62, 6.05, 3.98, 'Penerima menurut Kelompok');
    barList(s, pptx, Object.keys(d.benef).map(k => [d.benefLabels[k], d.benef[k]]), .9, 3.18, 5.42, 2.9, C.blue);
    panel(s, pptx, 6.9, 2.62, 5.87, 3.98, 'Jumlah Sekolah menurut Jenjang', 'F8FAFC');
    barList(s, pptx, Object.keys(d.schools).map(k => [d.schoolLabels[k], d.schools[k]]), 7.2, 3.18, 5.2, 2.62, C.green);
    addText(s, 'Catatan: data jenjang menunjukkan jumlah sekolah, bukan jumlah siswa per jenjang.', 7.2, 5.98, 5.08, .35, { fontSize: 9, italic: true, color: C.amber });
  }

  function sppgOpsSlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Operasional, Distribusi & Tenaga Kerja', 'Beban layanan, waktu tempuh, dan dukungan SDM', page);
    const porsiPerCook = sum(d.cooks) ? d.benefTotal / sum(d.cooks) : 0;
    kpi(s, pptx, .6, 1.23, 2.55, 1.12, fmt(avg(d.days), 1), 'Rata-rata hari salur', C.blue, 'Hari pada bulan lalu');
    kpi(s, pptx, 3.32, 1.23, 2.55, 1.12, fmt(d.dist['> 30 menit']), 'Distribusi >30 menit', C.amber, 'Perlu evaluasi rute');
    kpi(s, pptx, 6.04, 1.23, 2.55, 1.12, fmt(sum(d.workers)), 'Pekerja terlapor', C.green, `${fmt(d.workers.length)} SPPG mengisi`);
    kpi(s, pptx, 8.76, 1.23, 2.55, 1.12, fmt(porsiPerCook), 'Porsi per juru masak', C.red, 'Indikator beban agregat');
    panel(s, pptx, .6, 2.7, 5.92, 3.9, 'Waktu Tempuh Distribusi');
    barList(s, pptx, Object.entries(d.dist), .9, 3.3, 5.28, 2.55, C.amber);
    panel(s, pptx, 6.77, 2.7, 6.0, 3.9, 'Cakupan BPJS Ketenagakerjaan', 'F8FAFC');
    barList(s, pptx, Object.entries(d.bpjs), 7.08, 3.3, 5.3, 2.55, C.green);
    addText(s, d.dist['> 30 menit'] ? `Fokus: evaluasi rute, waktu pemorsian, dan armada pada ${fmt(d.dist['> 30 menit'])} unit.` : 'Tidak ada laporan distribusi lebih dari 30 menit pada data terisi.', 7.1, 5.95, 5.1, .38, { fontSize: 9.5, bold: true, color: d.dist['> 30 menit'] ? C.amber : C.green });
  }

  function sppgSafetySlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Keamanan Pangan & Pengelolaan Lingkungan', 'Sinyal risiko berdasarkan jawaban monitoring terakhir', page);
    const cards = [
      [d.safety.noExternal, 'Tanpa pengawasan eksternal', C.amber], [d.safety.noSample, 'Tanpa uji sampel', C.rose],
      [d.safety.digestive, 'Gangguan pencernaan', C.red], [d.safety.riskyWaste, 'Limbah berisiko', C.purple]
    ];
    cards.forEach((x, i) => kpi(s, pptx, .65 + i * 3.05, 1.32, 2.75, 1.22, fmt(x[0]), x[1], x[2], `${fmt(d.safety.filled)}/${fmt(d.total)} SPPG berdata`));
    panel(s, pptx, .65, 2.9, 7.0, 3.65, 'Interpretasi Risiko', 'FFF7ED');
    const risk = [];
    if (d.safety.digestive) risk.push(`${fmt(d.safety.digestive)} unit melaporkan kejadian gangguan pencernaan dan memerlukan verifikasi segera.`);
    if (d.safety.noSample) risk.push(`${fmt(d.safety.noSample)} unit belum melakukan uji sampel sebelum distribusi.`);
    if (d.safety.riskyWaste) risk.push(`${fmt(d.safety.riskyWaste)} unit melaporkan pengelolaan limbah yang berisiko.`);
    if (d.safety.noExternal) risk.push(`${fmt(d.safety.noExternal)} unit belum memiliki pengawasan kelaikan eksternal.`);
    if (!risk.length) risk.push(d.safety.filled ? 'Tidak ada sinyal risiko utama pada jawaban yang terisi.' : 'Data keamanan pangan belum tersedia untuk periode ini.');
    bullets(s, risk, .98, 3.52, 6.3, 2.55, C.slate);
    panel(s, pptx, 7.9, 2.9, 4.87, 3.65, 'Arahan Manajemen', 'F8FAFC');
    bullets(s, ['Verifikasi unit dengan kejadian kesehatan.', 'Pastikan uji sampel dilakukan sebelum distribusi.', 'Perkuat pengawasan Dinkes/Puskesmas.', 'Tindak lanjuti pengelolaan limbah berisiko.'], 8.22, 3.5, 4.22, 2.45, C.slate);
  }

  function sppgNakerSlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Profil & Kondisi Tenaga Kerja', 'Rekap respons Naker pada periode laporan', page);
    kpi(s, pptx, .6, 1.25, 2.55, 1.1, fmt(d.naker.responses), 'Respons Naker', C.purple, 'Bukan otomatis orang unik');
    kpi(s, pptx, 3.32, 1.25, 2.55, 1.1, fmt(d.naker.avgDays, 1), 'Hari kerja/minggu', C.blue, 'Rata-rata respons');
    kpi(s, pptx, 6.04, 1.25, 2.55, 1.1, fmt(d.naker.avgHours, 1), 'Jam kerja/hari', C.cyan, 'Rata-rata respons');
    kpi(s, pptx, 8.76, 1.25, 2.55, 1.1, fmt(d.naker.avgNewWage), 'Rata-rata upah', C.green, 'Rp/bulan');
    panel(s, pptx, .6, 2.72, 5.95, 3.9, 'Posisi Responden');
    barList(s, pptx, topEntries(d.naker.positions, 6), .9, 3.3, 5.3, 2.62, C.purple);
    panel(s, pptx, 6.8, 2.72, 5.97, 3.9, 'Kendala Pekerjaan Dominan', 'F8FAFC');
    barList(s, pptx, topEntries(d.naker.constraints, 6), 7.1, 3.3, 5.28, 2.62, C.amber);
    if (!d.naker.responses) addText(s, 'Belum ada respons Form Naker pada periode ini.', 7.1, 6.05, 5.15, .3, { fontSize: 10, bold: true, color: C.amber });
  }

  function sppgFinanceSlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Supplier, Belanja Lokal & Kendala Keuangan', 'Rekap nilai isian dalam Rp juta pada monitoring terbaru', page);
    kpi(s, pptx, .6, 1.24, 2.55, 1.12, fmt(sum(d.supplierCounts)), 'Supplier terlapor', C.blue, `${fmt(d.supplierCounts.length)} SPPG mengisi`);
    kpi(s, pptx, 3.32, 1.24, 2.55, 1.12, fmt(sum(Object.values(d.finance.supplierSpend)), 1), 'Belanja pemasok', C.green, 'Rp juta — data terisi');
    kpi(s, pptx, 6.04, 1.24, 2.55, 1.12, fmt(d.finance.originSpend['Dalam kota'], 1), 'Belanja dalam kota', C.cyan, 'Rp juta');
    kpi(s, pptx, 8.76, 1.24, 2.55, 1.12, fmt(d.finance.filled), 'SPPG berdata keuangan', C.amber, `${fmtPct(pct(d.finance.filled,d.total))} cakupan`);
    panel(s, pptx, .6, 2.72, 6.0, 3.92, 'Pengeluaran Bahan Baku per Kelompok Supplier');
    barList(s, pptx, Object.entries(d.finance.supplierSpend), .9, 3.3, 5.35, 2.7, C.green, null, ' jt');
    panel(s, pptx, 6.84, 2.72, 5.93, 3.92, 'Kendala Dominan', 'F8FAFC');
    const obs=topEntries(d.finance.obstacles,6);
    if(obs.length)barList(s,pptx,obs,7.14,3.3,5.28,2.55,C.amber);
    else addText(s,'Belum ada data kendala keuangan/distribusi yang terisi.',7.16,3.62,5.1,.5,{fontSize:11,bold:true,color:C.amber,align:'center'});
    const local=sum(Object.values(d.finance.originSpend));
    if(local)addText(s,`Proporsi belanja dalam kota ${fmtPct(pct(d.finance.originSpend['Dalam kota'],local))}; gunakan sebagai indikator awal penguatan ekonomi lokal.`,7.16,5.98,5.0,.4,{fontSize:9.2,bold:true,color:C.slate});
  }

  function kdmpStatusSlide(pptx,d,page){
    const s=pptx.addSlide();addHeader(s,pptx,'Status Operasional & Kategori Kinerja','Kondisi unit dan distribusi kategori hasil monitoring',page);
    const sts=Object.entries(d.status).map(([k,v])=>[k.charAt(0).toUpperCase()+k.slice(1),v]);
    panel(s,pptx,.6,1.25,5.92,5.45,'Status Master Unit');barList(s,pptx,sts,.92,1.95,5.25,3.15,C.red);
    panel(s,pptx,6.78,1.25,5.99,5.45,'Kategori Hasil KDMP','F8FAFC');
    const cats=Object.entries(d.scoreCategories);
    if(cats.length)barList(s,pptx,cats,7.1,1.95,5.3,2.6,C.purple);
    else addText(s,'Kategori hasil belum tersedia.',7.15,2.15,5.0,.45,{fontSize:12,bold:true,color:C.amber,align:'center'});
    addText(s,`Skor rata-rata: ${d.snapshots.length?fmt(d.avgScore,2):'—'} / 4,00. Cakupan form: ${fmt(d.snapshots.length)}/${fmt(d.total)} KDMP.`,7.12,5.1,5.05,.6,{fontSize:11,bold:true,color:C.slate,valign:'top'});
  }

  function kdmpOpenSlide(pptx,d,page){
    const s=pptx.addSlide();addHeader(s,pptx,'Suara Lapangan, Produk & Peluang Pengembangan','Ringkasan respons terbuka pada monitoring KDMP terbaru',page);
    const boxes=[['Kelebihan',d.open.strengths,C.green],['Kendala',d.open.constraints,C.amber],['Saran',d.open.suggestions,C.blue],['Produk Unggulan',d.open.products,C.purple]];
    boxes.forEach((b,i)=>{const x=.6+(i%2)*6.15,y=1.28+Math.floor(i/2)*2.72;panel(s,pptx,x,y,5.9,2.42,b[0],i%2?'F8FAFC':'FFFFFF');const vals=b[1].slice(0,4).map(t=>safe(t).slice(0,150));bullets(s,vals.length?vals:['Belum ada respons terisi.'],x+.28,y+.58,5.3,1.55,C.slate);addShape(s,pptx,pptx.ShapeType.rect,x,y,.07,2.42,{fill:{color:b[2]},line:{transparency:100}});});
  }

  function findingsSlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Temuan Utama & Unit Prioritas', 'Klasifikasi temuan lapangan dan sasaran tindak lanjut', page);
    panel(s, pptx, .6, 1.25, 6.1, 5.48, 'Kategori Temuan Dominan');
    barList(s, pptx, topEntries(d.findings.counts, 7), .9, 1.9, 5.45, 3.9, C.red);
    const totalFind = sum(Object.values(d.findings.counts));
    addText(s, `${fmt(totalFind)} sinyal kategori ditemukan dari teks temuan, rekomendasi, dan jawaban terbuka. Satu catatan dapat masuk lebih dari satu kategori.`, .92, 5.98, 5.35, .42, { fontSize: 8.5, italic: true, color: C.muted });
    panel(s, pptx, 6.95, 1.25, 5.82, 5.48, 'Daftar Prioritas Tertinggi', 'FFF7ED');
    const pri = d.priority.slice(0, 8);
    if (!pri.length) addText(s, 'Belum ada unit prioritas berdasarkan aturan laporan.', 7.28, 2.0, 5.1, .4, { fontSize: 11, color: C.green, bold: true });
    pri.forEach((p, i) => {
      const yy = 1.87 + i * .55;
      addShape(s, pptx, pptx.ShapeType.ellipse, 7.25, yy + .04, .28, .28, { fill: { color: i < 3 ? C.rose : C.amber }, line: { transparency: 100 } });
      addText(s, String(i + 1), 7.25, yy + .04, .28, .28, { fontSize: 8, color: C.white, bold: true, align: 'center' });
      addText(s, p.unit.nama, 7.68, yy, 3.45, .24, { fontSize: 9.2, bold: true, color: C.navy });
      addText(s, p.reasons.join(' · '), 7.68, yy + .25, 4.62, .2, { fontSize: 7.8, color: C.slate });
    });
  }

  function recommendationsSlide(pptx, d, cfg, page, edited) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Rekomendasi & Rencana Tindak Lanjut', 'Saran sistem harus diverifikasi dan disetujui penyusun laporan', page);
    panel(s, pptx, .6, 1.25, 7.35, 5.5, 'Rekomendasi Manajerial', 'F8FAFC');
    const list = edited ? edited.split(/\n+/).map(x => x.replace(/^\s*\d+[.)]\s*/, '')).filter(Boolean) : d.recommendations;
    bullets(s, list, .95, 1.92, 6.68, 4.45, C.slate);
    panel(s, pptx, 8.2, 1.25, 4.57, 5.5, 'Kerangka Aksi', 'FFF7ED');
    const actions = [['P1', 'Tindak lanjut risiko kritis', '0–14 hari', C.rose], ['P2', 'Perluasan monitoring & koreksi', '15–30 hari', C.amber], ['P3', 'Perbaikan data dan sistem', '31–90 hari', C.blue]];
    actions.forEach((x, i) => {
      const yy = 1.96 + i * 1.28;
      addShape(s, pptx, pptx.ShapeType.roundRect, 8.52, yy, .58, .58, { fill: { color: x[3] }, line: { transparency: 100 }, rectRadius: .06 });
      addText(s, x[0], 8.52, yy, .58, .58, { fontSize: 12, color: C.white, bold: true, align: 'center' });
      addText(s, x[1], 9.28, yy, 2.9, .3, { fontSize: 10.5, bold: true, color: C.navy });
      addText(s, 'Target: ' + x[2], 9.28, yy + .38, 2.7, .22, { fontSize: 9, color: C.slate });
    });
    addText(s, `Penyusun: ${cfg.author}`, 8.52, 6.0, 3.7, .25, { fontSize: 9.5, bold: true, color: C.slate });
  }

  function qualitySlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Kelengkapan Data & Catatan Interpretasi', 'Transparansi denominator agar keputusan tidak melampaui kualitas data', page);
    const entries = d.kind === 'sppg' ? [
      ['Form SPPG', d.completeness.sppgForms], ['Data penerima', d.completeness.benef], ['Data sekolah', d.completeness.schools], ['Keamanan pangan', d.completeness.safety]
    ] : [['Form KDMP', d.completeness.forms], ['Skor penilaian', d.completeness.scored], ['Kepatuhan', d.completeness.compliance]];
    panel(s, pptx, .6, 1.25, 6.15, 5.5, 'Cakupan Isian terhadap Unit');
    barList(s, pptx, entries.map(x => [x[0], pct(x[1], d.total)]), .95, 2.0, 5.4, 2.8, C.cyan, 100, '%');
    addText(s, `${fmt(d.total)} unit menjadi denominator laporan. Angka agregat hanya mewakili unit yang mengisi masing-masing bagian.`, .95, 5.22, 5.35, .6, { fontSize: 10, bold: true, color: C.slate, valign: 'top' });
    panel(s, pptx, 7.0, 1.25, 5.77, 5.5, 'Batasan & Cara Membaca', 'F8FAFC');
    const notes = d.kind === 'sppg' ? [
      'Data snapshot memakai monitoring SPPG terakhir per unit.', 'Frekuensi memakai seluruh kunjungan dalam periode.',
      'Data Naker dihitung sebagai respons, bukan otomatis orang unik.', 'Jumlah jenjang adalah jumlah sekolah, bukan siswa per jenjang.',
      'Narasi tren hanya sah jika tersedia lebih dari satu periode.', 'Rekomendasi otomatis perlu persetujuan penyusun.'
    ] : [
      'Skor memakai monitoring KDMP terakhir per unit.', 'Rata-rata hanya memakai butir yang terjawab.',
      'Kepatuhan dihitung dari jawaban Ya/Tidak yang tersedia.', 'Respons terbuka diklasifikasikan dengan kata kunci.',
      'Rekomendasi otomatis perlu persetujuan penyusun.'
    ];
    bullets(s, notes, 7.35, 1.95, 5.05, 4.3, C.slate);
  }

  function kdmpDimensionSlide(pptx, d, page) {
    const s = pptx.addSlide(); addHeader(s, pptx, 'Skor Delapan Dimensi KDMP', 'Rata-rata butir terjawab pada monitoring terbaru', page);
    kpi(s, pptx, .65, 1.25, 2.75, 1.15, d.snapshots.length ? fmt(d.avgScore, 2) : '—', 'Skor rata-rata / 4', C.red, `${fmt(d.snapshots.length)}/${fmt(d.total)} KDMP berform`);
    kpi(s, pptx, 3.6, 1.25, 2.75, 1.15, fmt(d.compliance.yes), 'Jawaban patuh (Ya)', C.green, `${fmt(d.compliance.total)} jawaban terisi`);
    kpi(s, pptx, 6.55, 1.25, 2.75, 1.15, fmtPct(pct(d.compliance.yes, d.compliance.total)), 'Tingkat kepatuhan', C.cyan, 'Dari jawaban tersedia');
    kpi(s, pptx, 9.5, 1.25, 2.75, 1.15, fmt(d.unmonitored), 'Belum dimonitor', C.amber, 'Target kunjungan');
    panel(s, pptx, .65, 2.72, 12.1, 3.95, 'Perbandingan Dimensi');
    const entries = Object.entries(d.dims).sort((a, b) => b[1] - a[1]);
    if (entries.length) barList(s, pptx, entries, 1.0, 3.3, 11.3, 2.86, C.red, 4, ' skor');
    else addText(s, 'Belum tersedia skor dimensi KDMP pada periode ini.', 1.0, 3.55, 10.8, .5, { fontSize: 14, bold: true, color: C.amber, align: 'center' });
  }

  async function buildPpt(cfg, d, geo, editedNarrative, editedRecommendations) {
    if (typeof global.PptxGenJS !== 'function') throw new Error('Library PowerPoint belum termuat. Periksa koneksi internet lalu refresh.');
    const pptx = new global.PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; pptx.author = cfg.author; pptx.company = 'DARMA'; pptx.subject = 'Laporan Manajemen'; pptx.title = cfg.title; pptx.lang = 'id-ID';
    pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos', lang: 'id-ID' };
    coverSlide(pptx, d, cfg);
    let p = 2;
    executiveSlide(pptx, d, cfg, p++, editedNarrative);
    mapSlide(pptx, d, cfg, geo, p++);
    coverageSlide(pptx, d, p++);
    if (d.kind === 'sppg') {
      sppgBeneficiarySlide(pptx, d, p++);
      if(cfg.mode==='lengkap'){
        sppgOpsSlide(pptx, d, p++);
        sppgSafetySlide(pptx, d, p++);
        sppgFinanceSlide(pptx, d, p++);
        sppgNakerSlide(pptx, d, p++);
      }
      findingsSlide(pptx, d, p++);
      recommendationsSlide(pptx, d, cfg, p++, editedRecommendations);
      qualitySlide(pptx, d, p++);
    } else {
      kdmpDimensionSlide(pptx, d, p++);
      if(cfg.mode==='lengkap'){
        kdmpStatusSlide(pptx,d,p++);
        kdmpOpenSlide(pptx,d,p++);
      }
      findingsSlide(pptx, d, p++);
      recommendationsSlide(pptx, d, cfg, p++, editedRecommendations);
      qualitySlide(pptx, d, p++);
    }
    return pptx;
  }

  async function generateManagementPpt() {
    const btn = document.getElementById('rptGenerateBtn');
    try {
      if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Menyusun PowerPoint...'; }
      if (!reportCache) previewManagementReport(false);
      if (!reportCache) throw new Error('Ringkasan laporan belum tersedia');
      const cfg = getConfig(), data = analyze(cfg); reportCache = { cfg, data };
      const geo = await getGeoJSON();
      const nar = document.getElementById('rptNarrative'); const rec = document.getElementById('rptRecommendations');
      const pptx = await buildPpt(cfg, data, geo, nar ? nar.value.trim() : '', rec ? rec.value.trim() : '');
      const period = cfg.start ? cfg.start.slice(0, 7) : new Date().toISOString().slice(0, 7);
      const fileName = `DARMA_Laporan_${cfg.type === 'sppg' ? 'SPPG_NAKER' : 'KDMP'}_${period}.pptx`;
      await pptx.writeFile({ fileName });
      if (typeof toast === 'function') toast('📊 PowerPoint manajemen berhasil dibuat');
      const st = document.getElementById('rptStatus'); if (st) st.innerHTML = `<i class="fas fa-check-circle" style="color:#16a34a"></i> ${html(fileName)} berhasil diunduh.`;
    } catch (e) {
      console.error('PPT generation error:', e);
      if (typeof toast === 'function') toast('Gagal membuat PPT: ' + e.message, 'e');
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-file-powerpoint"></i> Unduh PowerPoint'; }
    }
  }

  global.initReportPanel = initReportPanel;
  global.selectReportKind = selectReportKind;
  global.markReportStale = markReportStale;
  global.previewManagementReport = previewManagementReport;
  global.generateManagementPpt = generateManagementPpt;
  global.DarmaManagementReport = { analyze, buildPpt, mapSvg };
})(window);
