/* ============================================================
   SIMON-MBG — Cloudflare Worker + D1 (single source of truth)
   API:  /api/ping | /api/login | /api/state | /api/users | /api/units | /api/monitoring
   - Sesi berbasis token (tabel sessions), password tidak dikirim ke klien via /api/state
   - Hanya admin yang boleh mengelola users & replaceAll
   ============================================================ */
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = {
      'Access-Control-Allow-Origin': request.headers.get('Origin') || '*',
      'Access-Control-Allow-Methods': 'GET,POST,PUT,DELETE,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization',
      'Access-Control-Max-Age': '86400',
      'Content-Type': 'application/json'
    };
    if (request.method === 'OPTIONS') return new Response(null, { headers: cors });

    const path = url.pathname;
    if (!path.startsWith('/api/')) return json({ error: 'Not found' }, 404, cors);

    try {
      const r = path.slice(5); // hapus "/api/"

      // --- publik ---
      if (r === 'ping') return json({ ok: true, name: 'SIMON-MBG API' }, cors);
      if (r === 'login' && request.method === 'POST') return await login(request, env, cors);

      // --- butuh login ---
      const me = await authorize(request, env);
      if (!me) return json({ error: 'Unauthorized' }, 401, cors);

      if (r === 'state' && request.method === 'GET') {
        const [units, monitoring] = await Promise.all([allUnits(env), allMonitoring(env)]);
        return json({ user: pub(me), units, monitoring }, cors);
      }
      if (r === 'state' && request.method === 'PUT') {
        if (me.role !== 'admin') return json({ error: 'Forbidden' }, 403, cors);
        return await replaceAll(request, env, cors);
      }

      // --- users (admin) ---
      if (r === 'users' && request.method === 'GET') {
        if (me.role !== 'admin') return json({ error: 'Forbidden' }, 403, cors);
        return json({ users: await allUsers(env) }, cors);
      }
      if (r === 'users' && request.method === 'POST') {
        if (me.role !== 'admin') return json({ error: 'Forbidden' }, 403, cors);
        return await upsertUser(request, env, cors);
      }
      if (r.startsWith('users/') && request.method === 'DELETE') {
        if (me.role !== 'admin') return json({ error: 'Forbidden' }, 403, cors);
        return await deleteUser(env, decodeURIComponent(r.slice(6)), me, cors);
      }

      // --- units ---
      if (r === 'units' && request.method === 'POST') return await upsertUnit(request, env, cors);
      if (r.startsWith('units/') && request.method === 'DELETE') return await deleteUnit(env, decodeURIComponent(r.slice(6)), cors);

      // --- monitoring ---
      if (r === 'monitoring' && request.method === 'POST') return await upsertMonitoring(request, env, cors);
      if (r === 'monitoring' && request.method === 'DELETE') {
        if (me.role !== 'admin') return json({ error: 'Forbidden' }, 403, cors);
        await env.DB.prepare('DELETE FROM monitoring').run();
        return json({ ok: true }, cors);
      }
      if (r.startsWith('monitoring/') && request.method === 'DELETE') return await deleteMonitoring(env, decodeURIComponent(r.slice(11)), cors);

      return json({ error: 'Not found' }, 404, cors);
    } catch (e) {
      return json({ error: 'Server error: ' + (e && e.message || e) }, 500, cors);
    }
  }
};

/* ---------- helpers ---------- */
function json(obj, extra, third) {
  let st = 200, cors;
  if (typeof extra === 'number') { st = extra; cors = third; } else { cors = extra; }
  return new Response(JSON.stringify(obj), { status: st, headers: Object.assign({}, cors, { 'Content-Type': 'application/json' }) });
}
function pub(u) { return { id: u.id, nama: u.nama, username: u.username, role: u.role }; }
function token() {
  const a = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let s = ''; for (let i = 0; i < 32; i++) s += a[Math.floor(Math.random() * a.length)]; return s;
}
async function authorize(request, env) {
  const h = request.headers.get('Authorization') || '';
  const m = h.match(/^Bearer\s+(.+)$/);
  if (!m) return null;
  const row = await env.DB.prepare('SELECT username FROM sessions WHERE token=?').bind(m[1]).first();
  if (!row) return null;
  const u = await env.DB.prepare('SELECT * FROM users WHERE username=?').bind(row.username).first();
  return u || null;
}
async function login(request, env, cors) {
  let b; try { b = await request.json(); } catch (e) { return json({ error: 'Bad body' }, 400, cors); }
  const u = await env.DB.prepare('SELECT * FROM users WHERE username=?').bind((b.username || '').trim()).first();
  if (!u || u.password !== (b.password || '')) return json({ error: 'Invalid credentials' }, 401, cors);
  const t = token();
  await env.DB.prepare('INSERT INTO sessions(token, username, created_at) VALUES(?,?,?)').bind(t, u.username, Date.now()).run();
  return json({ token: t, user: pub(u) }, cors);
}
async function allUsers(env) { const { results } = await env.DB.prepare('SELECT id, nama, username, password, role FROM users ORDER BY nama').all(); return results; }
function unitRow(r) {
  return { id: r.id, jenis: r.jenis, nama: r.nama, ref: r.ref, status: r.status, kab: r.kab, kec: r.kec, desa: r.desa, alamat: r.alamat, lat: r.lat, lng: r.lng, pic: r.pic, telp: r.telp, note: r.note, yayasan: r.yayasan, kapasitas: r.kapasitas, sekolah: r.sekolah, slhs: r.slhs, mulai: r.mulai, anggota: r.anggota, peran: r.peran, usaha: r.usaha };
}
async function allUnits(env) { const { results } = await env.DB.prepare('SELECT * FROM units ORDER BY jenis, nama').all(); return results.map(unitRow); }
function monRow(r) { let form = null; try { form = r.form_json ? JSON.parse(r.form_json) : null; } catch (e) {} return { id: r.id, unitId: r.unit_id, tgl: r.tgl, petugas: r.petugas, jenis: r.jenis, formType: r.form_type, hasil: r.hasil, form, kebersihan: r.kebersihan, gizi: r.gizi, distribusi: r.distribusi, dok: r.dok, temuan: r.temuan, rekom: r.rekom }; }
async function allMonitoring(env) { const { results } = await env.DB.prepare('SELECT * FROM monitoring ORDER BY tgl DESC').all(); return results.map(monRow); }

const U_COLS = 'id,jenis,nama,ref,status,kab,kec,desa,alamat,lat,lng,pic,telp,note,yayasan,kapasitas,sekolah,slhs,mulai,anggota,peran,usaha,updated_at';
const U_Q = 'INSERT OR REPLACE INTO units(id,jenis,nama,ref,status,kab,kec,desa,alamat,lat,lng,pic,telp,note,yayasan,kapasitas,sekolah,slhs,mulai,anggota,peran,usaha,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)';
const M_COLS = 'id,unit_id,tgl,petugas,jenis,form_type,hasil,form_json,kebersihan,gizi,distribusi,dok,temuan,rekom,updated_at';
const M_Q = 'INSERT OR REPLACE INTO monitoring(id,unit_id,tgl,petugas,jenis,form_type,hasil,form_json,kebersihan,gizi,distribusi,dok,temuan,rekom,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)';
function bindUnit(u) {
  return [u.id||'', u.jenis||'', u.nama||'', u.ref||'', u.status||'aktif', u.kab||'', u.kec||'', u.desa||'', u.alamat||'', u.lat!=null?Number(u.lat):0, u.lng!=null?Number(u.lng):0, u.pic||'', u.telp||'', u.note||'', u.yayasan||'', u.kapasitas!=null?Number(u.kapasitas):0, u.sekolah!=null?Number(u.sekolah):0, u.slhs||'', u.mulai||'', u.anggota!=null?Number(u.anggota):0, u.peran||'', u.usaha||'', Date.now()];
}
function bindMon(m) {
  return [m.id||'', m.unitId||'', m.tgl||'', m.petugas||'', m.jenis||'', m.formType||'', m.hasil||'', JSON.stringify(m.form||null), m.kebersihan||'', m.gizi||'', m.distribusi||'', m.dok||'', m.temuan||'', m.rekom||'', Date.now()];
}

async function upsertUser(request, env, cors) {
  let u; try { u = await request.json(); } catch (e) { return json({ error: 'Bad body' }, 400, cors); }
  if (!u.id || !u.username || !u.password || !u.role) return json({ error: 'Field tidak lengkap' }, 400, cors);
  await env.DB.prepare('INSERT OR REPLACE INTO users(id,nama,username,password,role) VALUES(?,?,?,?,?)').bind(u.id, u.nama, u.username, u.password, u.role).run();
  return json({ ok: true }, cors);
}
async function deleteUser(env, id, me, cors) {
  if (id === me.id) return json({ error: 'Tidak bisa menghapus akun sendiri' }, 400, cors);
  const target = await env.DB.prepare('SELECT role FROM users WHERE id=?').bind(id).first();
  if (target && target.role === 'admin') {
    const ac = await env.DB.prepare('SELECT COUNT(*) c FROM users WHERE role=?').bind('admin').first();
    if ((ac.c || 0) <= 1) return json({ error: 'Minimal harus ada 1 admin' }, 400, cors);
  }
  await env.DB.prepare('DELETE FROM users WHERE id=?').bind(id).run();
  return json({ ok: true }, cors);
}
async function upsertUnit(request, env, cors) {
  let u; try { u = await request.json(); } catch (e) { return json({ error: 'Bad body' }, 400, cors); }
  await env.DB.prepare(U_Q).bind(...bindUnit(u)).run();
  return json({ ok: true }, cors);
}
async function deleteUnit(env, id, cors) {
  await env.DB.batch([env.DB.prepare('DELETE FROM monitoring WHERE unit_id=?').bind(id), env.DB.prepare('DELETE FROM units WHERE id=?').bind(id)]);
  return json({ ok: true }, cors);
}
async function upsertMonitoring(request, env, cors) {
  let m; try { m = await request.json(); } catch (e) { return json({ error: 'Bad body' }, 400, cors); }
  await env.DB.prepare(M_Q).bind(...bindMon(m)).run();
  return json({ ok: true }, cors);
}
async function deleteMonitoring(env, id, cors) {
  await env.DB.prepare('DELETE FROM monitoring WHERE id=?').bind(id).run();
  return json({ ok: true }, cors);
}
async function replaceAll(request, env, cors) {
  let b; try { b = await request.json(); } catch (e) { return json({ error: 'Bad body' }, 400, cors); }
  const stmts = [];
  if (Array.isArray(b.users)) {
    stmts.push(env.DB.prepare('DELETE FROM users'));
    b.users.forEach(u => stmts.push(env.DB.prepare('INSERT OR REPLACE INTO users(id,nama,username,password,role) VALUES(?,?,?,?,?)').bind(u.id, u.nama, u.username, u.password, u.role)));
  }
  if (Array.isArray(b.units)) {
    stmts.push(env.DB.prepare('DELETE FROM units'));
    b.units.forEach(u => stmts.push(env.DB.prepare(U_Q).bind(...bindUnit(u))));
  }
  if (Array.isArray(b.monitoring)) {
    stmts.push(env.DB.prepare('DELETE FROM monitoring'));
    b.monitoring.forEach(m => stmts.push(env.DB.prepare(M_Q).bind(...bindMon(m))));
  }
  if (stmts.length) await env.DB.batch(stmts);
  return json({ ok: true }, cors);
}
