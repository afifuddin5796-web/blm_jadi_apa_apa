// Uji asap (smoke test): memuat index.html di jsdom lalu menguji alur utama.
const { JSDOM, ResourceLoader, VirtualConsole } = require('jsdom');
const fs = require('fs');
const path = require('path');
const assert = require('assert');

const root = path.join(__dirname, '..');
class Lokal extends ResourceLoader {
  fetch(url) {
    const u = new URL(url);
    if (u.hostname !== 'localhost') return Promise.resolve(Buffer.from(''));
    const f = path.join(root, decodeURIComponent(u.pathname).replace(/^\//, '') || 'index.html');
    return Promise.resolve(fs.existsSync(f) ? fs.readFileSync(f) : Buffer.from(''));
  }
}
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const galat = [];
  const vc = new VirtualConsole();
  vc.on('jsdomError', e => galat.push(e.message));
  const dom = new JSDOM(fs.readFileSync(path.join(root, 'index.html'), 'utf8'), {
    url: 'http://localhost/', runScripts: 'dangerously', resources: new Lokal(), pretendToBeVisual: true, virtualConsole: vc,
  });
  const w = dom.window, d = w.document;
  w.scrollTo = () => {};
  await new Promise(r => (d.readyState === 'complete' ? r() : w.addEventListener('load', r)));
  await sleep(150);

  assert.deepStrictEqual(galat, [], 'Tidak boleh ada galat saat memuat: ' + galat.join(' | '));
  assert(w.XLSX && w.JSZip, 'Library lokal (vendor/) harus termuat');
  assert.strictEqual(w.eval('PT.length'), 10, 'Data contoh: 10 petak');

  // 0) Tab Tanya: jawaban untuk juru (data contoh)
  w.eval("show('tny')");
  const tanya = (q, v) => { const i = d.querySelector(`[data-q="${q}"]`); i.value = v; i.dispatchEvent(new w.Event('input', { bubbles: true })); };
  const nama = w.eval('PT[0].n'), di = w.eval('PT[0].di');
  tanya(1, nama); assert(/berada di DI/.test(d.querySelector('#a1').textContent), 'Tanya 1: debit');
  tanya(2, di); assert(/Petak tersier/.test(d.querySelector('#a2').textContent), 'Tanya 2: daerah yang diairi');
  tanya(3, nama); assert(/Kecamatan/.test(d.querySelector('#a3').textContent), 'Tanya 3: lokasi');
  tanya(5, di); d.querySelector('[data-act="all5"]').click();
  assert(/petak terpilih/.test(d.querySelector('#s5').textContent), 'Tanya 5: kebutuhan petak terpilih');
  tanya(6, ''); assert(d.querySelectorAll('#a6 tbody tr').length > 0, 'Tanya 6: juru dan DI');

  // 1) KML
  const kml = fs.readFileSync(path.join(root, 'data/contoh/contoh.kml'), 'utf8');
  w.__kml = kml;
  w.eval('kui(kparse(__kml))');
  d.querySelector('#kgo').click();
  assert.strictEqual(w.eval('PT.length'), 2, 'KML: 2 petak');
  assert.strictEqual(w.eval('D.bn.length'), 1, 'KML: 1 bendung');
  assert.strictEqual(w.eval('D.sal.length'), 1, 'KML: 1 saluran');
  assert(d.querySelectorAll('#map g.lt').length === 2, 'Label petak tampil');

  // 2) Excel (formulir sederhana)
  const buf = fs.readFileSync(path.join(root, 'data/contoh/laporan-tanaman-contoh.xlsx'));
  // ArrayBuffer harus dibuat di dalam 'window' jsdom (beda realm dengan Node)
  w.__f = { name: 'x.xlsx', arrayBuffer: async () => new w.Uint8Array(buf).buffer };
  await w.eval('xlOpen(__f)');
  d.querySelector('#xm').value = 'rep';
  d.querySelector('#xgo').click();
  assert.strictEqual(w.eval('C.length'), 3, 'Excel: 3 baris laporan tanaman (baris Jumlah dilewati)');

  // 3) Cadangan JSON bolak-balik
  const cad = w.eval('JSON.stringify(BK())');
  w.eval('PT.length=0;JR.length=0');
  const inp = d.querySelector('#frs');
  Object.defineProperty(inp, 'files', { value: [{ name: 'c.json', text: async () => cad }] });
  inp.dispatchEvent(new w.Event('change'));
  await sleep(100);
  assert.strictEqual(w.eval('PT.length'), 2, 'Pemulihan cadangan: 2 petak');

  assert.deepStrictEqual(galat, [], 'Galat saat uji: ' + galat.join(' | '));
  console.log('OK: semua uji lolos');
  process.exit(0);
})().catch(e => { console.error('GAGAL:', e.message); process.exit(1); });
