#!/usr/bin/env node
/*
 Menggabungkan banyak berkas KMZ/KML (dan, bila ada, laporan Excel 10 harian) menjadi SATU berkas
 cadangan JSON yang bisa dibuka di aplikasi (tab Data, Pulihkan dari cadangan).

 Skrip ini memakai kode aplikasi yang sama (index.html) sehingga hasilnya identik dengan impor manual.

 Pemakaian:
   node tools/gabung-kmz.js <folder-atau-zip> [-o wilayah.json] [--excel laporan.xlsx] [--laporan laporan.txt]

 Contoh:
   node tools/gabung-kmz.js data/privat/kmz -o data/privat/wilayah.json --excel data/privat/laporan.xlsx
*/
const fs = require('fs');
const path = require('path');
const { JSDOM, ResourceLoader, VirtualConsole } = require('jsdom');
const JSZip = require('jszip');

const root = path.join(__dirname, '..');
const arg = process.argv.slice(2);
const opt = (k, d) => { const i = arg.indexOf(k); return i >= 0 ? arg[i + 1] : d; };
const input = arg.find(a => !a.startsWith('-') && a !== opt('-o') && a !== opt('--excel') && a !== opt('--laporan'));
if (!input) { console.error('Pemakaian: node tools/gabung-kmz.js <folder-atau-zip> [-o wilayah.json] [--excel laporan.xlsx]'); process.exit(1); }
const out = opt('-o', 'wilayah.json'), excel = opt('--excel', null), laporan = opt('--laporan', null);
const log = []; const say = s => { console.log(s); log.push(s); };

class Lokal extends ResourceLoader {
  fetch(url) {
    const u = new URL(url);
    if (u.hostname !== 'localhost') return Promise.resolve(Buffer.from(''));
    const f = path.join(root, decodeURIComponent(u.pathname).replace(/^\//, '') || 'index.html');
    return Promise.resolve(fs.existsSync(f) ? fs.readFileSync(f) : Buffer.from(''));
  }
}
const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e => e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]);
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const t0 = Date.now();
  const vc = new VirtualConsole(); const galat = [];
  vc.on('jsdomError', e => galat.push(e.message));
  const dom = new JSDOM(fs.readFileSync(path.join(root, 'index.html'), 'utf8'), { url: 'http://localhost/', runScripts: 'dangerously', resources: new Lokal(), pretendToBeVisual: true, virtualConsole: vc });
  const w = dom.window, d = w.document;
  w.scrollTo = () => {};
  w.HTMLDialogElement.prototype.showModal = function () { this.setAttribute('open', ''); };
  await new Promise(r => (d.readyState === 'complete' ? r() : w.addEventListener('load', r)));
  await sleep(150);
  if (!w.XLSX) { console.error('Library vendor/ tidak termuat.'); process.exit(1); }

  // 1) kumpulkan berkas. KMZ dibuka di Node (ambil doc.kml) lalu diberikan ke aplikasi sebagai KML.
  const files = [];
  const isKm = n => /\.(kmz|kml)$/i.test(n) && !/^~\$|^\._/.test(path.basename(n)) && !/__MACOSX/.test(n);
  const kmlText = async (name, buf) => {
    if (/\.kml$/i.test(name)) return buf.toString('utf8');
    const z = await JSZip.loadAsync(buf);
    const ks = Object.keys(z.files).filter(k => !z.files[k].dir);
    const n = ks.find(k => /(^|\/)doc\.kml$/i.test(k)) || ks.find(k => /\.kml$/i.test(k));
    return n ? z.file(n).async('string') : null;
  };
  const add = async (name, buf) => {
    const t = await kmlText(name, buf);
    if (t == null) { say('  dilewati (tidak ada KML di dalamnya): ' + name); return; }
    files.push({ name: path.basename(name).replace(/\.kmz$/i, '.kml'), text: async () => t });
  };
  if (fs.statSync(input).isDirectory()) {
    for (const f of walk(input).filter(isKm).sort()) await add(f, fs.readFileSync(f));
  } else if (/\.zip$/i.test(input)) {
    const z = await JSZip.loadAsync(fs.readFileSync(input));
    for (const n of Object.keys(z.files).sort()) if (!z.files[n].dir && isKm(n)) await add(n, await z.file(n).async('nodebuffer'));
  } else await add(input, fs.readFileSync(input));
  if (!files.length) { console.error('Tidak ada berkas .kmz atau .kml.'); process.exit(1); }
  say(`Berkas dibaca: ${files.length}`);

  // 2) baca otomatis
  const inp = d.querySelector('#fk');
  Object.defineProperty(inp, 'files', { value: files, configurable: true });
  await inp.onchange({ target: inp });
  say('Pesan: ' + d.querySelector('#kmsg').textContent);
  say('\nKelompok penanda dan jenis yang dipilih otomatis:');
  d.querySelectorAll('#kl .krow').forEach(r => {
    const jenis = r.querySelector('select[data-r=t]'); const nama = jenis.options[jenis.selectedIndex].textContent;
    say(`  ${r.querySelector('b').textContent.padEnd(34)} ${r.querySelector('small').textContent.padEnd(60)} -> ${nama || 'Abaikan'}`);
  });
  d.querySelector('#kgo').click();
  say('\nHasil impor KMZ: ' + d.querySelector('#kmsg').textContent);
  const sebelum = w.eval('JSON.stringify(diNames())');
  const kmzDI = JSON.parse(sebelum);

  // 3) laporan Excel (opsional)
  if (excel) {
    const xb = fs.readFileSync(excel);
    w.__f = { name: path.basename(excel), arrayBuffer: async () => new w.Uint8Array(xb).buffer };
    await w.eval('xlOpen(__f)');
    if (d.querySelector('#xb').hidden) { say('\nPERINGATAN: format Excel tidak dikenali sebagai laporan 10 harian (lembar 5a dan 06), jadi tidak diimpor.'); }
    else {
      const excelDI = w.eval('XB.pt.map(p=>p.di)').filter(Boolean);
      d.querySelector('#xbg').click();
      say('\nHasil impor Excel: ' + d.querySelector('#msg').textContent);
      const key = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const kk = new Set(kmzDI.map(key)), ek = new Set(excelDI.map(key));
      const hanyaKmz = kmzDI.filter(x => !ek.has(key(x))), hanyaExcel = [...new Set(excelDI)].filter(x => !kk.has(key(x)));
      say(`\nPemeriksaan nama DI: ${kmzDI.length} DI di KMZ, ${ek.size} DI di Excel.`);
      say(`  Ada di KMZ tetapi tidak di Excel (${hanyaKmz.length}): ${hanyaKmz.slice(0, 40).join('; ') || '-'}`);
      say(`  Ada di Excel tetapi tidak di KMZ (${hanyaExcel.length}): ${hanyaExcel.slice(0, 40).join('; ') || '-'}`);
      say('  Bila ada nama yang seharusnya sama tetapi berbeda penulisan, samakan nama folder paling atas di KMZ dengan nama DI di Excel, lalu jalankan lagi.');
    }
  }

  // 4) ringkasan dan keluaran
  const ring = w.eval(`({di:new Set([...PT.map(p=>p.di),...D.bn.map(b=>b.di),...JR.map(j=>j.di)]).size,bn:D.bn.length,petak:PT.length,jr:JR.length,garis:D.sal.length,
    byk:JR.reduce((a,j)=>(a[j.k]=(a[j.k]||0)+1,a),{}),
    pj:JR.filter(j=>/^s/.test(j.k)).reduce((s,j)=>s+(j.pj||0),0)})`);
  say(`\nRingkasan: ${ring.di} DI, ${ring.bn} bendung, ${ring.petak} petak, ${ring.jr} bangunan dan saluran, ${ring.garis} garis saluran, panjang saluran ${(ring.pj / 1000).toFixed(1)} km`);
  say('Jenis bangunan dan saluran: ' + Object.entries(ring.byk).map(([k, v]) => `${k}=${v}`).join(', '));
  fs.writeFileSync(out, w.eval('JSON.stringify(BK())'));
  say(`\nDisimpan: ${out} (${(fs.statSync(out).size / 1048576).toFixed(2)} MB). Buka aplikasi, tab Data, Pulihkan dari cadangan.`);
  if (galat.length) say('Catatan galat skrip: ' + [...new Set(galat)].slice(0, 3).join(' | '));
  say(`Selesai dalam ${((Date.now() - t0) / 1000).toFixed(1)} detik.`);
  if (laporan) fs.writeFileSync(laporan, log.join('\n'));
  process.exit(0);
})().catch(e => { console.error('GAGAL:', e.message); process.exit(1); });
