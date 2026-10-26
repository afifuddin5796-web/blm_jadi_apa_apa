// Uji skrip penggabung KMZ: folder berisi KML dan KMZ, serta berkas zip berisi KMZ.
const { execFileSync } = require('child_process');
const fs = require('fs'), os = require('os'), path = require('path'), assert = require('assert');
const JSZip = require('jszip');

(async () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'gabung-'));
  const kml = fs.readFileSync(path.join(__dirname, '..', 'data/contoh/gaya-google-earth.kml'), 'utf8');
  fs.mkdirSync(path.join(tmp, 'kmz'));
  fs.writeFileSync(path.join(tmp, 'kmz', 'a.kml'), kml);
  const kmz = new JSZip(); kmz.file('doc.kml', kml.replace(/Contoh DI/g, 'DI Kedua'));
  fs.writeFileSync(path.join(tmp, 'kmz', 'b.kmz'), await kmz.generateAsync({ type: 'nodebuffer' }));
  const tool = path.join(__dirname, '..', 'tools', 'gabung-kmz.js');

  const out1 = path.join(tmp, 'folder.json');
  execFileSync('node', [tool, path.join(tmp, 'kmz'), '-o', out1], { stdio: 'pipe' });
  const j1 = JSON.parse(fs.readFileSync(out1, 'utf8'));
  assert.strictEqual(j1.app, 'irigasi');
  assert.strictEqual(j1.PT.length, 2, 'Folder: 2 petak (satu per DI)');
  assert.deepStrictEqual([...new Set(j1.PT.map(p => p.di))].sort(), ['Contoh DI', 'DI Kedua'], 'Nama DI dari folder paling atas');

  const zip = new JSZip(); zip.file('sub/b.kmz', await kmz.generateAsync({ type: 'nodebuffer' }));
  const zp = path.join(tmp, 'semua.zip'); fs.writeFileSync(zp, await zip.generateAsync({ type: 'nodebuffer' }));
  const out2 = path.join(tmp, 'zip.json');
  execFileSync('node', [tool, zp, '-o', out2], { stdio: 'pipe' });
  const j2 = JSON.parse(fs.readFileSync(out2, 'utf8'));
  assert.strictEqual(j2.PT.length, 1, 'Zip: 1 petak');
  console.log('OK: skrip penggabung lolos');
})().catch(e => { console.error('GAGAL:', e.message); process.exit(1); });
