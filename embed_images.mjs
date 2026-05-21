import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const HTML_FILE = path.join(__dirname, 'CNCAP2027_AEB_Scenarios.html');
const OUT_FILE  = path.join(__dirname, 'CNCAP2027_AEB_Scenarios_standalone.html');
const IMG_DIR   = path.join(__dirname, 'scenario_images');

// Show dir listing
console.log('Matching images:', fs.readdirSync(IMG_DIR).filter(f => /\.(jpg|jpeg|png|gif|webp|bmp)$/i.test(f)).length, 'files');

let html = fs.readFileSync(HTML_FILE, 'utf-8');

// Find all scenario_images/xxx.ext references
const re = /scenario_images\/([^\s"'<)]+)/gi;
const files = new Set();
let m;
while ((m = re.exec(html)) !== null) {
  files.add(m[1]);
}

console.log(`Unique images referenced: ${files.size}`);

for (const f of files) {
  const imgPath = path.join(IMG_DIR, f);
  if (!fs.existsSync(imgPath)) {
    console.warn(`  MISSING: ${f}`);
    continue;
  }
  const buf = fs.readFileSync(imgPath);
  const ext = path.extname(f).toLowerCase().replace('.jpg', 'jpeg');
  const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' :
               ext === 'webp' ? 'image/webp' : ext === 'bmp' ? 'image/bmp' : 'image/jpeg';
  const b64 = buf.toString('base64');
  const uri = `data:${mime};base64,${b64}`;

  // Replace in html — both src= and openLightbox('scenario_images/...')
  const escaped = f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  html = html.replace(new RegExp(`scenario_images/${escaped}`, 'g'), uri);
  console.log(`  OK: ${f}  (${(buf.length / 1024).toFixed(1)} KB)`);
}

fs.writeFileSync(OUT_FILE, html, 'utf-8');
console.log(`\nDone! Output: ${path.basename(OUT_FILE)}`);
