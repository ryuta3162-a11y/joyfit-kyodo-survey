/**
 * サービス画像を縮小JPEG→data URL化し、service-images-embedded.js を生成する。
 * 実行: node embed-service-images.mjs
 */
import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = __dirname;

// ファイル番号: 18=FW / 19=マシン / 20=有酸素（従来の割当を修正）
const mapping = [
    ['FWエリア', '18.png'],
    ['マシンエリア', '19.png'],
    ['有酸素エリア', '20.png'],
    ['ストレッチエリア', 'stretch_area.png'],
    ['サウナ', '23.png'],
    ['ミストサウナ', 'mist_sauna.png'],
    ['スタジオ', '17.png'],
    ['マシンピラティス(リフォーマー)', '16.png'],
    ['エステ', '21.png'],
    ['タンニング', '22.png']
];

const cache = new Map();

for (const file of [...new Set(mapping.map(([, f]) => f))]) {
    const input = path.join(root, file);
    const tmpOut = path.join(root, `_embed_${file.replace(/\W/g, '_')}.jpg`);
    const cmd = `npx --yes sharp-cli -i "${input}" -o "${tmpOut}" resize 280 --fit inside -f jpeg -q 68`;
    execSync(cmd, { cwd: root, shell: true, stdio: 'inherit' });
    const buf = fs.readFileSync(tmpOut);
    fs.unlinkSync(tmpOut);
    cache.set(file, `data:image/jpeg;base64,${buf.toString('base64')}`);
}

const obj = {};
for (const [service, file] of mapping) {
    obj[service] = cache.get(file);
}

const line = `const SERVICE_IMAGE_DATA_URLS = ${JSON.stringify(obj)};`;
const outPath = path.join(root, 'service-images-embedded.js');
fs.writeFileSync(outPath, `${line}\n`, 'utf8');
console.log('Wrote', outPath, 'chars:', line.length);

// 1ファイル配布用: 画像data URLを index 内に埋め込んだ index-standalone.html
const indexPath = path.join(root, 'index.html');
let indexHtml = fs.readFileSync(indexPath, 'utf8');
const needle = '<script src="service-images-embedded.js" charset="utf-8"></script>';
if (indexHtml.includes(needle)) {
    const standalone = indexHtml.replace(
        needle,
        `<script>\n${line}\n</script>`
    );
    fs.writeFileSync(path.join(root, 'index-standalone.html'), standalone, 'utf8');
    console.log('Wrote index-standalone.html (single file, ~' + Math.round(standalone.length / 1024) + ' KB)');
}
