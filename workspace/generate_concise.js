const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const pptxgen = require('pptxgenjs');
const html2pptx = require('/Users/seiiji/.claude/skills/pptx/scripts/html2pptx.js');

const outDir = path.resolve(__dirname);
const htmlDir = path.join(outDir, 'html_concise');
fs.mkdirSync(htmlDir, { recursive: true });

const C = { dark: '#1B2A4A', blue: '#2B6CB0', teal: '#00A99D', orange: '#E8734A', lightBg: '#EDF2F7', white: '#FFFFFF', text: '#1A202C', sub: '#4A5568', border: '#CBD5E0' };

const base = (bg = C.white) => `<!DOCTYPE html><html><head><style>
html{background:#fff}body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;display:flex;flex-direction:column;background:${bg}}
</style></head><body>`;
const end = '</body></html>';

async function createGradient(file, c1, c2, w = 1440, h = 810) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><defs><linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" style="stop-color:${c1}"/><stop offset="100%" style="stop-color:${c2}"/></linearGradient></defs><rect width="100%" height="100%" fill="url(#g)"/></svg>`;
  await sharp(Buffer.from(svg)).png().toFile(file);
}

const slides = [];

// Slide 1: Title
slides.push({ name: 'slide01.html', html: () => `<!DOCTYPE html><html><head><style>
html{background:#fff}body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;display:flex;flex-direction:column;background-image:url('${path.join(htmlDir,"bg_title.png")}');background-size:cover}
</style></head><body>
<div style="flex:1;display:flex;flex-direction:column;justify-content:center;margin:60pt 70pt">
  <div style="border-left:6pt solid ${C.teal};margin-bottom:20pt"><p style="color:#ffffff;font-size:14pt;margin:0 0 0 14pt;letter-spacing:2pt">AI CAREER AGENT</p></div>
  <h1 style="color:#ffffff;font-size:42pt;margin:0 0 4pt 0">YourCompany</h1>
  <h1 style="color:#ffffff;font-size:42pt;margin:0 0 16pt 0">Career Nexus</h1>
  <p style="color:${C.teal};font-size:20pt;margin:0 0 4pt 0;font-weight:bold">あなたのキャリアの選択肢、</p>
  <p style="color:${C.teal};font-size:20pt;margin:0;font-weight:bold">"半径5m"で決めていませんか？</p>
</div>
${end}` });

// Slide 2: 課題提起
slides.push({ name: 'slide02.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">大企業なのに"自分の周りしか見えない"</h2></div>
<div style="flex:1;display:flex;margin:24pt 40pt 24pt 40pt;gap:20pt">
  <div style="flex:1;display:flex;flex-direction:column;justify-content:center;gap:16pt">
    <div style="border-left:4pt solid ${C.teal};padding-left:12pt"><p style="color:${C.dark};font-size:14pt;font-weight:bold;margin:0">ロールモデルがいない</p></div>
    <div style="border-left:4pt solid ${C.blue};padding-left:12pt"><p style="color:${C.dark};font-size:14pt;font-weight:bold;margin:0">やりたいことと違う</p></div>
    <div style="border-left:4pt solid ${C.orange};padding-left:12pt"><p style="color:${C.dark};font-size:14pt;font-weight:bold;margin:0">部署の外が見えない</p></div>
    <div style="border-left:4pt solid #9F7AEA;padding-left:12pt"><p style="color:${C.dark};font-size:14pt;font-weight:bold;margin:0">異動が期待と違った</p></div>
  </div>
  <div style="flex:1;display:flex;flex-direction:column;justify-content:center;gap:14pt">
    <div style="background:${C.dark};border-radius:8pt;padding:20pt;text-align:center">
      <p style="color:${C.teal};font-size:40pt;font-weight:bold;margin:0">32%</p>
      <p style="color:#fff;font-size:12pt;margin:6pt 0 0 0">大卒3年以内離職率</p>
    </div>
    <div style="background:${C.dark};border-radius:8pt;padding:20pt;text-align:center">
      <p style="color:${C.orange};font-size:40pt;font-weight:bold;margin:0">5%</p>
      <p style="color:#fff;font-size:12pt;margin:6pt 0 0 0">日本の熱意ある社員割合</p>
      <p style="color:#ffffff88;font-size:9pt;margin:4pt 0 0 0">世界最低水準（Gallup 2023）</p>
    </div>
  </div>
</div>
${end}` });

// Slide 3: 原因 ― 3つのギャップ
slides.push({ name: 'slide03.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">"知れば辞めない"のに、知る手段がない</h2></div>
<div style="flex:1;display:flex;flex-direction:column;justify-content:center;gap:16pt;margin:20pt 50pt 30pt 50pt">
  <div style="display:flex;gap:16pt">
    <div style="flex:1;background:${C.teal};border-radius:10pt;padding:24pt;text-align:center">
      <p style="color:#fff;font-size:18pt;font-weight:bold;margin:0 0 10pt 0">情報のギャップ</p>
      <p style="color:#ffffffdd;font-size:12pt;margin:0">多様なキャリアを歩む社員がいるのに知る手段がない</p>
    </div>
    <div style="flex:1;background:${C.blue};border-radius:10pt;padding:24pt;text-align:center">
      <p style="color:#fff;font-size:18pt;font-weight:bold;margin:0 0 10pt 0">マッチングのギャップ</p>
      <p style="color:#ffffffdd;font-size:12pt;margin:0">社内公募は氷山の一角。可能性が埋もれている</p>
    </div>
    <div style="flex:1;background:${C.orange};border-radius:10pt;padding:24pt;text-align:center">
      <p style="color:#fff;font-size:18pt;font-weight:bold;margin:0 0 10pt 0">対話のギャップ</p>
      <p style="color:#ffffffdd;font-size:12pt;margin:0">気軽に壁打ちできる相手がいない</p>
    </div>
  </div>
  <div style="text-align:center"><p style="color:${C.dark};font-size:14pt;font-weight:bold;margin:0">これらのギャップをすべて埋めるのが YourCompany Career Nexus</p></div>
</div>
${end}` });

// Slide 4: 解決策 3フェーズ
slides.push({ name: 'slide04.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">3つのフェーズで進化</h2></div>
<div style="flex:1;display:flex;gap:16pt;margin:20pt 40pt 24pt 40pt">
  <div style="flex:1;border:2pt solid ${C.teal};border-radius:10pt;padding:18pt;text-align:center">
    <div style="background:${C.teal};border-radius:50%;width:40pt;height:40pt;display:flex;align-items:center;justify-content:center;margin:0 auto 10pt auto"><p style="color:#fff;font-size:20pt;font-weight:bold;margin:0">1</p></div>
    <p style="color:${C.dark};font-size:16pt;font-weight:bold;margin:0 0 8pt 0">発見</p>
    <p style="color:${C.sub};font-size:11pt;margin:0">ロールモデルマップ構築</p>
    <p style="color:${C.sub};font-size:11pt;margin:2pt 0 0 0">AIとキャリア壁打ち</p>
    <p style="color:${C.sub};font-size:11pt;margin:2pt 0 0 0">部署情報・風土の把握</p>
  </div>
  <div style="flex:1;border:2pt solid ${C.blue};border-radius:10pt;padding:18pt;text-align:center">
    <div style="background:${C.blue};border-radius:50%;width:40pt;height:40pt;display:flex;align-items:center;justify-content:center;margin:0 auto 10pt auto"><p style="color:#fff;font-size:20pt;font-weight:bold;margin:0">2</p></div>
    <p style="color:${C.dark};font-size:16pt;font-weight:bold;margin:0 0 8pt 0">継続</p>
    <p style="color:${C.sub};font-size:11pt;margin:0">個人別キャリアマップ</p>
    <p style="color:${C.sub};font-size:11pt;margin:2pt 0 0 0">AIによる励ましの声かけ</p>
    <p style="color:${C.sub};font-size:11pt;margin:2pt 0 0 0">モチベーション把握</p>
  </div>
  <div style="flex:1;border:2pt solid ${C.orange};border-radius:10pt;padding:18pt;text-align:center">
    <div style="background:${C.orange};border-radius:50%;width:40pt;height:40pt;display:flex;align-items:center;justify-content:center;margin:0 auto 10pt auto"><p style="color:#fff;font-size:20pt;font-weight:bold;margin:0">3</p></div>
    <p style="color:${C.dark};font-size:16pt;font-weight:bold;margin:0 0 8pt 0">組織</p>
    <p style="color:${C.sub};font-size:11pt;margin:0">スキル可視化</p>
    <p style="color:${C.sub};font-size:11pt;margin:2pt 0 0 0">最適人員配置の自動化</p>
    <p style="color:${C.sub};font-size:11pt;margin:2pt 0 0 0">勉強会・メンタリング</p>
  </div>
</div>
${end}` });

// Slide 5: デモ
slides.push({ name: 'slide05.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">ユーザー体験フロー</h2></div>
<div style="flex:1;display:flex;align-items:center;margin:20pt 30pt 24pt 30pt;gap:8pt">
  <div style="flex:1;background:${C.teal};border-radius:10pt;padding:18pt;text-align:center">
    <p style="color:#fff;font-size:28pt;font-weight:bold;margin:0 0 6pt 0">1</p>
    <p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">相談</p>
  </div>
  <p style="color:${C.border};font-size:24pt;margin:0">&#9654;</p>
  <div style="flex:1;background:${C.blue};border-radius:10pt;padding:18pt;text-align:center">
    <p style="color:#fff;font-size:28pt;font-weight:bold;margin:0 0 6pt 0">2</p>
    <p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">発見</p>
  </div>
  <p style="color:${C.border};font-size:24pt;margin:0">&#9654;</p>
  <div style="flex:1;background:#5A67D8;border-radius:10pt;padding:18pt;text-align:center">
    <p style="color:#fff;font-size:28pt;font-weight:bold;margin:0 0 6pt 0">3</p>
    <p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">対話</p>
  </div>
  <p style="color:${C.border};font-size:24pt;margin:0">&#9654;</p>
  <div style="flex:1;background:${C.orange};border-radius:10pt;padding:18pt;text-align:center">
    <p style="color:#fff;font-size:28pt;font-weight:bold;margin:0 0 6pt 0">4</p>
    <p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">提案</p>
  </div>
  <p style="color:${C.border};font-size:24pt;margin:0">&#9654;</p>
  <div style="flex:1;background:#9F7AEA;border-radius:10pt;padding:18pt;text-align:center">
    <p style="color:#fff;font-size:28pt;font-weight:bold;margin:0 0 6pt 0">5</p>
    <p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">行動</p>
  </div>
</div>
<div style="text-align:center;margin:0 0 20pt 0"><p style="color:${C.sub};font-size:11pt;margin:0">※ 詳細はデモ動画にて紹介</p></div>
${end}` });

// Slide 6: 進化ユースケース
slides.push({ name: 'slide06.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">進化後のユースケース</h2></div>
<div style="flex:1;display:flex;gap:20pt;margin:24pt 40pt 28pt 40pt">
  <div style="flex:1;background:${C.blue};border-radius:10pt;padding:24pt;display:flex;flex-direction:column;justify-content:center">
    <p style="color:#fff;font-size:18pt;font-weight:bold;margin:0 0 14pt 0">Phase 2</p>
    <p style="color:#ffffffdd;font-size:13pt;margin:0 0 8pt 0">AIが日々のモチベーションを把握</p>
    <p style="color:#ffffffdd;font-size:13pt;margin:0">自発的な声かけで社員を支える</p>
  </div>
  <div style="flex:1;background:${C.orange};border-radius:10pt;padding:24pt;display:flex;flex-direction:column;justify-content:center">
    <p style="color:#fff;font-size:18pt;font-weight:bold;margin:0 0 14pt 0">Phase 3</p>
    <p style="color:#ffffffdd;font-size:13pt;margin:0 0 8pt 0">最適人員配置の自動化</p>
    <p style="color:#ffffffdd;font-size:13pt;margin:0">勉強会・メンタリングを自動マッチング</p>
  </div>
</div>
${end}` });

// Slide 7: 独自性
slides.push({ name: 'slide07.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">独自性</h2></div>
<div style="flex:1;display:flex;gap:16pt;margin:28pt 40pt 32pt 40pt;align-items:center">
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-top:5pt solid ${C.teal}">
    <p style="color:${C.teal};font-size:36pt;font-weight:bold;margin:0">01</p>
    <p style="color:${C.dark};font-size:15pt;font-weight:bold;margin:10pt 0 6pt 0">自社データ x AI</p>
    <p style="color:${C.sub};font-size:11pt;margin:0">その会社ならではの提案</p>
  </div>
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-top:5pt solid ${C.blue}">
    <p style="color:${C.blue};font-size:36pt;font-weight:bold;margin:0">02</p>
    <p style="color:${C.dark};font-size:15pt;font-weight:bold;margin:10pt 0 6pt 0">対話で深掘り</p>
    <p style="color:${C.sub};font-size:11pt;margin:0">個人に寄り添う提案</p>
  </div>
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-top:5pt solid ${C.orange}">
    <p style="color:${C.orange};font-size:36pt;font-weight:bold;margin:0">03</p>
    <p style="color:${C.dark};font-size:15pt;font-weight:bold;margin:10pt 0 6pt 0">蓄積が組織価値に</p>
    <p style="color:${C.sub};font-size:11pt;margin:0">個人から組織への相乗効果</p>
  </div>
</div>
${end}` });

// Slide 8: アーキテクチャ
slides.push({ name: 'slide08.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">アーキテクチャ</h2></div>
<div style="flex:1;display:flex;flex-direction:column;gap:10pt;margin:20pt 50pt 24pt 50pt;justify-content:center">
  <div style="display:flex;gap:10pt">
    <div style="width:110pt;background:${C.dark};border-radius:6pt;padding:10pt;display:flex;align-items:center"><p style="color:#fff;font-size:11pt;font-weight:bold;margin:0">ユーザー接点</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:${C.lightBg};border-radius:6pt;padding:10pt;text-align:center"><p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0">Slack / Teams</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:110pt;background:${C.teal};border-radius:6pt;padding:10pt;display:flex;align-items:center"><p style="color:#fff;font-size:11pt;font-weight:bold;margin:0">AIエージェント</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:#E6FFFA;border-radius:6pt;padding:10pt;text-align:center;border:1pt solid ${C.teal}"><p style="color:${C.dark};font-size:10pt;margin:0">壁打ち</p></div>
      <div style="flex:1;background:#E6FFFA;border-radius:6pt;padding:10pt;text-align:center;border:1pt solid ${C.teal}"><p style="color:${C.dark};font-size:10pt;margin:0">励まし</p></div>
      <div style="flex:1;background:#E6FFFA;border-radius:6pt;padding:10pt;text-align:center;border:1pt solid ${C.teal}"><p style="color:${C.dark};font-size:10pt;margin:0">マッチング</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:110pt;background:${C.blue};border-radius:6pt;padding:10pt;display:flex;align-items:center"><p style="color:#fff;font-size:11pt;font-weight:bold;margin:0">データ / RAG</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:#EBF4FF;border-radius:6pt;padding:10pt;text-align:center;border:1pt solid ${C.blue}"><p style="color:${C.dark};font-size:10pt;margin:0">インタビュー</p></div>
      <div style="flex:1;background:#EBF4FF;border-radius:6pt;padding:10pt;text-align:center;border:1pt solid ${C.blue}"><p style="color:${C.dark};font-size:10pt;margin:0">公募・スキル</p></div>
      <div style="flex:1;background:#EBF4FF;border-radius:6pt;padding:10pt;text-align:center;border:1pt solid ${C.blue}"><p style="color:${C.dark};font-size:10pt;margin:0">キャリアマップ</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:110pt;background:#718096;border-radius:6pt;padding:10pt;display:flex;align-items:center"><p style="color:#fff;font-size:11pt;font-weight:bold;margin:0">社内連携</p></div>
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;padding:10pt;text-align:center;border:1pt dashed #718096"><p style="color:${C.sub};font-size:10pt;margin:0">既存キャリアシステム / 人事システム（将来拡張）</p></div>
  </div>
</div>
${end}` });

// Slide 9: 期待される効果
slides.push({ name: 'slide09.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:20pt 40pt"><h2 style="color:#fff;font-size:24pt;margin:0">期待される効果</h2></div>
<div style="flex:1;display:flex;gap:18pt;margin:28pt 40pt 32pt 40pt;align-items:center">
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-bottom:5pt solid ${C.teal}">
    <p style="color:${C.teal};font-size:16pt;font-weight:bold;margin:0">離職率低下</p>
  </div>
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-bottom:5pt solid ${C.blue}">
    <p style="color:${C.blue};font-size:16pt;font-weight:bold;margin:0">エンゲージメント向上</p>
  </div>
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-bottom:5pt solid ${C.orange}">
    <p style="color:${C.orange};font-size:16pt;font-weight:bold;margin:0">人事工数削減</p>
  </div>
  <div style="flex:1;text-align:center;padding:24pt;background:${C.lightBg};border-radius:10pt;border-bottom:5pt solid #9F7AEA">
    <p style="color:#9F7AEA;font-size:16pt;font-weight:bold;margin:0">組織流動性向上</p>
  </div>
</div>
${end}` });

// Slide 10: クロージング
slides.push({ name: 'slide10.html', html: () => `<!DOCTYPE html><html><head><style>
html{background:#fff}body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;display:flex;flex-direction:column;background-image:url('${path.join(htmlDir,"bg_closing.png")}');background-size:cover}
</style></head><body>
<div style="flex:1;display:flex;flex-direction:column;justify-content:center;align-items:center;margin:40pt">
  <h1 style="color:#fff;font-size:34pt;margin:0 0 4pt 0;text-align:center">半径5mの外に、</h1>
  <h1 style="color:#fff;font-size:34pt;margin:0 0 20pt 0;text-align:center">あなたのキャリアがある</h1>
  <p style="color:${C.teal};font-size:18pt;font-weight:bold;margin:0">YourCompany Career Nexus</p>
</div>
${end}` });

async function main() {
  await createGradient(path.join(htmlDir, 'bg_title.png'), '#1B2A4A', '#0D1B2A');
  await createGradient(path.join(htmlDir, 'bg_closing.png'), '#1B2A4A', '#162033');

  for (const s of slides) {
    fs.writeFileSync(path.join(htmlDir, s.name), s.html());
  }

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'YourCompany Career Nexus - Concise';
  pptx.author = 'Career Nexus Team';

  for (const s of slides) {
    await html2pptx(path.join(htmlDir, s.name), pptx);
  }

  const outFile = path.join(outDir, 'career_nexus_concise.pptx');
  await pptx.writeFile({ fileName: outFile });
  console.log('Created:', outFile);
}

main().catch(console.error);
