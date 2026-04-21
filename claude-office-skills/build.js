const pptxgen = require('pptxgenjs');
const path = require('path');
const html2pptx = require('./html2pptx.js');

(async () => {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'Inner Compass — 社内ロールモデル発見型キャリアAIエージェント';
  pptx.author = 'Google AI Agent Hackathon Submission';

  const slideDir = path.join(__dirname, 'slides');
  for (let i = 1; i <= 10; i++) {
    const file = path.join(slideDir, `slide${i}.html`);
    await html2pptx(file, pptx);
    console.log(`✓ slide${i}.html processed`);
  }

  await pptx.writeFile({ fileName: path.join(__dirname, 'inner-compass.pptx') });
  console.log('Saved: inner-compass.pptx');
})().catch(e => { console.error(e); process.exit(1); });
