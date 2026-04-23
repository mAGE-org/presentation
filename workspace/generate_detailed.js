const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const pptxgen = require('pptxgenjs');
const html2pptx = require('/Users/seiiji/.claude/skills/pptx/scripts/html2pptx.js');

const outDir = path.resolve(__dirname);
const htmlDir = path.join(outDir, 'html_detailed');
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
<div style="flex:1;display:flex;flex-direction:column;justify-content:center;margin:50pt 60pt">
  <div style="border-left:6pt solid ${C.teal};margin-bottom:16pt"><p style="color:#ffffff;font-size:14pt;margin:0 0 0 14pt;letter-spacing:2pt">AI CAREER AGENT</p></div>
  <h1 style="color:#ffffff;font-size:36pt;margin:0 0 4pt 0">YourCompany</h1>
  <h1 style="color:#ffffff;font-size:36pt;margin:0 0 12pt 0">Career Nexus</h1>
  <p style="color:${C.teal};font-size:18pt;margin:0 0 4pt 0;font-weight:bold">あなたのキャリアの選択肢、</p>
  <p style="color:${C.teal};font-size:18pt;margin:0 0 30pt 0;font-weight:bold">"半径5m"で決めていませんか？</p>
  <p style="color:#ffffffcc;font-size:12pt;margin:0">その会社ならではのキャリアの実現により、個人と会社がいきいきと働く世界を</p>
</div>
${end}` });

// Slide 2: 課題提起
slides.push({ name: 'slide02.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">課題提起 ― 大企業なのに"自分の周りしか見えない"問題</h2></div>
<div style="flex:1;display:flex;flex-direction:column;margin:20pt 40pt 20pt 40pt">
  <p style="color:${C.dark};font-size:14pt;font-weight:bold;margin:0 0 14pt 0">同僚が転職していく、そのリアルな理由</p>
  <div style="display:flex;gap:12pt;margin-bottom:16pt">
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;border-left:4pt solid ${C.teal};padding:10pt 12pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">ロールモデルがいない</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">直属の上司・先輩の将来が自分の将来に見えてしまう</p>
    </div>
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;border-left:4pt solid ${C.blue};padding:10pt 12pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">やりたいことと違う</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">でも社内に他の選択肢があることを知らない</p>
    </div>
  </div>
  <div style="display:flex;gap:12pt;margin-bottom:16pt">
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;border-left:4pt solid ${C.orange};padding:10pt 12pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">部署の外が見えない</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">大企業なのに、知っている社員はほんの一握り</p>
    </div>
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;border-left:4pt solid #9F7AEA;padding:10pt 12pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">異動が期待と違った</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">異動先の風土・同僚・業務を事前に知る手段がない</p>
    </div>
  </div>
  <div style="background:${C.dark};border-radius:6pt;padding:14pt 18pt">
    <p style="color:${C.teal};font-size:13pt;font-weight:bold;margin:0">辞める人の多くは「会社が嫌い」なのではなく、「この会社での可能性が見えない」から辞めている</p>
  </div>
</div>
${end}` });

// Slide 3: 定量データ
slides.push({ name: 'slide03.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">数字で見る課題の深刻さ</h2></div>
<div style="flex:1;display:flex;gap:16pt;margin:20pt 40pt 20pt 40pt">
  <div style="flex:1;display:flex;flex-direction:column;gap:12pt">
    <div style="background:${C.lightBg};border-radius:8pt;padding:14pt;text-align:center">
      <p style="color:${C.teal};font-size:36pt;font-weight:bold;margin:0">32-34%</p>
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:4pt 0 0 0">大卒 入社3年以内の離職率</p>
      <p style="color:${C.sub};font-size:9pt;margin:2pt 0 0 0">3人に1人が3年で辞める</p>
    </div>
    <div style="background:${C.lightBg};border-radius:8pt;padding:14pt;text-align:center">
      <p style="color:${C.orange};font-size:36pt;font-weight:bold;margin:0">5-6%</p>
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:4pt 0 0 0">日本の「熱意ある社員」割合</p>
      <p style="color:${C.sub};font-size:9pt;margin:2pt 0 0 0">世界約140カ国中で最低水準（Gallup 2023）</p>
    </div>
  </div>
  <div style="flex:1;display:flex;flex-direction:column;gap:12pt">
    <div style="background:${C.dark};border-radius:8pt;padding:14pt">
      <p style="color:#fff;font-size:11pt;font-weight:bold;margin:0 0 8pt 0">若手の離職理由TOP</p>
      <p style="color:${C.teal};font-size:10pt;margin:0 0 4pt 0">1. やりたい仕事と違った ― 約25-30%</p>
      <p style="color:${C.teal};font-size:10pt;margin:0 0 4pt 0">2. キャリアの見通しが立たない ― 約25-28%</p>
      <p style="color:${C.teal};font-size:10pt;margin:0">3. 人間関係 ― 約20-25%</p>
      <p style="color:#ffffff99;font-size:8pt;margin:8pt 0 0 0">出典: 内閣府「子供・若者白書」、エン・ジャパン調査</p>
    </div>
    <div style="background:${C.lightBg};border-radius:8pt;padding:14pt;text-align:center">
      <p style="color:${C.blue};font-size:28pt;font-weight:bold;margin:0">150-200万円</p>
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:4pt 0 0 0">新卒1人あたりの離職コスト</p>
      <p style="color:${C.sub};font-size:9pt;margin:2pt 0 0 0">100人離職で年間1.5-2億円の損失</p>
    </div>
  </div>
</div>
<p style="color:${C.sub};font-size:7pt;margin:0 40pt 10pt 40pt">出典: 厚生労働省「雇用動向調査」「新規学卒就職者の離職状況」/ Gallup "State of the Global Workplace 2023" / リクルート「就職白書」</p>
${end}` });

// Slide 4: 原因
slides.push({ name: 'slide04.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">本質的な原因 ― "知れば辞めない"のに、知る手段がない</h2></div>
<div style="flex:1;display:flex;flex-direction:column;margin:16pt 40pt 16pt 40pt">
  <p style="color:${C.dark};font-size:11pt;margin:0 0 12pt 0">既存の手段はいずれも"発見"を生む設計になっていない</p>
  <div style="display:flex;gap:10pt;margin-bottom:16pt">
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;padding:10pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">1on1</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">上司の知見・視野に依存。部署外のキャリアは話題にすら上がらない</p>
    </div>
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;padding:10pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">社内公募</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">公募は社内のごく一部。マッチする可能性が多数埋もれている</p>
    </div>
    <div style="flex:1;background:${C.lightBg};border-radius:6pt;padding:10pt">
      <p style="color:${C.dark};font-size:11pt;font-weight:bold;margin:0 0 4pt 0">外部コーチング</p>
      <p style="color:${C.sub};font-size:9pt;margin:0">自社固有のキャリア制度・部署の風土を知らない。一般論に留まる</p>
    </div>
  </div>
  <p style="color:${C.dark};font-size:12pt;font-weight:bold;margin:0 0 10pt 0">3つの構造的ギャップ</p>
  <div style="display:flex;gap:10pt">
    <div style="flex:1;background:${C.teal};border-radius:6pt;padding:12pt">
      <p style="color:#fff;font-size:12pt;font-weight:bold;margin:0 0 4pt 0">情報のギャップ</p>
      <p style="color:#ffffffdd;font-size:9pt;margin:0">多様なキャリアを歩む社員がいるのに、その存在・経験を知る手段がない</p>
    </div>
    <div style="flex:1;background:${C.blue};border-radius:6pt;padding:12pt">
      <p style="color:#fff;font-size:12pt;font-weight:bold;margin:0 0 4pt 0">マッチングのギャップ</p>
      <p style="color:#ffffffdd;font-size:9pt;margin:0">社内公募は氷山の一角。自分に合うポジションが埋もれている</p>
    </div>
    <div style="flex:1;background:${C.orange};border-radius:6pt;padding:12pt">
      <p style="color:#fff;font-size:12pt;font-weight:bold;margin:0 0 4pt 0">対話のギャップ</p>
      <p style="color:#ffffffdd;font-size:9pt;margin:0">一方通行の情報では「自分ならどうか」を考えられない</p>
    </div>
  </div>
</div>
${end}` });

// Slide 5: 解決策
slides.push({ name: 'slide05.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">解決策 ― YourCompany Career Nexus</h2></div>
<div style="margin:12pt 40pt 0 40pt"><p style="color:${C.dark};font-size:11pt;margin:0">3つのフェーズで進化するAIキャリアエージェント</p></div>
<div style="flex:1;display:flex;gap:14pt;margin:12pt 40pt 20pt 40pt">
  <div style="flex:1;border:2pt solid ${C.teal};border-radius:8pt;padding:14pt;display:flex;flex-direction:column">
    <div style="background:${C.teal};border-radius:4pt;padding:6pt 10pt;margin-bottom:10pt"><p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">Phase 1: 発見</p></div>
    <p style="color:${C.dark};font-size:10pt;font-weight:bold;margin:0 0 6pt 0">ロールモデルマップ & キャリア壁打ち</p>
    <ul style="color:${C.sub};font-size:9pt;margin:0;padding-left:14pt;line-height:1.6">
      <li>自社内データでAIとキャリア壁打ち</li>
      <li>インタビュー記事からロールモデルマップ構築</li>
      <li>ロールモデルとの"疑似対話"</li>
      <li>社内公募と連動したキャリアパス提案</li>
      <li>キャリア制度・部署情報・風土データを取り込み</li>
    </ul>
  </div>
  <div style="flex:1;border:2pt solid ${C.blue};border-radius:8pt;padding:14pt;display:flex;flex-direction:column">
    <div style="background:${C.blue};border-radius:4pt;padding:6pt 10pt;margin-bottom:10pt"><p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">Phase 2: 継続</p></div>
    <p style="color:${C.dark};font-size:10pt;font-weight:bold;margin:0 0 6pt 0">キャリアマップ & モチベーション支援</p>
    <ul style="color:${C.sub};font-size:9pt;margin:0;padding-left:14pt;line-height:1.6">
      <li>会話内容から個人別キャリアマップ構築</li>
      <li>日々の会話でモチベーション把握</li>
      <li>AIが自発的に励ましの声かけ</li>
      <li>公募更新通知・新人メンタリング</li>
    </ul>
  </div>
  <div style="flex:1;border:2pt solid ${C.orange};border-radius:8pt;padding:14pt;display:flex;flex-direction:column">
    <div style="background:${C.orange};border-radius:4pt;padding:6pt 10pt;margin-bottom:10pt"><p style="color:#fff;font-size:13pt;font-weight:bold;margin:0">Phase 3: 組織</p></div>
    <p style="color:${C.dark};font-size:10pt;font-weight:bold;margin:0 0 6pt 0">スキル可視化 & 最適配置</p>
    <ul style="color:${C.sub};font-size:9pt;margin:0;padding-left:14pt;line-height:1.6">
      <li>全社員のキャリア・スキル可視化</li>
      <li>プロジェクトへの最適人員配置を自動化</li>
      <li>勉強会・メンタリングをマッチング</li>
      <li>流動的な組織の実現</li>
    </ul>
  </div>
</div>
${end}` });

// Slide 6: デモ
slides.push({ name: 'slide06.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">デモ ― ユーザー体験フロー</h2></div>
<div style="margin:10pt 40pt 0 40pt"><p style="color:${C.sub};font-size:11pt;margin:0">「今の仕事の延長しか見えない若手社員」のケース</p></div>
<div style="flex:1;display:flex;align-items:center;margin:10pt 30pt 20pt 30pt;gap:6pt">
  <div style="flex:1;background:${C.teal};border-radius:8pt;padding:14pt;text-align:center">
    <p style="color:#fff;font-size:24pt;font-weight:bold;margin:0">1</p>
    <p style="color:#fff;font-size:11pt;font-weight:bold;margin:6pt 0 4pt 0">悩みの相談</p>
    <p style="color:#ffffffcc;font-size:8pt;margin:0">Slack/Teamsから気軽にAIに相談</p>
  </div>
  <p style="color:${C.border};font-size:20pt;margin:0">&#9654;</p>
  <div style="flex:1;background:${C.blue};border-radius:8pt;padding:14pt;text-align:center">
    <p style="color:#fff;font-size:24pt;font-weight:bold;margin:0">2</p>
    <p style="color:#fff;font-size:11pt;font-weight:bold;margin:6pt 0 4pt 0">ロールモデル発見</p>
    <p style="color:#ffffffcc;font-size:8pt;margin:0">意外なキャリアを歩んだ社員を提案</p>
  </div>
  <p style="color:${C.border};font-size:20pt;margin:0">&#9654;</p>
  <div style="flex:1;background:#5A67D8;border-radius:8pt;padding:14pt;text-align:center">
    <p style="color:#fff;font-size:24pt;font-weight:bold;margin:0">3</p>
    <p style="color:#fff;font-size:11pt;font-weight:bold;margin:6pt 0 4pt 0">疑似対話</p>
    <p style="color:#ffffffcc;font-size:8pt;margin:0">考え方・学び方をAI対話で深掘り</p>
  </div>
  <p style="color:${C.border};font-size:20pt;margin:0">&#9654;</p>
  <div style="flex:1;background:${C.orange};border-radius:8pt;padding:14pt;text-align:center">
    <p style="color:#fff;font-size:24pt;font-weight:bold;margin:0">4</p>
    <p style="color:#fff;font-size:11pt;font-weight:bold;margin:6pt 0 4pt 0">キャリアプラン提案</p>
    <p style="color:#ffffffcc;font-size:8pt;margin:0">自社ならではの具体的パスが見える</p>
  </div>
  <p style="color:${C.border};font-size:20pt;margin:0">&#9654;</p>
  <div style="flex:1;background:#9F7AEA;border-radius:8pt;padding:14pt;text-align:center">
    <p style="color:#fff;font-size:24pt;font-weight:bold;margin:0">5</p>
    <p style="color:#fff;font-size:11pt;font-weight:bold;margin:6pt 0 4pt 0">次のアクション</p>
    <p style="color:#ffffffcc;font-size:8pt;margin:0">公募情報・勉強会・メンター紹介</p>
  </div>
</div>
<div style="margin:0 40pt 16pt 40pt;text-align:center"><p style="color:${C.sub};font-size:10pt;margin:0">※ 詳細はデモ動画にて紹介</p></div>
${end}` });

// Slide 7: 進化ユースケース
slides.push({ name: 'slide07.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">進化後のユースケース</h2></div>
<div style="flex:1;display:flex;gap:16pt;margin:20pt 40pt 20pt 40pt">
  <div style="flex:1;border:2pt solid ${C.blue};border-radius:8pt;padding:16pt;display:flex;flex-direction:column">
    <div style="background:${C.blue};border-radius:4pt;padding:8pt 12pt;margin-bottom:12pt"><p style="color:#fff;font-size:14pt;font-weight:bold;margin:0">Phase 2: モチベーション支援</p></div>
    <ul style="color:${C.dark};font-size:10pt;margin:0;padding-left:14pt;line-height:1.8">
      <li>会話履歴からキャリア志向・モチベーションを日々把握・更新</li>
      <li>AIから自発的に声かけでモチベーションアップ</li>
    </ul>
    <div style="background:${C.lightBg};border-radius:4pt;padding:10pt;margin-top:10pt">
      <p style="color:${C.sub};font-size:9pt;margin:0 0 4pt 0;font-style:italic">「先日話していたスキルアップの件、社内にこんな勉強会がありますよ」</p>
      <p style="color:${C.sub};font-size:9pt;margin:0;font-style:italic">「同じ悩みを乗り越えた○○さんの話を聞いてみませんか？」</p>
    </div>
  </div>
  <div style="flex:1;border:2pt solid ${C.orange};border-radius:8pt;padding:16pt;display:flex;flex-direction:column">
    <div style="background:${C.orange};border-radius:4pt;padding:8pt 12pt;margin-bottom:12pt"><p style="color:#fff;font-size:14pt;font-weight:bold;margin:0">Phase 3: 組織最適化</p></div>
    <ul style="color:${C.dark};font-size:10pt;margin:0;padding-left:14pt;line-height:1.8">
      <li>全社員のスキルマップでプロジェクトへの最適人員配置</li>
      <li>流動的な組織 ― 異動のハードルを下げる</li>
      <li>人と人を繋いで勉強会・メンタリングを自動開催</li>
    </ul>
    <div style="background:${C.lightBg};border-radius:4pt;padding:10pt;margin-top:10pt">
      <p style="color:${C.sub};font-size:9pt;margin:0 0 4pt 0;font-style:italic">スキルの近い社員同士の学習コミュニティ形成</p>
      <p style="color:${C.sub};font-size:9pt;margin:0;font-style:italic">シニア社員と若手の自動メンタリングマッチング</p>
    </div>
  </div>
</div>
${end}` });

// Slide 8: 独自性
slides.push({ name: 'slide08.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">独自性 ― なぜ Career Nexus でなければならないか</h2></div>
<div style="flex:1;display:flex;gap:14pt;margin:20pt 40pt 24pt 40pt">
  <div style="flex:1;background:${C.lightBg};border-radius:8pt;padding:18pt;border-top:4pt solid ${C.teal};display:flex;flex-direction:column">
    <p style="color:${C.teal};font-size:32pt;font-weight:bold;margin:0">01</p>
    <p style="color:${C.dark};font-size:13pt;font-weight:bold;margin:8pt 0 8pt 0">自社データ x AI</p>
    <p style="color:${C.sub};font-size:10pt;margin:0;line-height:1.6">インタビュー記事・スキルデータ・公募情報をRAGで取り込み。汎用キャリアアドバイスではなく、その会社ならではのキャリア提案が可能</p>
  </div>
  <div style="flex:1;background:${C.lightBg};border-radius:8pt;padding:18pt;border-top:4pt solid ${C.blue};display:flex;flex-direction:column">
    <p style="color:${C.blue};font-size:32pt;font-weight:bold;margin:0">02</p>
    <p style="color:${C.dark};font-size:13pt;font-weight:bold;margin:8pt 0 8pt 0">対話で深掘り</p>
    <p style="color:${C.sub};font-size:10pt;margin:0;line-height:1.6">AIとの会話で悩みや志向を深掘りし、一般的なアンケートでは得られない深い理解に基づく個人最適化されたプランを提示</p>
  </div>
  <div style="flex:1;background:${C.lightBg};border-radius:8pt;padding:18pt;border-top:4pt solid ${C.orange};display:flex;flex-direction:column">
    <p style="color:${C.orange};font-size:32pt;font-weight:bold;margin:0">03</p>
    <p style="color:${C.dark};font-size:13pt;font-weight:bold;margin:8pt 0 8pt 0">蓄積が組織価値に</p>
    <p style="color:${C.sub};font-size:10pt;margin:0;line-height:1.6">個人の会話蓄積で従来得られない詳細情報が溜まり、人員配置最適化・メンタリングマッチング等の新ユースケースが生まれる</p>
  </div>
</div>
${end}` });

// Slide 9: アーキテクチャ
slides.push({ name: 'slide09.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">アーキテクチャ概要</h2></div>
<div style="flex:1;display:flex;flex-direction:column;gap:8pt;margin:16pt 40pt 16pt 40pt">
  <div style="display:flex;gap:10pt">
    <div style="width:120pt;background:${C.dark};border-radius:4pt;padding:8pt;display:flex;align-items:center"><p style="color:#fff;font-size:10pt;font-weight:bold;margin:0">ユーザー接点</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:${C.lightBg};border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.border}"><p style="color:${C.dark};font-size:10pt;margin:0;font-weight:bold">Slack</p></div>
      <div style="flex:1;background:${C.lightBg};border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.border}"><p style="color:${C.dark};font-size:10pt;margin:0;font-weight:bold">Microsoft Teams</p></div>
      <div style="flex:1;background:${C.lightBg};border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.border}"><p style="color:${C.sub};font-size:10pt;margin:0">日常ツールで気軽に利用</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:120pt;background:${C.teal};border-radius:4pt;padding:8pt;display:flex;align-items:center"><p style="color:#fff;font-size:10pt;font-weight:bold;margin:0">AIエージェント</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:#E6FFFA;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.teal}"><p style="color:${C.dark};font-size:9pt;margin:0">キャリア壁打ち</p></div>
      <div style="flex:1;background:#E6FFFA;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.teal}"><p style="color:${C.dark};font-size:9pt;margin:0">励まし・通知</p></div>
      <div style="flex:1;background:#E6FFFA;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.teal}"><p style="color:${C.dark};font-size:9pt;margin:0">マッチング</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:120pt;background:${C.blue};border-radius:4pt;padding:8pt;display:flex;align-items:center"><p style="color:#fff;font-size:10pt;font-weight:bold;margin:0">RAG / 検索</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:#EBF4FF;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.blue}"><p style="color:${C.dark};font-size:9pt;margin:0">インタビュー記事</p></div>
      <div style="flex:1;background:#EBF4FF;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.blue}"><p style="color:${C.dark};font-size:9pt;margin:0">公募情報</p></div>
      <div style="flex:1;background:#EBF4FF;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.blue}"><p style="color:${C.dark};font-size:9pt;margin:0">スキルデータ</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:120pt;background:${C.orange};border-radius:4pt;padding:8pt;display:flex;align-items:center"><p style="color:#fff;font-size:10pt;font-weight:bold;margin:0">データ基盤</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:#FFF5F0;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.orange}"><p style="color:${C.dark};font-size:9pt;margin:0">ロールモデルマップ</p></div>
      <div style="flex:1;background:#FFF5F0;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.orange}"><p style="color:${C.dark};font-size:9pt;margin:0">個人キャリアマップ</p></div>
      <div style="flex:1;background:#FFF5F0;border-radius:4pt;padding:8pt;text-align:center;border:1pt solid ${C.orange}"><p style="color:${C.dark};font-size:9pt;margin:0">全社員スキルマップ</p></div>
    </div>
  </div>
  <div style="display:flex;gap:10pt">
    <div style="width:120pt;background:#718096;border-radius:4pt;padding:8pt;display:flex;align-items:center"><p style="color:#fff;font-size:10pt;font-weight:bold;margin:0">社内連携</p></div>
    <div style="flex:1;display:flex;gap:8pt">
      <div style="flex:1;background:${C.lightBg};border-radius:4pt;padding:8pt;text-align:center;border:1pt dashed #718096"><p style="color:${C.sub};font-size:9pt;margin:0">既存キャリアシステム</p></div>
      <div style="flex:1;background:${C.lightBg};border-radius:4pt;padding:8pt;text-align:center;border:1pt dashed #718096"><p style="color:${C.sub};font-size:9pt;margin:0">人事システム</p></div>
      <div style="flex:1;background:${C.lightBg};border-radius:4pt;padding:8pt;text-align:center;border:1pt dashed #718096"><p style="color:${C.sub};font-size:9pt;margin:0">自社データ連携（将来拡張）</p></div>
    </div>
  </div>
</div>
<p style="color:${C.sub};font-size:9pt;margin:0 40pt 12pt 40pt">※ 破線は将来連携を示す。既存システムと簡単に繋ぐことができる拡張性のある設計</p>
${end}` });

// Slide 10: 期待される効果
slides.push({ name: 'slide10.html', html: () => `${base(C.white)}
<div style="background:${C.dark};padding:18pt 40pt"><h2 style="color:#fff;font-size:22pt;margin:0">期待される効果</h2></div>
<div style="flex:1;display:flex;gap:16pt;margin:24pt 40pt 24pt 40pt;align-items:center">
  <div style="flex:1;text-align:center;padding:20pt;background:${C.lightBg};border-radius:8pt;border-bottom:4pt solid ${C.teal}">
    <p style="color:${C.teal};font-size:14pt;font-weight:bold;margin:0 0 8pt 0">離職率</p>
    <p style="color:${C.dark};font-size:11pt;margin:0">「可能性が見えない」離職の削減</p>
  </div>
  <div style="flex:1;text-align:center;padding:20pt;background:${C.lightBg};border-radius:8pt;border-bottom:4pt solid ${C.blue}">
    <p style="color:${C.blue};font-size:14pt;font-weight:bold;margin:0 0 8pt 0">エンゲージメント</p>
    <p style="color:${C.dark};font-size:11pt;margin:0">キャリア展望の明確化による向上</p>
  </div>
  <div style="flex:1;text-align:center;padding:20pt;background:${C.lightBg};border-radius:8pt;border-bottom:4pt solid ${C.orange}">
    <p style="color:${C.orange};font-size:14pt;font-weight:bold;margin:0 0 8pt 0">人事工数</p>
    <p style="color:${C.dark};font-size:11pt;margin:0">人員配置・マッチングの自動化</p>
  </div>
  <div style="flex:1;text-align:center;padding:20pt;background:${C.lightBg};border-radius:8pt;border-bottom:4pt solid #9F7AEA">
    <p style="color:#9F7AEA;font-size:14pt;font-weight:bold;margin:0 0 8pt 0">組織流動性</p>
    <p style="color:${C.dark};font-size:11pt;margin:0">異動ハードル低下、適材適所の実現</p>
  </div>
</div>
${end}` });

// Slide 11: クロージング
slides.push({ name: 'slide11.html', html: () => `<!DOCTYPE html><html><head><style>
html{background:#fff}body{width:720pt;height:405pt;margin:0;padding:0;font-family:Arial,sans-serif;display:flex;flex-direction:column;background-image:url('${path.join(htmlDir,"bg_closing.png")}');background-size:cover}
</style></head><body>
<div style="flex:1;display:flex;flex-direction:column;justify-content:center;align-items:center;margin:40pt">
  <div style="border-left:6pt solid ${C.teal};padding-left:20pt;margin-bottom:30pt">
    <h1 style="color:#fff;font-size:30pt;margin:0 0 4pt 0">半径5mの外に、</h1>
  <h1 style="color:#fff;font-size:30pt;margin:0">あなたのキャリアがある</h1>
  </div>
  <p style="color:#ffffffcc;font-size:14pt;margin:0 0 8pt 0">YourCompany Career Nexus</p>
  <p style="color:#ffffffaa;font-size:11pt;margin:0">すべての社員に"自社の中のまだ見ぬ可能性"を届ける</p>
  <p style="color:#ffffffaa;font-size:11pt;margin:4pt 0 0 0">個人のキャリア実現と組織の活性化を同時に叶えるAIエージェント</p>
</div>
${end}` });

async function main() {
  // Create gradient backgrounds
  await createGradient(path.join(htmlDir, 'bg_title.png'), '#1B2A4A', '#0D1B2A');
  await createGradient(path.join(htmlDir, 'bg_closing.png'), '#1B2A4A', '#162033');

  // Write HTML files
  for (const s of slides) {
    fs.writeFileSync(path.join(htmlDir, s.name), s.html());
  }

  // Generate PPTX
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'YourCompany Career Nexus - Detailed';
  pptx.author = 'Career Nexus Team';

  for (const s of slides) {
    await html2pptx(path.join(htmlDir, s.name), pptx);
  }

  const outFile = path.join(outDir, 'career_nexus_detailed.pptx');
  await pptx.writeFile({ fileName: outFile });
  console.log('Created:', outFile);
}

main().catch(console.error);
