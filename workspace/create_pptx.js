const pptxgen = require('pptxgenjs');
const path = require('path');

async function createPresentation() {
  const pptx = new pptxgen();
  pptx.defineLayout({ name: 'CUSTOM', width: 17.778, height: 10 });
  pptx.layout = 'CUSTOM';
  pptx.author = 'Career Nexus AI';
  pptx.title = 'YourCompany Career Nexus';

  const imgDir = path.resolve(__dirname);
  const BG = 'F0EFEA'; // slide background color
  const WHITE = 'FFFFFF';

  // ============================================================
  // Slide 1: Title & Hook
  // ============================================================
  const s1 = pptx.addSlide();
  s1.background = { path: path.join(imgDir, 'image1.png') };
  s1.addText('あなたのキャリアの選択肢、"半径5m"で決めていませんか？', {
    x: 0.3, y: 6.0, w: 17.2, h: 1.5,
    fontSize: 40, fontFace: 'Arial', color: '1A1A1A', bold: true,
    align: 'center', valign: 'middle',
    fill: { color: BG }
  });
  s1.addText('[YourCompany] Career Nexus | 自社の可能性を、自分の可能性にするAIエージェント', {
    x: 1.5, y: 7.6, w: 14.8, h: 0.9,
    fontSize: 22, fontFace: 'Arial', color: '666666',
    align: 'center', valign: 'middle',
    fill: { color: BG }
  });

  // ============================================================
  // Slide 2: Problem
  // ============================================================
  const s2 = pptx.addSlide();
  s2.background = { path: path.join(imgDir, 'image2.png') };

  // Left labels
  s2.addText('入社', { x: 0.2, y: 1.6, w: 1.3, h: 0.5, fontSize: 16, color: 'FFFFFF', bold: true, align: 'center', fill: { color: '555555' } });
  s2.addText('1年', { x: 2.3, y: 1.6, w: 1.0, h: 0.5, fontSize: 16, color: 'FFFFFF', bold: true, align: 'center', fill: { color: '555555' } });
  s2.addText('3年', { x: 4.0, y: 1.6, w: 1.0, h: 0.5, fontSize: 16, color: 'FFFFFF', bold: true, align: 'center', fill: { color: '555555' } });

  // Card 1: 大企業の深刻な流出
  s2.addText([
    { text: '     大企業の深刻な流出\n\n', options: { fontSize: 20, bold: true, color: '1A1A1A' } },
    { text: '従業員1,000人以上の大企業でも、入社3\n年以内で', options: { fontSize: 14, color: '333333' } },
    { text: '約25〜27%', options: { fontSize: 14, color: 'D4A017', bold: true } },
    { text: 'が離職。\n（※日本全体の大卒3年以内離職\n率は', options: { fontSize: 14, color: '333333' } },
    { text: '約32〜34%', options: { fontSize: 14, color: 'D4A017', bold: true } },
    { text: '）', options: { fontSize: 14, color: '333333' } }
  ], {
    x: 6.2, y: 0.5, w: 5.8, h: 2.8,
    valign: 'top',
    fill: { color: 'FDF6E3' },
    line: { color: 'D4A017', width: 1.5 }
  });

  // Card 2: 隠れた巨額損失
  s2.addText([
    { text: '     隠れた巨額損失\n\n', options: { fontSize: 20, bold: true, color: '1A1A1A' } },
    { text: '新卒1人の離職損失は', options: { fontSize: 14, color: '333333' } },
    { text: '約150〜200万円', options: { fontSize: 14, color: '2596BE', bold: true } },
    { text: '。\n年間100人が離職する企業では、', options: { fontSize: 14, color: '333333' } },
    { text: '年間\n1.5〜2億円以上', options: { fontSize: 14, color: '2596BE', bold: true } },
    { text: 'の直接的な財務損失に。', options: { fontSize: 14, color: '333333' } }
  ], {
    x: 6.2, y: 3.5, w: 5.8, h: 2.6,
    valign: 'top',
    fill: { color: 'EBF5FB' },
    line: { color: '2596BE', width: 1.5 }
  });

  // Card 3: エンゲージメントの危機
  s2.addText([
    { text: '     エンゲージメントの危機\n\n', options: { fontSize: 20, bold: true, color: '1A1A1A' } },
    { text: '日本の「熱意ある社員」はわずか', options: { fontSize: 14, color: '333333' } },
    { text: '5〜6%', options: { fontSize: 14, color: 'D4A017', bold: true } },
    { text: '\n（世界最低水準。世界平均は', options: { fontSize: 14, color: '333333' } },
    { text: '23%', options: { fontSize: 14, color: 'D4A017', bold: true } },
    { text: '）。', options: { fontSize: 14, color: '333333' } }
  ], {
    x: 6.2, y: 6.3, w: 5.8, h: 2.3,
    valign: 'top',
    fill: { color: 'FDF6E3' },
    line: { color: 'D4A017', width: 1.5 }
  });

  // Right side reason boxes
  const reasonBoxOpts = { align: 'center', valign: 'middle', fill: { color: 'F0F4F8' }, line: { color: 'CCCCCC', width: 1 } };
  s2.addText([
    { text: '1. 業務内容の\nミスマッチ\n（', options: { fontSize: 17, color: '333333', bold: true } },
    { text: '約25〜30%', options: { fontSize: 17, color: 'D4A017', bold: true } },
    { text: '）', options: { fontSize: 17, color: '333333', bold: true } }
  ], { x: 13.0, y: 0.5, w: 4.3, h: 2.5, ...reasonBoxOpts });

  s2.addText([
    { text: '2. キャリアの見通し・\n成長実感の欠如\n（', options: { fontSize: 17, color: '333333', bold: true } },
    { text: '約25〜28%', options: { fontSize: 17, color: 'D4A017', bold: true } },
    { text: '）', options: { fontSize: 17, color: '333333', bold: true } }
  ], { x: 13.0, y: 3.3, w: 4.3, h: 2.5, ...reasonBoxOpts });

  s2.addText([
    { text: '3. 局所的な\n人間関係への依存\n（', options: { fontSize: 17, color: '333333', bold: true } },
    { text: '約20〜25%', options: { fontSize: 17, color: 'D4A017', bold: true } },
    { text: '）', options: { fontSize: 17, color: '333333', bold: true } }
  ], { x: 13.0, y: 6.1, w: 4.3, h: 2.5, ...reasonBoxOpts });

  // ============================================================
  // Slide 3: Root Cause
  // ============================================================
  const s3 = pptx.addSlide();
  s3.background = { path: path.join(imgDir, 'image3.png') };
  s3.addText('辞める本当の理由:「会社が嫌い」ではなく、「この会社での可能性が見えない」', {
    x: 0.3, y: 0.15, w: 17.0, h: 1.0,
    fontSize: 28, color: '1A1A1A', bold: true, fill: { color: BG }
  });
  s3.addText([
    { text: 'ロールモデルは3フロア上にいるかもしれないのに、出会う手段がない。\n', options: { fontSize: 19, color: '333333', bold: true } },
    { text: '大企業特有の「サイロ化」が、"ここにいても成長できない"という誤解を生んでいる。', options: { fontSize: 19, color: '333333', bold: true } }
  ], {
    x: 0.8, y: 1.3, w: 16.0, h: 1.3,
    align: 'center', fill: { color: BG }
  });
  s3.addText('現状', { x: 2.0, y: 2.8, w: 5.5, h: 0.7, fontSize: 24, color: '1A1A1A', bold: true, align: 'center', fill: { color: BG } });
  s3.addText('見えていない可能性', { x: 9.5, y: 2.8, w: 7.5, h: 0.7, fontSize: 24, color: '1A1A1A', bold: true, align: 'center', fill: { color: BG } });
  s3.addText('直属の上司・先輩の\n将来が自分の将来に\n見えてしまう…', {
    x: 0.8, y: 4.2, w: 4.8, h: 2.0, fontSize: 16, color: 'FFFFFF', bold: true, align: 'center',
    fill: { color: '444444', transparency: 30 }
  });
  s3.addText('上司', { x: 5.2, y: 7.0, w: 2.0, h: 0.5, fontSize: 14, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('従業員', { x: 2.2, y: 8.3, w: 2.0, h: 0.5, fontSize: 14, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('他部署の\nリーダー', { x: 9.2, y: 3.6, w: 2.2, h: 0.8, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('スペシャリスト', { x: 13.0, y: 3.6, w: 2.5, h: 0.5, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('メンター', { x: 8.3, y: 5.6, w: 2.0, h: 0.5, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('プロジェクト\nマネージャー', { x: 14.5, y: 5.3, w: 2.5, h: 0.8, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('海外拠点', { x: 8.8, y: 8.2, w: 2.0, h: 0.5, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('新規事業部', { x: 11.3, y: 8.7, w: 2.5, h: 0.5, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s3.addText('研究開発', { x: 14.2, y: 7.8, w: 2.0, h: 0.5, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });

  // ============================================================
  // Slide 4: Comparison Table
  // ============================================================
  const s4 = pptx.addSlide();
  s4.background = { path: path.join(imgDir, 'image4.png') };
  s4.addText('既存の施策が抱える「3つの構造的ギャップ」', {
    x: 0.5, y: 0.2, w: 16.5, h: 1.2, fontSize: 34, color: '1A1A1A', bold: true, align: 'center', fill: { color: BG }
  });
  // Headers
  s4.addText('1on1面談', { x: 5.0, y: 1.7, w: 2.8, h: 0.9, fontSize: 18, color: WHITE, bold: true, align: 'center', valign: 'middle', fill: { color: '666666' } });
  s4.addText('社内公募', { x: 7.9, y: 1.7, w: 2.8, h: 0.9, fontSize: 18, color: WHITE, bold: true, align: 'center', valign: 'middle', fill: { color: '666666' } });
  s4.addText('外部キャリア\nコーチング', { x: 10.8, y: 1.6, w: 2.8, h: 1.1, fontSize: 18, color: WHITE, bold: true, align: 'center', valign: 'middle', fill: { color: '666666' } });
  s4.addText('Career Nexus', { x: 13.8, y: 1.7, w: 3.6, h: 0.9, fontSize: 22, color: '2596BE', bold: true, align: 'center', valign: 'middle', fill: { color: 'E8F6FC' }, line: { color: '2596BE', width: 2 } });

  // Row labels
  s4.addText('1. 情報網の広さ\n（情報のギャップ）：\n大企業内の多様な経験を\n知れるか？', {
    x: 0.2, y: 2.8, w: 4.7, h: 2.5, fontSize: 15, color: '1A1A1A', bold: true, valign: 'middle', fill: { color: 'F5F5F0' }
  });
  s4.addText('✕\n\n（上司の知見・視野\nに依存）', {
    x: 5.0, y: 2.8, w: 2.8, h: 2.5, fontSize: 14, color: '333333', align: 'center', valign: 'middle', fill: { color: WHITE }
  });
  s4.addText('', { x: 7.9, y: 2.8, w: 2.8, h: 2.5, fill: { color: WHITE } });
  s4.addText('', { x: 10.8, y: 2.8, w: 2.8, h: 2.5, fill: { color: WHITE } });
  s4.addText('◎\n\n（全社網羅）', {
    x: 13.8, y: 2.8, w: 3.6, h: 2.5, fontSize: 15, color: '2596BE', align: 'center', valign: 'middle', bold: true, fill: { color: 'F0FAFF' }
  });

  s4.addText('2. 潜在ニーズの発掘\n（マッチングのギャップ）：\n表面化していないポジション\nや志向に気づけるか？', {
    x: 0.2, y: 5.4, w: 4.7, h: 2.5, fontSize: 15, color: '1A1A1A', bold: true, valign: 'middle', fill: { color: 'EAEAE5' }
  });
  s4.addText('', { x: 5.0, y: 5.4, w: 2.8, h: 2.5, fill: { color: 'F5F5F0' } });
  s4.addText('△\n\n（公募は氷山の一角。\n能動的な層にしか届かない）', {
    x: 7.9, y: 5.4, w: 2.8, h: 2.5, fontSize: 12, color: '333333', align: 'center', valign: 'middle', fill: { color: 'F5F5F0' }
  });
  s4.addText('', { x: 10.8, y: 5.4, w: 2.8, h: 2.5, fill: { color: 'F5F5F0' } });
  s4.addText('◎\n\n（潜在層へのアプローチ）', {
    x: 13.8, y: 5.4, w: 3.6, h: 2.5, fontSize: 15, color: '2596BE', align: 'center', valign: 'middle', bold: true, fill: { color: 'F0FAFF' }
  });

  s4.addText('3. 自社固有の実態把握\n（対話のギャップ）：\n自社の風土や制度に基づく\n壁打ちができるか？', {
    x: 0.2, y: 8.0, w: 4.7, h: 1.9, fontSize: 15, color: '1A1A1A', bold: true, valign: 'middle', fill: { color: 'F5F5F0' }
  });
  s4.addText('', { x: 5.0, y: 8.0, w: 2.8, h: 1.9, fill: { color: WHITE } });
  s4.addText('', { x: 7.9, y: 8.0, w: 2.8, h: 1.9, fill: { color: WHITE } });
  s4.addText('✕\n\n（社内の人脈やリアルな\n風土を知らない一般論）', {
    x: 10.8, y: 8.0, w: 2.8, h: 1.9, fontSize: 12, color: '333333', align: 'center', valign: 'middle', fill: { color: WHITE }
  });
  s4.addText('◎\n\n（自社データ特化）', {
    x: 13.8, y: 8.0, w: 3.6, h: 1.9, fontSize: 15, color: '2596BE', align: 'center', valign: 'middle', bold: true, fill: { color: 'F0FAFF' }
  });

  // ============================================================
  // Slide 5: Solution - 3 Phases
  // ============================================================
  const s5 = pptx.addSlide();
  s5.background = { path: path.join(imgDir, 'image5.png') };
  s5.addText('誰もが社内の優秀な社員を\nロールモデルにできる仕組み', {
    x: 0.3, y: 0.3, w: 9.0, h: 2.2, fontSize: 34, color: '1A1A1A', bold: true, fill: { color: BG }
  });
  s5.addText('3つのフェーズで進化する\n「AIキャリアエージェント」', {
    x: 0.3, y: 2.7, w: 8.0, h: 1.4, fontSize: 22, color: '333333', fill: { color: BG }
  });

  // Phase cards
  const phaseCardOpts = { valign: 'top', fill: { color: WHITE }, shadow: { type: 'outer', blur: 6, offset: 2, color: '999999', opacity: 0.3 } };
  s5.addText([
    { text: 'Phase 1:\n発見 (Discovery)\n\n', options: { fontSize: 22, bold: true, color: '1A1A1A' } },
    { text: 'ロールモデルマップ構築と、自社\nデータに基づくキャリア壁打ち。', options: { fontSize: 15, color: '555555' } }
  ], { x: 0.5, y: 5.0, w: 5.2, h: 4.0, ...phaseCardOpts });

  s5.addText([
    { text: 'Phase 2:\n継続 (Continuation)\n\n', options: { fontSize: 22, bold: true, color: '1A1A1A' } },
    { text: '個人別キャリアマップの構築と、AI\nからの自発的なモチベーション支援。', options: { fontSize: 15, color: '555555' } }
  ], { x: 6.1, y: 5.0, w: 5.2, h: 4.0, ...phaseCardOpts });

  s5.addText([
    { text: 'Phase 3:\n組織 (Organization)\n\n', options: { fontSize: 22, bold: true, color: '1A1A1A' } },
    { text: '全社スキルの可視化と、最適な人員\n配置・自動メンタリングマッチング。', options: { fontSize: 15, color: '555555' } }
  ], { x: 11.7, y: 5.0, w: 5.2, h: 4.0, ...phaseCardOpts });

  // ============================================================
  // Slide 6: User Experience Flow
  // ============================================================
  const s6 = pptx.addSlide();
  s6.background = { path: path.join(imgDir, 'image6.png') };
  s6.addText('ユーザー体験：今の仕事の延長しか見えない若手社員のケース', {
    x: 0.3, y: 0.15, w: 17.0, h: 1.0, fontSize: 28, color: '1A1A1A', bold: true, fill: { color: BG }
  });

  const stepW = 3.15;
  const steps = [
    { title: 'Step 1: 気軽な相談', body: '「今後のキャリアに悩\nんでいて...」\nいつも使うチャットツ\nールからAIに相談。', bg: WHITE },
    { title: 'Step 2: 意外な発見', body: '「あなたと似た経歴か\nら、全く異なる部署で\n活躍中の○○さんがい\nます。」\n潜在的なロールモデル\nをAIが提示。', bg: 'E8F8F8' },
    { title: 'Step 3:\n疑似対話で深掘り', body: 'ロールモデルの考え\n方・ターニングポイン\nトをAIとの対話でシミ\nュレーション。', bg: WHITE },
    { title: 'Step 4:\n自社ならではの提案', body: '自社の制度や風土を反\n映した、現実的なキャ\nリアパスの提示。', bg: WHITE },
    { title: 'Step 5: アクション', body: '関連する社内公募の通\n知や、社内勉強会・メ\nンターの紹介へシーム\nレスに接続。', bg: 'FDF6E3' }
  ];
  steps.forEach((s, i) => {
    const x = 0.3 + i * 3.5;
    s6.addText(s.title, { x, y: 4.3, w: stepW, h: 0.9, fontSize: 15, color: '1A1A1A', bold: true, fill: { color: BG } });
    s6.addText(s.body, { x, y: 5.3, w: stepW, h: 2.8, fontSize: 13, color: '333333', valign: 'top', fill: { color: s.bg } });
  });

  // ============================================================
  // Slide 7: Differentiation - RAG
  // ============================================================
  const s7 = pptx.addSlide();
  s7.background = { path: path.join(imgDir, 'image7.png') };
  s7.addText('汎用AIとの決定的な違い：「自社のリアル」を学習したアシスト', {
    x: 0.3, y: 0.15, w: 17.0, h: 1.0, fontSize: 28, color: '1A1A1A', bold: true, fill: { color: BG }
  });
  s7.addText('インターネット上の一般的なアドバイスではなく、自社のキャリア制度や各部署の風土、先輩社員の\nリアルなターニングポイントに基づいた、極めて解像度の高い提案を実現します。', {
    x: 0.3, y: 1.3, w: 17.0, h: 1.0, fontSize: 16, color: '333333', fill: { color: BG }
  });
  s7.addText('インタビュー記事', { x: 2.5, y: 3.6, w: 3.5, h: 0.5, fontSize: 16, color: '333333', fill: { color: BG } });
  s7.addText('スキルデータ', { x: 2.5, y: 5.2, w: 3.0, h: 0.5, fontSize: 16, color: '333333', fill: { color: BG } });
  s7.addText('公募情報', { x: 2.5, y: 6.7, w: 2.5, h: 0.5, fontSize: 16, color: '333333', fill: { color: BG } });
  s7.addText('社内データ', { x: 1.2, y: 8.3, w: 3.8, h: 0.5, fontSize: 16, color: '555555', align: 'center', fill: { color: BG } });
  s7.addText('自社データ特化\nAI処理 (RAG)', { x: 7.2, y: 7.8, w: 4.0, h: 1.0, fontSize: 18, color: '333333', align: 'center', bold: true, fill: { color: BG } });
  s7.addText('ロールモデルマップ\n(Role Model Map)', { x: 12.5, y: 8.2, w: 5.0, h: 1.0, fontSize: 18, color: '333333', align: 'center', bold: true, fill: { color: BG } });
  s7.addText('キャリアパス', { x: 13.2, y: 2.8, w: 2.5, h: 0.5, fontSize: 12, color: '333333', fill: { color: BG } });
  s7.addText('ロールモデル', { x: 15.5, y: 2.8, w: 2.0, h: 0.5, fontSize: 12, color: '333333', fill: { color: BG } });
  s7.addText('ターニング\nポイント', { x: 15.8, y: 4.3, w: 1.8, h: 0.8, fontSize: 12, color: '333333', fill: { color: BG } });
  s7.addText('メンター', { x: 15.8, y: 5.8, w: 1.8, h: 0.5, fontSize: 12, color: '333333', fill: { color: BG } });
  s7.addText('ターニング\nポイント', { x: 11.2, y: 6.8, w: 1.8, h: 0.8, fontSize: 12, color: '333333', fill: { color: BG } });
  s7.addText('キャリアパス', { x: 12.8, y: 7.3, w: 2.0, h: 0.5, fontSize: 12, color: '333333', fill: { color: BG } });

  // ============================================================
  // Slide 8: Phase 2 & 3 Circular
  // ============================================================
  const s8 = pptx.addSlide();
  s8.background = { path: path.join(imgDir, 'image8.png') };
  s8.addText('個人の「対話」の蓄積が、組織全体の「流動性」に変わる', {
    x: 0.5, y: 0.15, w: 16.5, h: 1.0, fontSize: 30, color: '1A1A1A', bold: true, align: 'center', fill: { color: BG }
  });

  s8.addText('Phase 2（継続・支援）', { x: 0.2, y: 3.8, w: 4.8, h: 0.6, fontSize: 18, color: '1A1A1A', bold: true, fill: { color: BG } });
  s8.addText('日々の会話からモチベーション\nを把握し、AIが「あの勉強会が\nありますよ」「同じ悩みを乗り\n越えた○○さんの話を聞きま\nせんか？」と自発的に支援。', {
    x: 0.2, y: 4.6, w: 4.8, h: 2.8, fontSize: 14, color: '333333', fill: { color: BG }
  });

  s8.addText('Phase 3（組織への還元）', { x: 12.8, y: 3.8, w: 4.8, h: 0.6, fontSize: 18, color: '1A1A1A', bold: true, fill: { color: BG } });
  s8.addText('個人ごとの対話データが全社\nのスキルマップとして結実。\n異動のハードルを下げ、人と\n人を繋ぐ自動マッチングで\n真の「適材適所」を実現。', {
    x: 12.8, y: 4.6, w: 4.8, h: 2.8, fontSize: 14, color: '333333', fill: { color: BG }
  });

  // Circle labels
  s8.addText('会話データの\n蓄積', { x: 3.8, y: 1.3, w: 2.5, h: 0.8, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s8.addText('AIとの対話', { x: 6.5, y: 2.6, w: 2.5, h: 0.5, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s8.addText('動機づけ\n・声かけ', { x: 8.8, y: 3.3, w: 2.0, h: 0.8, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s8.addText('自発的な\nアクション', { x: 6.0, y: 5.6, w: 2.5, h: 0.8, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s8.addText('動的な\n組織への進化', { x: 3.5, y: 7.0, w: 2.8, h: 0.8, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s8.addText('最適な人員配置', { x: 6.8, y: 8.3, w: 3.5, h: 0.5, fontSize: 14, color: '333333', align: 'center', bold: true, fill: { color: BG } });
  s8.addText('スキル・\nモチベーション\nの可視化', { x: 11.2, y: 1.3, w: 2.8, h: 1.0, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });
  s8.addText('勉強会・\nメンタリングの\n自動マッチング', { x: 11.5, y: 6.3, w: 3.2, h: 1.0, fontSize: 13, color: '333333', align: 'center', fill: { color: BG } });

  // ============================================================
  // Slide 9: Architecture
  // ============================================================
  const s9 = pptx.addSlide();
  s9.background = { path: path.join(imgDir, 'image9.png') };
  s9.addText('既存の業務環境に溶け込む、拡張性の高いアーキテクチャ', {
    x: 0.3, y: 0.15, w: 17.0, h: 1.0, fontSize: 28, color: '1A1A1A', bold: true, fill: { color: BG }
  });
  s9.addText('特別な専用アプリは不要。社員が毎日開く日常ツールの中で完結。既存の人事システムやキャリア\nデータと容易に連携可能。将来的な機能拡張にも柔軟に対応するデータ基盤。', {
    x: 0.3, y: 1.3, w: 17.0, h: 1.0, fontSize: 16, color: '333333', fill: { color: BG }
  });
  s9.addText('ユーザー接点\n[Slack / Microsoft Teams]', { x: 10.2, y: 3.1, w: 7.0, h: 1.0, fontSize: 17, color: '333333', fill: { color: BG } });
  s9.addText('AIエージェント群\nキャリア壁打ち / 励まし・通知 / マッチング', { x: 10.2, y: 4.6, w: 7.0, h: 1.0, fontSize: 17, color: '333333', fill: { color: BG } });
  s9.addText('データ基盤 & RAG検索\nロールモデルマップ / 個人キャリアマップ /\nスキルマップのベクトル検索', { x: 10.2, y: 6.1, w: 7.0, h: 1.3, fontSize: 17, color: '333333', fill: { color: BG } });
  s9.addText('社内連携\n既存キャリアシステム /\n人事システム / 自社データ', { x: 10.2, y: 7.8, w: 7.0, h: 1.2, fontSize: 17, color: '333333', fill: { color: BG } });

  // ============================================================
  // Slide 10: Why Career Nexus
  // ============================================================
  const s10 = pptx.addSlide();
  s10.background = { path: path.join(imgDir, 'image10.png') };
  s10.addText('なぜ Career Nexus で\nでなければならないか？', {
    x: 0.3, y: 0.3, w: 8.5, h: 2.2, fontSize: 34, color: '1A1A1A', bold: true, fill: { color: BG }
  });

  const cardOpts10 = { align: 'center', valign: 'top', fill: { color: WHITE } };
  s10.addText([
    { text: '自社データ特化\n(Company-Specific)\n\n', options: { fontSize: 18, bold: true, color: '1A1A1A' } },
    { text: 'アンケートや一般論ではなく、\n「自社ならでは」の文脈・風土・\nロールモデルに基づく具体的で\nリアルな提案。', options: { fontSize: 14, color: '555555' } }
  ], { x: 0.5, y: 5.5, w: 5.3, h: 3.5, ...cardOpts10 });

  s10.addText([
    { text: '寄り添う対話\n(Empathetic Dialogue)\n\n', options: { fontSize: 18, bold: true, color: '1A1A1A' } },
    { text: '一方的な情報提供ではなく、対\n話を通じて個人の志向や悩みを\n深掘り。潜在的なニーズを引き\n出すパーソナライズ。', options: { fontSize: 14, color: '555555' } }
  ], { x: 6.2, y: 5.5, w: 5.3, h: 3.5, ...cardOpts10 });

  s10.addText([
    { text: '組織への還元\n(Organizational ROI)\n\n', options: { fontSize: 18, bold: true, color: '1A1A1A' } },
    { text: '個人の対話がシステムに蓄積す\nるほど、組織全体の人員配置や\nマッチング精度が向上するデー\nタループの構築。', options: { fontSize: 14, color: '555555' } }
  ], { x: 11.9, y: 5.5, w: 5.3, h: 3.5, ...cardOpts10 });

  // ============================================================
  // Slide 11: Expected Effects
  // ============================================================
  const s11 = pptx.addSlide();
  s11.background = { path: path.join(imgDir, 'image11.png') };
  s11.addText('期待される効果：組織のサイロ化を打破するROI', {
    x: 0.3, y: 0.15, w: 17.0, h: 1.0, fontSize: 30, color: '1A1A1A', bold: true, fill: { color: BG }
  });

  const effOpts = { valign: 'top', fill: { color: WHITE } };
  s11.addText([
    { text: '[離職率低下] Retention\n\n', options: { fontSize: 20, color: '2596BE', bold: true } },
    { text: '「可能性が見えない」という理由に\nよる、若手・優秀層の致命的な離職\nをストップ。', options: { fontSize: 15, color: '333333' } }
  ], { x: 0.5, y: 1.5, w: 8.0, h: 3.2, ...effOpts });

  s11.addText([
    { text: '[エンゲージメント向上] Engagement\n\n', options: { fontSize: 20, color: '2596BE', bold: true } },
    { text: '自社内での自律的なキャリア展望が\n明確になることによる、抜本的な\nモチベーション回復。', options: { fontSize: 15, color: '333333' } }
  ], { x: 9.2, y: 1.5, w: 8.0, h: 3.2, ...effOpts });

  s11.addText([
    { text: '[人事工数削減] HR Efficiency\n\n', options: { fontSize: 20, color: '2596BE', bold: true } },
    { text: '手作業に頼っていた人員配置の\n検討や、社内メンター制度・勉強会\nのマッチングを自動化。', options: { fontSize: 15, color: '333333' } }
  ], { x: 0.5, y: 5.5, w: 8.0, h: 3.2, ...effOpts });

  s11.addText([
    { text: '[組織流動性] Fluidity\n\n', options: { fontSize: 20, color: '2596BE', bold: true } },
    { text: '部署間の壁を越えた人材の交わりを\n促し、大企業の中に「動的な組織」\nと真の適材適所を実現。', options: { fontSize: 15, color: '333333' } }
  ], { x: 9.2, y: 5.5, w: 8.0, h: 3.2, ...effOpts });

  // ============================================================
  // Slide 12: Closing
  // ============================================================
  const s12 = pptx.addSlide();
  s12.background = { path: path.join(imgDir, 'image12.png') };
  s12.addText('「半径5mの外に、あなたのキャリアがある。」', {
    x: 0.5, y: 2.8, w: 16.5, h: 1.5, fontSize: 36, color: '1A1A1A', bold: true, align: 'center', fill: { color: BG, transparency: 20 }
  });
  s12.addText('[YourCompany] Career Nexus が、自社の中のまだ見ぬ可能性を届け、\n個人と会社がいきいきと働く世界を実現します。', {
    x: 1.5, y: 4.5, w: 14.5, h: 1.5, fontSize: 21, color: '444444', align: 'center', fill: { color: BG, transparency: 20 }
  });

  const outputPath = path.resolve(__dirname, '..', 'output', 'Career_Nexus_AI_Editable.pptx');
  await pptx.writeFile({ fileName: outputPath });
  console.log('Saved:', outputPath);
}

createPresentation().catch(console.error);
