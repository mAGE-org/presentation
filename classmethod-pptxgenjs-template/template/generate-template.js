const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");

const { COLORS, FONTS, SIZE, SLIDE, LAYOUT, makeShadow, makeCardShadow } = require("./theme");

// ---------------------------------------------------------------------------
// Icon helper
// ---------------------------------------------------------------------------
function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, "#" + color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

function classmethodLogoBase64() {
  const svg = `<svg width="160" height="35" fill="none" xmlns="http://www.w3.org/2000/svg">
  <g clip-path="url(#a)" fill="#000">
    <path d="M156.26 1.84689c1.042 0 1.87.81747 1.87 1.84688V30.9429c0 1.2111-1.012 2.2102-2.238 2.2102H4.10807c-1.22629 0-2.23798-.9991-2.23798-2.2102V4.05709c0-1.21107 1.01169-2.2102 2.23798-2.2102H132.899c.092 0 .215.03027.276.1211l4.997 5.48011c.736.72664 1.901.72664 2.637 0l4.997-5.48011c.061-.06055.153-.1211.276-.1211h10.178ZM156.26 0h-10.945c-.122 0-.214.0605536-.306.121107L140.134 5.4801l-.031.03028c-.367.33305-.95.33305-1.287 0l-.031-.03028-4.844-5.328716C133.849.0605536 133.757 0 133.635 0H4.10807C1.83943 0 0 1.81661 0 4.05709V30.9732c0 2.2405 1.83943 4.0571 4.10807 4.0571H155.923c2.268 0 4.108-1.8166 4.108-4.0571V3.72405C160 1.66522 158.314 0 156.26 0Z"/>
    <path d="M17.5659 13.9883c.3066.0606.4599.1817.4599.5147v2.0588c0 .2725-.1533.4239-.4599.4239h-4.3227c-2.4219 0-2.8204.8175-2.8204 4.1177 0 3.2699.3678 4.1176 2.8204 4.1176h4.3227c.3066 0 .4599.1211.4599.3936v2.0891c0 .3028-.1533.4542-.4599.4844-1.4409.3028-3.3416.3634-4.3227.3634-5.6409 0-6.7139-1.6955-6.7139-7.4481 0-5.7829 1.073-7.4481 6.7139-7.4481.9811 0 2.8818.0605 4.3227.333ZM23.299 10.3848c.184 0 .3679.1513.3679.3633v17.4394c0 .212-.1839.3634-.3679.3634h-3.127c-.2759 0-.3986-.1817-.3986-.3936V10.7784c0-.2725.184-.3936.3679-.3936h3.1577ZM99.8502 20.9217v.8175c0 .5147-.1533.6963-.8584.6963h-8.3081c.0613 2.3616.6744 2.7855 3.0044 2.7855h4.7212c.1839 0 .3679.1514.3679.3633v2.18c0 .2422-.1226.3936-.3679.4238-1.4102.3028-3.311.3634-5.0584.3634-5.3344 0-6.6527-1.4836-6.6527-7.4481 0-5.9948 1.2876-7.4481 6.6527-7.4481 5.181 0 6.4686 1.4533 6.4993 7.2664Zm-9.1972-.9991h5.2424c-.0306-2.3011-.6438-2.9066-2.5752-2.9066-1.9927 0-2.6365.6055-2.6672 2.9066ZM108.067 25.2204h-1.165c-1.165 0-1.564-.2119-1.564-1.3322v-6.8728h2.514c.184 0 .368-.1211.368-.3936V14.018c0-.1514-.123-.3633-.368-.3633h-2.514v-2.9066c0-.1817-.153-.3633-.368-.3633h-3.188c-.215 0-.368.1816-.368.3633v14.442c0 3.3002 2.054 3.3608 4.353 3.3608.583 0 1.319-.0606 2.238-.2423.307-.0605.43-.1816.43-.4844V25.614c0-.2422-.154-.3936-.368-.3936ZM138.632 21.1034c0 5.9645-1.257 7.4481-6.745 7.4481-5.487 0-6.744-1.4836-6.744-7.4481 0-5.9948 1.257-7.4481 6.744-7.4481 5.518 0 6.745 1.4533 6.745 7.4481Zm-9.535 0c0 3.3001.399 4.1176 2.821 4.1176s2.82-.8175 2.82-4.1176c0-3.3002-.398-4.1177-2.82-4.1177-2.453.0303-2.821.8175-2.821 4.1177ZM119.593 13.8363c-2.023-.3633-4.261-.1816-5.641 0v-3.0882c0-.2422-.214-.3633-.367-.3633h-3.189c-.245 0-.368.2119-.368.3633v17.4394c0 .212.154.3936.368.3936h3.189c.184 0 .367-.1514.367-.3633V21.133c0-2.7854.092-4.1176 2.821-4.1176 2.728 0 2.851 1.3019 2.82 4.1176-.03 2.8461 0 7.0848 0 7.0848 0 .2422.184.3633.368.3633h3.189c.153 0 .367-.1211.367-.3633v-7.8114c0-4.2388-1.011-6.0251-3.924-6.5701ZM153.072 10.2637c.276 0 .398.2119.398.3633v16.3798c0 .7569-.306.9991-1.103 1.1505-1.135.2119-3.434.3936-5.273.3936-5.702 0-6.867-1.5441-6.867-7.4178 0-5.8132 1.165-7.4784 6.438-7.4784 1.839 0 1.9.0605 2.881.1514V10.627c0-.2725.215-.3936.368-.3936h3.158v.0303Zm-5.917 6.782c-2.667 0-3.004.6963-3.004 4.0874 0 3.0882.214 4.1176 2.851 4.1176.644 0 1.747-.0606 2.544-.1817v-8.0233h-2.391ZM37.8608 18.1665v8.5078c0 .9689-.1533 1.3322-1.5941 1.5441-1.6555.2422-3.5256.3633-5.0278.3633-3.9855 0-6.0702-.4238-6.1315-4.5112.0613-3.5424 1.9008-4.2993 5.4263-4.2993h3.3723v-.9386c0-1.2413-.797-1.8166-2.4832-1.8166h-4.7825c-.1226 0-.3679-.0908-.3679-.3936v-2.1497c0-.1816.092-.3633.3679-.4238 1.6248-.3028 3.1577-.3936 4.7825-.3936 5.2731 0 6.438 1.6652 6.438 4.5112Zm-6.9592 4.0874c-1.3795 0-1.8394.4541-1.8394 1.5744 0 1.2413.5825 1.6047 2.238 1.6047.5212 0 1.8701-.0606 2.6365-.2725v-2.9066h-3.0351ZM45.2198 16.8343h5.5183c.1533 0 .3679-.0908.3679-.3936v-2.1193c0-.1211-.092-.3634-.4599-.3936-.9197-.0909-2.146-.1817-3.6175-.2422-.4292-.0303-1.7168-.0303-1.8701-.0303-3.3723 0-5.6409.7872-5.6409 3.8149v.4238c0 1.7864.5212 3.1488 2.7898 3.936 0 0 3.7402 1.3322 3.8628 1.3625.5825.2725.7971.6055.7971.9991v.2725c0 .6661-.5212.9386-1.5635.9386h-5.5183c-.1533 0-.3679.0908-.3679.3936v2.1194c0 .1211.0919.3633.4598.3936.9198.0908 2.146.1816 3.6176.2422.4292.0302 1.7168.0302 1.8701.0302 3.3723 0 5.6409-.7872 5.6409-3.8148v-.4239c0-1.7863-.5212-3.1488-2.7898-3.936 0 0-3.7402-1.3322-3.8628-1.3625-.5825-.2724-.7971-.6055-.7971-.9991v-.2725c0-.6661.5212-.9386 1.5635-.9386ZM58.3097 16.8343h5.5183c.1532 0 .3678-.0908.3678-.3936v-2.1193c0-.1211-.0919-.3634-.4598-.3936-.9197-.0909-2.146-.1817-3.6176-.2422-.4292-.0303-1.7168-.0303-1.8701-.0303-3.3722 0-5.6409.7872-5.6409 3.8149v.4238c0 1.7864.5212 3.1488 2.7898 3.936 0 0 3.7402 1.3322 3.8628 1.3625.5825.2725.7971.6055.7971.9991v.2725c0 .6661-.5211.9386-1.5635.9386h-5.5183c-.1533 0-.3679.0908-.3679.3936v2.1194c0 .1211.092.3633.4599.3936.9197.0908 2.146.1816 3.6175.2422.4292.0302 1.7168.0302 1.8701.0302 3.3723 0 5.6409-.7872 5.6409-3.8148v-.4239c0-1.7863-.5211-3.1488-2.7898-3.936 0 0-3.7401-1.3322-3.8628-1.3625-.5825-.2724-.7971-.6055-.7971-.9991v-.2725c0-.6661.5519-.9386 1.5636-.9386ZM79.2788 13.6553c-1.8394 0-3.0963.2725-3.8321.6661-.7358-.3936-1.8701-.6661-3.8322-.6661-4.7518 0-5.7942 1.423-5.7942 6.2673v8.2655c0 .212.184.3634.3679.3634h3.1577c.1839 0 .3679-.1514.3679-.3634v-7.2664c0-3.1791.3066-3.7846 1.8701-3.9057 1.4715.0908 1.8394.6358 1.8701 3.3002v7.8719c0 .1514.1226.3634.3678.3634h3.1577c.184 0 .3679-.1514.3679-.3634v-7.8719c.0307-2.6341.3986-3.1791 1.8701-3.3002 1.5635.1211 1.8701.7266 1.8701 3.9057v7.2664c0 .2423.2146.3634.3679.3634h3.1883c.184 0 .3679-.1514.3679-.3634v-8.2655c.0613-4.814-.9503-6.2673-5.7329-6.2673Z"/>
  </g>
  <defs>
    <clipPath id="a">
      <path fill="#fff" d="M0 0h160v35H0z"/>
    </clipPath>
  </defs>
</svg>`;
  return sharp(Buffer.from(svg))
    .resize(640, 140)
    .png()
    .toBuffer()
    .then(buf => "image/png;base64," + buf.toString("base64"));
}

// ---------------------------------------------------------------------------
// Shared slide helpers (pres passed for shape constants)
// ---------------------------------------------------------------------------

function addHeaderBar(pres, slide, title) {
  slide.addText(title, {
    x: 0.5, y: 0.08, w: 9, h: 0.6,
    fontSize: SIZE.TITLE - 4, fontFace: FONTS.HEADING,
    color: COLORS.GRAY_MEDIUM, bold: true, margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: 0, y: LAYOUT.HEADER_LINE_Y, w: SLIDE.W, h: 0,
    line: { color: COLORS.BLACK, width: 1 },
  });
}

function addPageNumber(slide, num) {
  slide.addText(String(num), {
    x: SLIDE.W - 0.8, y: 0.12, w: 0.5, h: 0.5,
    fontSize: SIZE.CAPTION, fontFace: FONTS.BODY,
    color: COLORS.GRAY_LIGHTEST, align: "right", margin: 0,
  });
}

function addFooterNote(slide, text) {
  slide.addText(text, {
    x: LAYOUT.MARGIN_X, y: LAYOUT.FOOTER_Y, w: LAYOUT.CONTENT_W, h: 0.3,
    fontSize: SIZE.NOTE, fontFace: FONTS.BODY,
    color: COLORS.GRAY_LIGHT, align: "left", margin: 0,
  });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
async function main() {
  const {
    FaRocket, FaChartLine, FaUsers, FaCog,
    FaLightbulb, FaShieldAlt, FaBolt, FaGlobe,
    FaCheckCircle, FaArrowRight,
    FaQuoteLeft,
  } = require("react-icons/fa");

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Classmethod";
  pres.title = "Classmethod Slide Template";

  const SHAPES = pres.shapes;

  // Pre-render icons
  const icons = {};
  const iconList = [
    ["rocket", FaRocket, COLORS.BLUE],
    ["chart", FaChartLine, COLORS.BLUE],
    ["users", FaUsers, COLORS.BLUE],
    ["cog", FaCog, COLORS.BLUE],
    ["lightbulb", FaLightbulb, COLORS.BLUE],
    ["shield", FaShieldAlt, COLORS.BLUE],
    ["bolt", FaBolt, COLORS.BLUE],
    ["globe", FaGlobe, COLORS.BLUE],
    ["check", FaCheckCircle, COLORS.BLUE],
    ["arrow", FaArrowRight, COLORS.GRAY_LIGHTEST],
    ["quote", FaQuoteLeft, COLORS.GRAY_LIGHTER],
    ["rocket_white", FaRocket, COLORS.WHITE],
    ["chart_white", FaChartLine, COLORS.WHITE],
    ["users_white", FaUsers, COLORS.WHITE],
    ["check_white", FaCheckCircle, COLORS.WHITE],
    ["lightbulb_white", FaLightbulb, COLORS.WHITE],
    ["shield_white", FaShieldAlt, COLORS.WHITE],
    ["globe_white", FaGlobe, COLORS.WHITE],
    ["cog_white", FaCog, COLORS.WHITE],
  ];
  for (const [name, Comp, color] of iconList) {
    icons[name] = await iconToBase64Png(Comp, color);
  }
  const logo = await classmethodLogoBase64();

  let pageNum = 0;

  // =======================================================================
  // SLIDE 1: Title Slide
  // =======================================================================
  {
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    s.addImage({ data: logo, x: 3.5, y: 0.8, w: 3.0, h: 0.66 });

    s.addText("XXXX プレゼンテーションタイトル", {
      x: 0.5, y: 1.7, w: 9, h: 1.4,
      fontSize: SIZE.TITLE + 2, fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK, bold: true, align: "center", valign: "middle",
    });

    s.addText("XXXX サブタイトル・説明文をここに入力", {
      x: 1, y: 3.1, w: 8, h: 0.6,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT, align: "center",
    });

    s.addShape(SHAPES.LINE, {
      x: 3.5, y: 3.9, w: 3, h: 0,
      line: { color: COLORS.BORDER, width: 1 },
    });

    s.addText("XXXX 20XX/XX/XX  発表者名", {
      x: 2, y: 4.1, w: 6, h: 0.5,
      fontSize: SIZE.SMALL, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST, align: "center",
    });
  }

  // =======================================================================
  // SLIDE 2: Section Divider
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    s.addShape(SHAPES.RECTANGLE, {
      x: 0.6, y: 2.0, w: 0.07, h: 1.5,
      fill: { color: COLORS.BLUE },
    });

    s.addText("XXXX セクションタイトル", {
      x: 0.9, y: 2.0, w: 8.5, h: 0.8,
      fontSize: SIZE.H2 + 4, fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK, bold: true, valign: "middle", margin: 0,
    });

    s.addText("XXXX セクションの説明テキスト", {
      x: 0.9, y: 2.9, w: 8.5, h: 0.6,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT, margin: 0,
    });
  }

  // =======================================================================
  // SLIDE 3: Standard Content (Header + Bullets)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX スライドタイトル");
    addPageNumber(s, pageNum);

    s.addText([
      { text: "XXXX 見出しテキスト", options: { bold: true, fontSize: SIZE.H3, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX 箇条書き項目1：説明テキストをここに入力", options: { bullet: true, breakLine: true } },
      { text: "XXXX 箇条書き項目2：説明テキストをここに入力", options: { bullet: true, breakLine: true } },
      { text: "XXXX 箇条書き項目3：説明テキストをここに入力", options: { bullet: true, breakLine: true } },
      { text: "XXXX 箇条書き項目4：説明テキストをここに入力", options: { bullet: true } },
    ], {
      x: LAYOUT.MARGIN_X, y: LAYOUT.CONTENT_TOP, w: LAYOUT.CONTENT_W, h: LAYOUT.CONTENT_H,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
      valign: "top", paraSpaceAfter: 4,
    });
  }

  // =======================================================================
  // SLIDE 4: No Header (Full Content)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();

    s.addText("XXXX フルコンテンツのタイトル", {
      x: LAYOUT.MARGIN_X, y: 0.3, w: LAYOUT.CONTENT_W, h: 0.8,
      fontSize: SIZE.TITLE, fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK, bold: true, margin: 0,
    });

    s.addText([
      { text: "XXXX このレイアウトはヘッダーバーがなく、コンテンツ領域を広く使えます。情報量が多いスライドや、独自のレイアウトが必要な場合に適しています。", options: { breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX 箇条書き項目1", options: { bullet: true, breakLine: true } },
      { text: "XXXX 箇条書き項目2", options: { bullet: true, breakLine: true } },
      { text: "XXXX 箇条書き項目3", options: { bullet: true } },
    ], {
      x: LAYOUT.MARGIN_X, y: 1.2, w: LAYOUT.CONTENT_W, h: 4.0,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
      valign: "top", paraSpaceAfter: 4,
    });

    addPageNumber(s, pageNum);
  }

  // =======================================================================
  // SLIDE 5: Image Center
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 図表のタイトル");
    addPageNumber(s, pageNum);

    const imgW = 7, imgH = 3.6;
    const imgX = (SLIDE.W - imgW) / 2;
    const imgY = LAYOUT.CONTENT_TOP + 0.15;
    s.addShape(SHAPES.RECTANGLE, {
      x: imgX, y: imgY, w: imgW, h: imgH,
      fill: { color: COLORS.BG_ACCENT },
      line: { color: COLORS.BORDER, width: 1, dashType: "dash" },
    });
    s.addText("XXXX 画像・図表プレースホルダー", {
      x: imgX, y: imgY, w: imgW, h: imgH,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST, align: "center", valign: "middle",
    });

    addFooterNote(s, "XXXX 出典・注釈テキスト");
  }

  // =======================================================================
  // SLIDE 6: Content + Image Right (50/50)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX テキスト＋図（右）");
    addPageNumber(s, pageNum);

    const splitX = 5.1;
    s.addText([
      { text: "XXXX 見出し", options: { bold: true, fontSize: SIZE.H3, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX 左側にテキスト説明を配置します。箇条書きや段落テキストで内容を表現します。", options: { breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX ポイント1の説明", options: { bullet: true, breakLine: true } },
      { text: "XXXX ポイント2の説明", options: { bullet: true, breakLine: true } },
      { text: "XXXX ポイント3の説明", options: { bullet: true } },
    ], {
      x: LAYOUT.MARGIN_X, y: LAYOUT.CONTENT_TOP, w: splitX - LAYOUT.MARGIN_X - 0.3, h: LAYOUT.CONTENT_H,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
      valign: "top", paraSpaceAfter: 4,
    });

    s.addShape(SHAPES.RECTANGLE, {
      x: splitX, y: LAYOUT.HEADER_LINE_Y,
      w: SLIDE.W - splitX, h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fill: { color: COLORS.BG_ACCENT },
    });
    s.addText("XXXX 画像エリア", {
      x: splitX, y: LAYOUT.HEADER_LINE_Y,
      w: SLIDE.W - splitX, h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST, align: "center", valign: "middle",
    });
  }

  // =======================================================================
  // SLIDE 7: Content + Image Left (50/50)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX テキスト＋図（左）");
    addPageNumber(s, pageNum);

    const splitX = 4.9;
    s.addShape(SHAPES.RECTANGLE, {
      x: 0, y: LAYOUT.HEADER_LINE_Y,
      w: splitX, h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fill: { color: COLORS.BG_ACCENT },
    });
    s.addText("XXXX 画像エリア", {
      x: 0, y: LAYOUT.HEADER_LINE_Y,
      w: splitX, h: SLIDE.H - LAYOUT.HEADER_LINE_Y,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHTEST, align: "center", valign: "middle",
    });

    s.addText([
      { text: "XXXX 見出し", options: { bold: true, fontSize: SIZE.H3, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX 右側にテキストを配置するレイアウトです。", options: { breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX ポイント1の説明", options: { bullet: true, breakLine: true } },
      { text: "XXXX ポイント2の説明", options: { bullet: true, breakLine: true } },
      { text: "XXXX ポイント3の説明", options: { bullet: true } },
    ], {
      x: splitX + 0.3, y: LAYOUT.CONTENT_TOP, w: SLIDE.W - splitX - 0.3 - LAYOUT.MARGIN_X, h: LAYOUT.CONTENT_H,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
      valign: "top", paraSpaceAfter: 4,
    });
  }

  // =======================================================================
  // SLIDE 8: Two Column Layout
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 2カラムレイアウト");
    addPageNumber(s, pageNum);

    const colW = (LAYOUT.CONTENT_W - LAYOUT.GAP) / 2;
    const colY = LAYOUT.CONTENT_TOP;

    s.addText([
      { text: "XXXX 左カラム見出し", options: { bold: true, fontSize: SIZE.H3, breakLine: true, color: COLORS.GRAY_DARK } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX 左カラムの内容1", options: { bullet: true, breakLine: true } },
      { text: "XXXX 左カラムの内容2", options: { bullet: true, breakLine: true } },
      { text: "XXXX 左カラムの内容3", options: { bullet: true } },
    ], {
      x: LAYOUT.MARGIN_X, y: colY, w: colW, h: LAYOUT.CONTENT_H,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
      valign: "top", paraSpaceAfter: 4,
    });

    s.addText([
      { text: "XXXX 右カラム見出し", options: { bold: true, fontSize: SIZE.H3, breakLine: true, color: COLORS.GRAY_DARK } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "XXXX 右カラムの内容1", options: { bullet: true, breakLine: true } },
      { text: "XXXX 右カラムの内容2", options: { bullet: true, breakLine: true } },
      { text: "XXXX 右カラムの内容3", options: { bullet: true } },
    ], {
      x: LAYOUT.MARGIN_X + colW + LAYOUT.GAP, y: colY, w: colW, h: LAYOUT.CONTENT_H,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
      valign: "top", paraSpaceAfter: 4,
    });

    s.addShape(SHAPES.LINE, {
      x: LAYOUT.MARGIN_X + colW + LAYOUT.GAP / 2, y: colY,
      w: 0, h: LAYOUT.CONTENT_H,
      line: { color: COLORS.BORDER, width: 1 },
    });
  }

  // =======================================================================
  // SLIDE 9: Three Column Layout
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 3カラムレイアウト");
    addPageNumber(s, pageNum);

    const gap = 0.25;
    const colW = (LAYOUT.CONTENT_W - gap * 2) / 3;
    const colY = LAYOUT.CONTENT_TOP;
    const labels = ["XXXX カラム1", "XXXX カラム2", "XXXX カラム3"];

    for (let i = 0; i < 3; i++) {
      const x = LAYOUT.MARGIN_X + i * (colW + gap);
      s.addText([
        { text: labels[i], options: { bold: true, fontSize: SIZE.H3, breakLine: true, color: COLORS.GRAY_DARK } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "XXXX 項目A", options: { bullet: true, breakLine: true } },
        { text: "XXXX 項目B", options: { bullet: true, breakLine: true } },
        { text: "XXXX 項目C", options: { bullet: true } },
      ], {
        x, y: colY, w: colW, h: LAYOUT.CONTENT_H,
        fontSize: SIZE.BODY, fontFace: FONTS.BODY, color: COLORS.GRAY_MEDIUM,
        valign: "top", paraSpaceAfter: 4,
      });

      if (i < 2) {
        s.addShape(SHAPES.LINE, {
          x: x + colW + gap / 2, y: colY, w: 0, h: LAYOUT.CONTENT_H,
          line: { color: COLORS.BORDER, width: 1 },
        });
      }
    }
  }

  // =======================================================================
  // SLIDE 10: Key Stats / Numbers
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 主要指標");
    addPageNumber(s, pageNum);

    const stats = [
      { num: "98%", label: "XXXX 指標ラベル1", icon: "rocket" },
      { num: "2.5x", label: "XXXX 指標ラベル2", icon: "chart" },
      { num: "500+", label: "XXXX 指標ラベル3", icon: "users" },
    ];

    const cardW = 2.6, gap = 0.35;
    const totalW = cardW * 3 + gap * 2;
    const startX = (SLIDE.W - totalW) / 2;
    const cardY = LAYOUT.CONTENT_TOP + 0.15;
    const cardH = 3.8;

    for (let i = 0; i < stats.length; i++) {
      const x = startX + i * (cardW + gap);

      // Card background
      s.addShape(SHAPES.RECTANGLE, {
        x, y: cardY, w: cardW, h: cardH,
        fill: { color: COLORS.WHITE },
        line: { color: COLORS.BORDER, width: 0.5 },
        shadow: makeCardShadow(),
      });
      // Blue top accent
      s.addShape(SHAPES.RECTANGLE, {
        x, y: cardY, w: cardW, h: 0.06,
        fill: { color: COLORS.BLUE },
      });
      // Icon in circle
      const circleSize = 0.6;
      s.addShape(SHAPES.OVAL, {
        x: x + (cardW - circleSize) / 2, y: cardY + 0.3, w: circleSize, h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      const whiteKey = stats[i].icon + "_white";
      if (icons[whiteKey]) {
        s.addImage({
          data: icons[whiteKey],
          x: x + (cardW - 0.34) / 2, y: cardY + 0.3 + (circleSize - 0.34) / 2, w: 0.34, h: 0.34,
        });
      }
      // Big number
      s.addText(stats[i].num, {
        x, y: cardY + 1.15, w: cardW, h: 1.1,
        fontSize: SIZE.STAT_NUMBER - 8, fontFace: FONTS.HEADING,
        color: COLORS.BLUE, bold: true, align: "center", valign: "middle", margin: 0,
      });
      // Label
      s.addText(stats[i].label, {
        x: x + 0.2, y: cardY + 2.5, w: cardW - 0.4, h: 1.0,
        fontSize: SIZE.BODY, fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT, align: "center", valign: "top", margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 11: Timeline / Process
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX プロセス・手順");
    addPageNumber(s, pageNum);

    const steps = [
      { num: "01", title: "XXXX ステップ1", desc: "XXXX 説明文をここに入力" },
      { num: "02", title: "XXXX ステップ2", desc: "XXXX 説明文をここに入力" },
      { num: "03", title: "XXXX ステップ3", desc: "XXXX 説明文をここに入力" },
      { num: "04", title: "XXXX ステップ4", desc: "XXXX 説明文をここに入力" },
    ];

    const stepW = 2.0, gapX = 0.35;
    const totalW = stepW * 4 + gapX * 3;
    const startX = (SLIDE.W - totalW) / 2;
    const baseY = LAYOUT.CONTENT_TOP + 0.6;

    // Connecting line behind circles
    const lineY = baseY + 0.35;
    s.addShape(SHAPES.LINE, {
      x: startX + stepW / 2, y: lineY,
      w: totalW - stepW, h: 0,
      line: { color: COLORS.BORDER, width: 2 },
    });

    for (let i = 0; i < steps.length; i++) {
      const x = startX + i * (stepW + gapX);
      const circleCx = x + stepW / 2;
      const circleSize = 0.7;

      s.addShape(SHAPES.OVAL, {
        x: circleCx - circleSize / 2, y: baseY, w: circleSize, h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      s.addText(steps[i].num, {
        x: circleCx - circleSize / 2, y: baseY, w: circleSize, h: circleSize,
        fontSize: SIZE.BODY, fontFace: FONTS.HEADING,
        color: COLORS.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });

      s.addText(steps[i].title, {
        x, y: baseY + 0.9, w: stepW, h: 0.5,
        fontSize: SIZE.BODY, fontFace: FONTS.HEADING,
        color: COLORS.GRAY_DARK, bold: true, align: "center", valign: "top", margin: 0,
      });
      s.addText(steps[i].desc, {
        x, y: baseY + 1.4, w: stepW, h: 1.2,
        fontSize: SIZE.SMALL, fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT, align: "center", valign: "top", margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 12: Comparison / Before-After
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 比較");
    addPageNumber(s, pageNum);

    const colW = 4.1, gap = 0.4;
    const startX = (SLIDE.W - colW * 2 - gap) / 2;
    const cardY = LAYOUT.CONTENT_TOP + 0.1;
    const cardH = 3.8;

    const sides = [
      { label: "XXXX Before / 現状", color: COLORS.GRAY_LIGHT, items: ["XXXX 課題・問題点1", "XXXX 課題・問題点2", "XXXX 課題・問題点3"] },
      { label: "XXXX After / 改善後", color: COLORS.BLUE, items: ["XXXX 改善効果1", "XXXX 改善効果2", "XXXX 改善効果3"] },
    ];

    for (let i = 0; i < 2; i++) {
      const x = startX + i * (colW + gap);
      const side = sides[i];

      s.addShape(SHAPES.RECTANGLE, {
        x, y: cardY, w: colW, h: cardH,
        fill: { color: COLORS.WHITE },
        line: { color: COLORS.BORDER, width: 1 },
        shadow: makeCardShadow(),
      });
      s.addShape(SHAPES.RECTANGLE, {
        x, y: cardY, w: colW, h: 0.07,
        fill: { color: side.color },
      });
      s.addText(side.label, {
        x: x + 0.3, y: cardY + 0.3, w: colW - 0.6, h: 0.5,
        fontSize: SIZE.H3, fontFace: FONTS.HEADING,
        color: side.color, bold: true, margin: 0,
      });
      const items = side.items.map((item, idx) => ({
        text: item,
        options: { bullet: true, breakLine: idx < side.items.length - 1 },
      }));
      s.addText(items, {
        x: x + 0.3, y: cardY + 1.0, w: colW - 0.6, h: cardH - 1.3,
        fontSize: SIZE.BODY, fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM, valign: "top", paraSpaceAfter: 6,
      });
    }
  }

  // =======================================================================
  // SLIDE 13: Icon Grid 2x3
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX 特徴・機能一覧");
    addPageNumber(s, pageNum);

    const gridIcons = [
      { key: "rocket", title: "XXXX 機能1", desc: "XXXX 説明テキスト" },
      { key: "chart", title: "XXXX 機能2", desc: "XXXX 説明テキスト" },
      { key: "shield", title: "XXXX 機能3", desc: "XXXX 説明テキスト" },
      { key: "lightbulb", title: "XXXX 機能4", desc: "XXXX 説明テキスト" },
      { key: "globe", title: "XXXX 機能5", desc: "XXXX 説明テキスト" },
      { key: "cog", title: "XXXX 機能6", desc: "XXXX 説明テキスト" },
    ];

    const cols = 3, rows = 2;
    const cellW = 2.7, cellH = 1.8;
    const gapX = 0.25, gapY = 0.2;
    const totalW = cellW * cols + gapX * (cols - 1);
    const totalH = cellH * rows + gapY * (rows - 1);
    const startX = (SLIDE.W - totalW) / 2;
    const startY = LAYOUT.CONTENT_TOP + (LAYOUT.CONTENT_H - totalH) / 2;

    for (let idx = 0; idx < gridIcons.length; idx++) {
      const row = Math.floor(idx / cols);
      const col = idx % cols;
      const x = startX + col * (cellW + gapX);
      const y = startY + row * (cellH + gapY);

      const circleSize = 0.55;
      s.addShape(SHAPES.OVAL, {
        x: x + 0.15, y: y + 0.1, w: circleSize, h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      const whiteIconKey = gridIcons[idx].key + "_white";
      if (icons[whiteIconKey]) {
        s.addImage({
          data: icons[whiteIconKey],
          x: x + 0.15 + (circleSize - 0.3) / 2, y: y + 0.1 + (circleSize - 0.3) / 2,
          w: 0.3, h: 0.3,
        });
      }

      s.addText(gridIcons[idx].title, {
        x: x + 0.15, y: y + 0.75, w: cellW - 0.3, h: 0.35,
        fontSize: SIZE.BODY, fontFace: FONTS.HEADING,
        color: COLORS.GRAY_DARK, bold: true, margin: 0,
      });
      s.addText(gridIcons[idx].desc, {
        x: x + 0.15, y: y + 1.1, w: cellW - 0.3, h: 0.6,
        fontSize: SIZE.SMALL, fontFace: FONTS.BODY,
        color: COLORS.GRAY_LIGHT, margin: 0,
      });
    }
  }

  // =======================================================================
  // SLIDE 14: Quote / Callout
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    s.addImage({
      data: icons.quote, x: 1.2, y: 1.0, w: 0.6, h: 0.6,
    });

    s.addText("XXXX ここに引用文やキーメッセージを入力します。印象的なフレーズや重要な発言を大きく表示するレイアウトです。", {
      x: 1.2, y: 1.7, w: 7.5, h: 2.2,
      fontSize: SIZE.H2, fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK, italic: true, valign: "top", margin: 0,
    });

    s.addShape(SHAPES.LINE, {
      x: 1.2, y: 4.1, w: 1.5, h: 0,
      line: { color: COLORS.BLUE, width: 2 },
    });
    s.addText("XXXX 発言者名  /  XXXX 役職・所属", {
      x: 1.2, y: 4.3, w: 7, h: 0.4,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT, margin: 0,
    });
  }

  // =======================================================================
  // SLIDE 15: Agenda / Table of Contents
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();

    s.addText("XXXX アジェンダ", {
      x: LAYOUT.MARGIN_X, y: 0.4, w: LAYOUT.CONTENT_W, h: 0.8,
      fontSize: SIZE.TITLE, fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK, bold: true, margin: 0,
    });

    const agendaItems = [
      "XXXX アジェンダ項目1",
      "XXXX アジェンダ項目2",
      "XXXX アジェンダ項目3",
      "XXXX アジェンダ項目4",
      "XXXX アジェンダ項目5",
    ];

    const itemH = 0.7, startY = 1.4;

    for (let i = 0; i < agendaItems.length; i++) {
      const y = startY + i * (itemH + 0.12);
      const circleSize = 0.45;

      s.addShape(SHAPES.OVAL, {
        x: LAYOUT.MARGIN_X, y: y + 0.12, w: circleSize, h: circleSize,
        fill: { color: COLORS.BLUE },
      });
      s.addText(String(i + 1), {
        x: LAYOUT.MARGIN_X, y: y + 0.12, w: circleSize, h: circleSize,
        fontSize: SIZE.BODY - 2, fontFace: FONTS.HEADING,
        color: COLORS.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });

      s.addText(agendaItems[i], {
        x: LAYOUT.MARGIN_X + 0.65, y: y, w: LAYOUT.CONTENT_W - 0.65, h: itemH,
        fontSize: SIZE.H3 - 2, fontFace: FONTS.BODY,
        color: COLORS.GRAY_MEDIUM, valign: "middle", margin: 0,
      });

      if (i < agendaItems.length - 1) {
        s.addShape(SHAPES.LINE, {
          x: LAYOUT.MARGIN_X + 0.65, y: y + itemH + 0.03, w: LAYOUT.CONTENT_W - 0.65, h: 0,
          line: { color: COLORS.BORDER, width: 0.5 },
        });
      }
    }
  }

  // =======================================================================
  // SLIDE 16: Table (data presentation)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    addHeaderBar(pres, s, "XXXX テーブルレイアウト");
    addPageNumber(s, pageNum);

    const headerRow = [
      { text: "XXXX 項目", options: { bold: true, color: COLORS.GRAY_MEDIUM, fill: { color: COLORS.BG_ACCENT } } },
      { text: "XXXX 列2", options: { bold: true, color: COLORS.GRAY_MEDIUM, fill: { color: COLORS.BG_ACCENT } } },
      { text: "XXXX 列3", options: { bold: true, color: COLORS.GRAY_MEDIUM, fill: { color: COLORS.BG_ACCENT } } },
      { text: "XXXX 列4", options: { bold: true, color: COLORS.GRAY_MEDIUM, fill: { color: COLORS.BG_ACCENT } } },
    ];

    const dataRows = [];
    for (let r = 0; r < 5; r++) {
      dataRows.push([
        { text: `XXXX \u30C7\u30FC\u30BF${r + 1}-1`, options: { color: COLORS.GRAY_MEDIUM } },
        { text: `XXXX \u30C7\u30FC\u30BF${r + 1}-2`, options: { color: COLORS.GRAY_MEDIUM } },
        { text: `XXXX \u30C7\u30FC\u30BF${r + 1}-3`, options: { color: COLORS.GRAY_MEDIUM } },
        { text: `XXXX \u30C7\u30FC\u30BF${r + 1}-4`, options: { color: COLORS.GRAY_MEDIUM } },
      ]);
    }

    s.addTable([headerRow, ...dataRows], {
      x: LAYOUT.MARGIN_X, y: LAYOUT.CONTENT_TOP + 0.1,
      w: LAYOUT.CONTENT_W, h: 3.5,
      fontSize: SIZE.SMALL, fontFace: FONTS.BODY,
      border: { pt: 0.5, color: COLORS.BORDER },
      colW: [2.2, 2.2, 2.2, 2.2],
      rowH: [0.5, 0.5, 0.5, 0.5, 0.5, 0.5],
      align: "left",
      valign: "middle",
    });

    addFooterNote(s, "XXXX 出典・注釈テキスト");
  }

  // =======================================================================
  // SLIDE 17: Blue accent section (dark variant)
  // =======================================================================
  {
    pageNum++;
    const s = pres.addSlide();
    s.background = { color: COLORS.BLUE };

    s.addText("XXXX キーメッセージ", {
      x: 1.2, y: 1.4, w: 7.6, h: 1.2,
      fontSize: SIZE.TITLE + 6, fontFace: FONTS.HEADING,
      color: COLORS.WHITE, bold: true, valign: "middle", margin: 0,
    });

    s.addText("XXXX 強調したいメッセージやセクションの導入文をここに入力します。青背景のスライドは視覚的なアクセントになります。", {
      x: 1.2, y: 2.8, w: 7.6, h: 1.5,
      fontSize: SIZE.BODY + 2, fontFace: FONTS.BODY,
      color: COLORS.WHITE, valign: "top", margin: 0,
    });
  }

  // =======================================================================
  // SLIDE 18: Thank You / Closing
  // =======================================================================
  {
    const s = pres.addSlide();
    s.background = { color: COLORS.BG_ACCENT };

    s.addImage({ data: logo, x: 3.5, y: 1.2, w: 3.0, h: 0.66 });

    s.addText("XXXX ご清聴ありがとうございました", {
      x: 0.5, y: 2.1, w: 9, h: 1.0,
      fontSize: SIZE.TITLE, fontFace: FONTS.HEADING,
      color: COLORS.GRAY_DARK, bold: true, align: "center", valign: "middle",
    });

    s.addShape(SHAPES.LINE, {
      x: 3.5, y: 3.3, w: 3, h: 0,
      line: { color: COLORS.BORDER, width: 1 },
    });

    s.addText([
      { text: "XXXX お名前", options: { bold: true, breakLine: true } },
      { text: "XXXX 部署・所属", options: { breakLine: true } },
      { text: "XXXX email@example.com", options: {} },
    ], {
      x: 2, y: 3.5, w: 6, h: 1.4,
      fontSize: SIZE.BODY, fontFace: FONTS.BODY,
      color: COLORS.GRAY_LIGHT, align: "center", paraSpaceAfter: 4,
    });
  }

  // =======================================================================
  // Write file
  // =======================================================================
  const outputPath = path.join(__dirname, "classmethod-template.pptx");
  await pres.writeFile({ fileName: outputPath });
  console.log("Template generated:", outputPath);
}

main().catch(err => { console.error(err); process.exit(1); });
