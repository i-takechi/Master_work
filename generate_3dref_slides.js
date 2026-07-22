const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "大阪大学 修士課程セミナー";
pptx.subject = "3DRef論文紹介";
pptx.title = "3DRef: 3D Dataset and Benchmark for Reflection Detection in RGB and Lidar Data";
pptx.company = "Osaka University";
pptx.lang = "ja-JP";
pptx.theme = {
  headFontFace: "Yu Gothic",
  bodyFontFace: "Yu Gothic",
  lang: "ja-JP",
};

const SW = 10;
const SH = 5.625;
const M = 0.5;
const TOP = 0.95;
const BOTTOM = 5.12;
const CW = 9.0;
const CH = BOTTOM - TOP;

const T = {
  bg: "FFFFFF",
  navy: "17365D",
  blue: "2E75B6",
  blue2: "5B9BD5",
  pale: "EAF2F8",
  pale2: "F4F7FA",
  orange: "ED7D31",
  orangePale: "FCE4D6",
  green: "2A9D8F",
  purple: "8E5EA2",
  text: "20242A",
  sub: "4D5966",
  muted: "6F7B86",
  line: "D5DCE3",
  white: "FFFFFF",
  gray: "8C8C8C",
};

const F = {
  cover: 31,
  coverSub: 16,
  title: 23,
  body: 20,
  bodySmall: 17,
  label: 14,
  cardTitle: 18,
  cardBody: 16.5,
  stat: 34,
  tableHead: 13.5,
  tableCell: 12.5,
  cite: 10.5,
};

const A = path.join(__dirname, "3dref_assets");
const FIG = {
  overview: path.join(A, "fig-000.png"),
  platform: path.join(A, "fig-002.jpg"),
  reflective: path.join(A, "fig-003.jpg"),
  seq1: path.join(A, "fig-004.jpg"),
  seq2: path.join(A, "fig-005.jpg"),
  seq3: path.join(A, "fig-006.jpg"),
  angle: path.join(A, "fig-007.png"),
};

const S = pptx.ShapeType;
const C = pptx.ChartType;
const mkShadow = () => ({ type: "outer", color: "000000", blur: 2, angle: 45, distance: 1, opacity: 0.12 });

function tx(slide, text, opts = {}) {
  slide.addText(text, {
    fontFace: "Yu Gothic",
    color: T.text,
    margin: 0,
    breakLine: false,
    shrinkText: true,
    valign: "mid",
    ...opts,
  });
}

function rect(slide, x, y, w, h, fill, line = fill, radius = false) {
  slide.addShape(radius ? S.roundRect : S.rect, {
    x, y, w, h,
    rectRadius: radius ? 0.06 : undefined,
    fill: { color: fill },
    line: { color: line, pt: line === fill ? 0.2 : 1 },
  });
}

function base(title, section = "") {
  const sl = pptx.addSlide();
  sl.background = { color: T.bg };
  rect(sl, 0, 0, SW, 0.72, T.navy);
  tx(sl, title, { x: M, y: 0.09, w: CW, h: 0.52, fontSize: F.title, bold: true, color: T.white, valign: "mid" });
  if (section) {
    tx(sl, section, { x: 8.55, y: 0.74, w: 0.95, h: 0.18, fontSize: F.cite, color: T.muted, align: "right" });
  }
  return sl;
}

function cite(slide, text) {
  tx(slide, text, { x: M, y: 4.94, w: CW, h: 0.16, fontSize: F.cite, color: T.muted, valign: "mid" });
}

function note(slide, lines) {
  slide.addNotes(lines.join("\n"));
}

function card(slide, x, y, w, h, title, body, accent = T.blue, fill = T.pale2, opts = {}) {
  slide.addShape(S.rect, { x, y, w, h, fill: { color: fill }, line: { color: T.line, pt: 0.7 }, shadow: mkShadow() });
  rect(slide, x, y, 0.06, h, accent);
  tx(slide, title, { x: x + 0.22, y: y + 0.1, w: w - 0.36, h: 0.38, fontSize: opts.titleSize || F.cardTitle, bold: true, color: T.navy, valign: "top" });
  tx(slide, body, { x: x + 0.22, y: y + 0.52, w: w - 0.36, h: h - 0.64, fontSize: opts.bodySize || F.cardBody, color: T.sub, valign: "top", breakLine: true, margin: [2, 0, 4, 0] });
}

function pill(slide, text, x, y, w, color = T.blue, fill = T.pale) {
  rect(slide, x, y, w, 0.36, fill, color, true);
  tx(slide, text, { x, y, w, h: 0.36, fontSize: F.label, color, bold: true, align: "center" });
}

function stat(slide, x, y, w, h, value, label, color = T.blue) {
  rect(slide, x, y, w, h, "FFFFFF", T.line, true);
  tx(slide, value, { x: x + 0.08, y: y + 0.14, w: w - 0.16, h: 0.52, fontSize: F.stat, bold: true, color, align: "center" });
  tx(slide, label, { x: x + 0.08, y: y + 0.72, w: w - 0.16, h: h - 0.82, fontSize: F.label, color: T.sub, align: "center", valign: "top" });
}

function fitImage(slide, imgPath, x, y, w, h, iw, ih, altText = "") {
  const r = iw / ih;
  let dw = w;
  let dh = dw / r;
  if (dh > h) {
    dh = h;
    dw = dh * r;
  }
  slide.addImage({ path: imgPath, x: x + (w - dw) / 2, y: y + (h - dh) / 2, w: dw, h: dh, altText });
  return { x: x + (w - dw) / 2, y: y + (h - dh) / 2, w: dw, h: dh };
}

function addChart(slide, type, data, opts) {
  slide.addChart(type, data, {
    chartColors: [T.blue, T.orange, T.green, T.purple],
    showTitle: false,
    showLegend: true,
    legendPos: "b",
    legendFontFace: "Yu Gothic",
    legendFontSize: 10.5,
    legendColor: T.sub,
    catAxisLabelFontFace: "Yu Gothic",
    catAxisLabelFontSize: 11.5,
    catAxisLabelColor: T.sub,
    valAxisLabelFontFace: "Yu Gothic",
    valAxisLabelFontSize: 10.5,
    valAxisLabelColor: T.sub,
    valGridLine: { color: "DDE3E8", pt: 0.5 },
    catGridLine: { color: T.bg, transparency: 100 },
    chartArea: { fill: { color: T.bg }, line: { color: T.bg, transparency: 100 } },
    plotArea: { fill: { color: T.bg }, line: { color: T.bg, transparency: 100 } },
    showValue: true,
    dataLabelColor: T.text,
    dataLabelFontFace: "Yu Gothic",
    dataLabelFontSize: 10.5,
    showCatName: false,
    showValAxisTitle: false,
    showCatAxisTitle: false,
    ...opts,
  });
}

function addTable(slide, rows, opts = {}) {
  const styled = rows.map((row, ri) => row.map((cell, ci) => {
    const raw = typeof cell === "object" ? cell : { text: String(cell), options: {} };
    const isHead = ri === 0;
    return {
      text: raw.text,
      options: {
        fontFace: "Yu Gothic",
        fontSize: isHead ? F.tableHead : F.tableCell,
        bold: isHead ? true : !!raw.options?.bold,
        color: isHead ? T.white : (raw.options?.color || T.text),
        fill: { color: isHead ? T.navy : (raw.options?.fill || (ri % 2 ? "FFFFFF" : "F3F6F8")) },
        align: raw.options?.align || (ci === 0 ? "left" : "center"),
        valign: "mid",
        margin: 0.05,
        ...raw.options,
      },
    };
  }));
  slide.addTable(styled, {
    x: opts.x, y: opts.y, w: opts.w,
    colW: opts.colW,
    rowH: opts.rowH || 0.36,
    border: { pt: 0.45, color: T.line },
    margin: 0.03,
  });
}

// 1. Title
{
  const sl = pptx.addSlide();
  sl.background = { color: T.navy };
  rect(sl, 0, 4.72, SW, 0.905, "102A49");
  tx(sl, "3DRef", { x: 0.65, y: 0.55, w: 2.2, h: 0.65, fontSize: 34, bold: true, color: T.blue2, valign: "mid" });
  tx(sl, "RGB・LiDAR反射検出のための\n3Dデータセットとベンチマーク", { x: 0.65, y: 1.28, w: 8.7, h: 1.45, fontSize: F.cover, bold: true, color: T.white, valign: "top", breakLine: true });
  tx(sl, "Zhao & Schwertfeger, 3DV 2024", { x: 0.68, y: 2.9, w: 8.5, h: 0.35, fontSize: F.coverSub, color: "C8D9EA" });
  pill(sl, "論文紹介セミナー", 0.68, 3.55, 2.0, T.white, "284B70");
  tx(sl, "大阪大学 修士課程セミナー  |  2026年7月23日", { x: 0.68, y: 4.92, w: 8.6, h: 0.3, fontSize: 14.5, color: "C8D9EA" });
  note(sl, [
    "本日は、3DV 2024で発表された3DRefを紹介します。",
    "焦点は、反射を含む環境をRGBと複数LiDARでどうデータ化したか、そして何がベンチマークで分かったかです。",
    "最後に、データセットとしての強みと限界を批判的に整理します。",
  ]);
}

// 2. Motivation
{
  const sl = base("反射面は『最初の物体で反射する』というLiDARの前提を崩す", "Motivation");
  fitImage(sl, FIG.reflective, 0.55, 1.0, 4.9, 3.72, 2243, 1498, "鏡、ガラス、白板、テレビなどの反射物体");
  card(sl, 5.72, 1.05, 3.73, 1.05, "観測の破綻", "反射点・欠測・透過後の障害物が同じスキャンに混在", T.orange, T.orangePale);
  card(sl, 5.72, 2.30, 3.73, 1.05, "ロボットへの影響", "phantom wall → 地図・自己位置推定・経路計画の失敗", T.blue, T.pale);
  card(sl, 5.72, 3.55, 3.73, 1.05, "必要なデータ", "表面だけでなく、反射点とガラス越し障害物も3Dで区別", T.green, "E7F4F2");
  cite(sl, "Source: Zhao & Schwertfeger (2024), Fig. 3; Tibebu et al. (2021)");
  note(sl, [
    "通常のLiDAR処理は、レーザが最初に当たった実物表面から距離が返ると仮定します。",
    "鏡やガラスでは、反射・透過・吸収が起こり、仮想的な点や穴が生成されます。",
    "これがphantom wallとなり、存在しない壁を避けるなどナビゲーションを壊します。",
    "次に、この問題へ3DRefがどんな問いを立てたかを見ます。",
  ]);
}

// 3. Research question and contributions
{
  const sl = base("本論文は『反射を3D・マルチモーダルに学習できる基盤』を構築する", "Question");
  rect(sl, 0.72, 1.06, 8.56, 1.08, T.pale, T.blue, true);
  tx(sl, "研究問い", { x: 0.95, y: 1.2, w: 1.25, h: 0.3, fontSize: F.label, bold: true, color: T.blue });
  tx(sl, "多様な反射面をRGBと3D LiDARで精密にラベル化し、\n現在の検出手法を公平に比較できるか？", { x: 2.15, y: 1.18, w: 6.75, h: 0.66, fontSize: F.body, bold: true, color: T.navy, valign: "mid" });
  const gap = 0.24;
  const w = (8.56 - gap * 2) / 3;
  card(sl, 0.72, 2.45, w, 2.15, "① 51,823サンプル", "48,024点群＋3,799 RGB\n3種LiDAR・3シーケンス", T.blue, T.pale2);
  card(sl, 0.72 + w + gap, 2.45, w, 2.15, "② 3D Ground Truth", "色付きメッシュを基準に\nray castingで自動ラベル", T.green, T.pale2, { titleSize: 15.5 });
  card(sl, 0.72 + (w + gap) * 2, 2.45, w, 2.15, "③ 標準ベンチマーク", "LiDAR 3手法＋RGB 4手法\nreturn channelも評価", T.orange, T.pale2);
  cite(sl, "Source: Zhao & Schwertfeger (2024), Sec. 1");
  note(sl, [
    "本論文の貢献は、新しい検出モデルそのものではなく、学習と比較のための基盤です。",
    "第一に規模、第二に3Dメッシュを使った統一ラベル、第三にLiDARとRGB双方のベンチマークです。",
    "データセット論文としては、収集・GT生成・評価を一つのパイプラインにした点が核心です。",
  ]);
}

// 4. Prior datasets
{
  const sl = base("既存データは2D中心だが、3DRefは反射の種類と3D点を同時に扱う", "Gap");
  addChart(sl, C.bar, [{
    name: "サンプル数",
    labels: ["GDD", "GSD", "Mirror3D", "TROSD", "3DRef"],
    values: [3900, 4102, 5894, 11060, 51823],
  }], {
    x: 0.55, y: 1.02, w: 5.55, h: 3.85,
    barDir: "bar",
    catAxisLabelPos: "low",
    showLegend: false,
    chartColors: [T.blue],
    valAxisMinVal: 0,
    valAxisMaxVal: 55000,
    valAxisMajorUnit: 10000,
    showValue: true,
    dataLabelPosition: "outEnd",
  });
  card(sl, 6.32, 1.07, 3.12, 1.02, "従来", "RGB / RGB-Dの2Dマスク\n対象は主にglassかmirror", T.gray, T.pale2);
  card(sl, 6.32, 2.29, 3.12, 1.42, "3DRef", "RGB＋multi-return LiDAR\nglass / mirror / other / reflection / obstacle", T.blue, T.pale);
  rect(sl, 6.32, 3.92, 3.12, 0.82, T.orangePale, T.orange, true);
  tx(sl, "約4.7倍", { x: 6.47, y: 3.99, w: 1.2, h: 0.38, fontSize: 24, bold: true, color: T.orange, align: "center" });
  tx(sl, "TROSD比のサンプル規模", { x: 7.65, y: 4.02, w: 1.62, h: 0.32, fontSize: 12.5, color: T.sub, align: "center" });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Table 1（3DRefは点群＋RGBの合計）");
  note(sl, [
    "先行データセットは画像マスク中心で、ガラスか鏡のどちらかに限定されることが多いです。",
    "3DRefは点群48,024とRGB 3,799を合わせて51,823サンプルです。",
    "ただし点群と画像は同一形式のサンプルではないため、規模比較は概算として読む必要があります。",
    "本質的な差は、3D位置と反射現象の種類を明示的に持つことです。",
  ]);
}

// 5. Sensor platform
{
  const sl = base("3種類のmulti-return LiDARと360度カメラが相補的な観測を与える", "Acquisition");
  fitImage(sl, FIG.platform, 0.55, 1.0, 4.0, 3.95, 2983, 3617, "3DRefのデータ収集プラットフォーム");
  addTable(sl, [
    ["Sensor", "方式", "記録リターン"],
    ["Ouster OS0-128", "128ch spinning", "strongest / 2nd"],
    ["Livox Avia", "solid-state", "1st / 2nd / 3rd"],
    ["Hesai QT64", "64ch spinning", "first / last"],
    ["Insta360", "fisheye RGB", "3072×3072"],
  ], { x: 4.75, y: 1.18, w: 4.7, colW: [1.65, 1.45, 1.6], rowH: 0.54 });
  pill(sl, "PTPで1 ms以内", 4.92, 4.17, 1.76, T.blue, T.pale);
  pill(sl, "MA-LIOで時空間整合", 6.88, 4.17, 2.32, T.green, "E7F4F2");
  cite(sl, "Source: Zhao & Schwertfeger (2024), Fig. 2 and Table 2");
  note(sl, [
    "プラットフォームには、戻り方の異なる3種類のLiDARが搭載されています。",
    "Livoxは最大3リターン、Hesaiはfirst/last、Ousterはstrongest/second strongestです。",
    "PTP同期に加え、MA-LIOでオンライン外部較正と残留時間差を補正します。",
    "センサ差そのものが、後のベンチマーク結果にも現れます。",
  ]);
}

// 6. Annotation pipeline
{
  const sl = base("色付き3Dメッシュを共通基準にすることで点群と画像を自動ラベル化する", "Annotation");
  fitImage(sl, FIG.overview, 0.55, 1.0, 3.55, 3.95, 891, 1049, "ラベル付きメッシュから点群と画像マスクを生成する概要");
  const xs = [4.35, 5.65, 6.95, 8.25];
  const labels = ["Mesh構築", "表面を着色", "Ray casting", "Point / Mask"];
  labels.forEach((v, i) => {
    rect(sl, xs[i], 1.42, 1.0, 0.72, i === 3 ? "E7F4F2" : T.pale, i === 3 ? T.green : T.blue, true);
    tx(sl, v, { x: xs[i] + 0.05, y: 1.52, w: 0.9, h: 0.45, fontSize: 13, bold: true, color: T.navy, align: "center" });
    if (i < 3) {
      sl.addShape(S.chevron, { x: xs[i] + 1.05, y: 1.61, w: 0.22, h: 0.3, fill: { color: T.blue }, line: { color: T.blue } });
    }
  });
  card(sl, 4.35, 2.55, 4.9, 1.0, "Point cloud", "LiDAR原点から各点へrayを飛ばし、交差したメッシュ面のラベルを付与", T.blue, T.pale2);
  card(sl, 4.35, 3.75, 4.9, 1.0, "RGB mask", "カメラ姿勢・内部パラメータでラベルメッシュをrenderし、画素マスクを生成", T.green, T.pale2);
  cite(sl, "Source: Zhao & Schwertfeger (2024), Fig. 1 and Sec. 4.2");
  note(sl, [
    "アノテーションの中心は、各センサごとに手作業でラベルするのではなく、共通の3DメッシュをGTにすることです。",
    "まずPolar scannerでメッシュを作り、ガラスなどによる穴を手動修正します。",
    "点群はray casting、RGBはラベルメッシュのレンダリングで自動生成します。",
    "これによりモダリティ間でラベルの意味と位置が揃います。",
  ]);
}

// 7. Labels and scale
{
  const sl = base("6クラスの現象ラベルを48,024点群と3,799画像へ展開する", "Dataset");
  const labels = [
    ["Normal", "実在する通常表面", T.gray],
    ["Glass", "ガラス表面", T.green],
    ["Mirror", "鏡面", T.blue],
    ["Other reflective", "白板・TV・タイル等", T.purple],
    ["Reflection", "仮想物体の反射点", T.orange],
    ["Obstacle behind", "透明物体の背後", "C79A00"],
  ];
  labels.forEach((d, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.55 + col * 2.12;
    const y = 1.08 + row * 1.18;
    rect(sl, x, y, 1.88, 0.94, "FFFFFF", T.line, true);
    rect(sl, x, y, 0.09, 0.94, d[2]);
    tx(sl, d[0], { x: x + 0.22, y: y + 0.10, w: 1.52, h: 0.28, fontSize: 15.5, bold: true, color: T.navy });
    tx(sl, d[1], { x: x + 0.22, y: y + 0.44, w: 1.52, h: 0.32, fontSize: 12.5, color: T.sub, valign: "top" });
  });
  stat(sl, 7.10, 1.08, 2.35, 1.35, "48,024", "点群（16,008 / LiDAR）", T.blue);
  stat(sl, 7.10, 2.63, 2.35, 1.35, "3,799", "RGB画像＋マスク", T.green);
  const seqs = [
    ["Seq.1", "廊下・鏡", "3,732 / sensor"],
    ["Seq.2", "室内", "4,702 / sensor"],
    ["Seq.3", "フロア", "7,574 / sensor"],
  ];
  seqs.forEach((d, i) => {
    const x = 0.65 + i * 2.1;
    pill(sl, `${d[0]}  ${d[1]}`, x, 3.72, 1.9, i === 2 ? T.orange : T.blue, i === 2 ? T.orangePale : T.pale);
    tx(sl, d[2], { x, y: 4.17, w: 1.9, h: 0.28, fontSize: 12.5, color: T.sub, align: "center" });
  });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Sec. 4.2–4.3");
  note(sl, [
    "ラベルは表面材質だけでなく、観測現象まで分けています。",
    "特にReflectionは仮想物体の点、Obstacle behindはガラス越しの実物点です。",
    "3シーケンスは鏡の廊下、ガラスの多いオフィス、TVや白板を含むフロアです。",
    "未ラベルを含めると実装上はlabel 0から6までですが、意味のある評価対象をここでは6クラスで示しています。",
  ]);
}

// 8. Multi-return
{
  const sl = base("後段リターンはReflectionとObstacle behindへ選択的に集中する", "Analysis");
  addChart(sl, C.bar, [
    { name: "2nd return", labels: ["Normal", "Glass", "Mirror", "OtherRef", "Reflection", "Obstacle"], values: [0.97, 0.39, 1.26, 0.18, 22.07, 33.38] },
    { name: "3rd return", labels: ["Normal", "Glass", "Mirror", "OtherRef", "Reflection", "Obstacle"], values: [0.01, 0.00, 0.00, 0.00, 0.77, 4.43] },
  ], {
    x: 0.55, y: 1.03, w: 6.18, h: 3.84,
    barDir: "col",
    grouping: "clustered",
    valAxisMinVal: 0,
    valAxisMaxVal: 40,
    valAxisMajorUnit: 10,
    showValue: true,
    dataLabelPosition: "outEnd",
  });
  rect(sl, 6.95, 1.18, 2.45, 1.16, T.orangePale, T.orange, true);
  tx(sl, "33.38%", { x: 7.13, y: 1.33, w: 2.08, h: 0.46, fontSize: 30, bold: true, color: T.orange, align: "center" });
  tx(sl, "Obstacleの2nd return", { x: 7.13, y: 1.82, w: 2.08, h: 0.25, fontSize: 13, color: T.sub, align: "center" });
  card(sl, 6.95, 2.58, 2.45, 1.03, "解釈", "表面クラスはほぼ1st\n反射・透過後点は後段へ", T.blue, T.pale2, { bodySize: 14.2 });
  card(sl, 6.95, 3.82, 2.45, 0.88, "示唆", "return番号は現象識別の手掛かり", T.green, "E7F4F2", { bodySize: 14.0 });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Table 4");
  note(sl, [
    "この表が、multi-returnを使う動機を最も直接に示します。",
    "通常表面・ガラス・鏡は99%前後がfirst returnです。",
    "一方、Reflectionの22%、Obstacle behindの33%がsecond returnです。",
    "return番号だけで完全分類はできませんが、現象ラベルとの強い相関があります。",
  ]);
}

// 9. Incident angle
{
  const sl = base("入射角が浅くなるほど直接反射が減り、誤反射との比率が変わる", "Analysis");
  fitImage(sl, FIG.angle, 0.6, 1.02, 5.75, 3.80, 800, 637, "入射角と各ラベル比率の関係");
  card(sl, 6.58, 1.18, 2.82, 1.02, "0–20°", "glass / mirrorの直接returnが多い", T.blue, T.pale);
  card(sl, 6.58, 2.42, 2.82, 1.02, "30–60°", "reflection / obstacleの比率が相対的に増加", T.orange, T.orangePale);
  card(sl, 6.58, 3.66, 2.82, 1.02, "設計上の示唆", "複数視点・角度を学習データに含める必要", T.green, "E7F4F2");
  cite(sl, "Source: Zhao & Schwertfeger (2024), Fig. 7（各角度でクラス比率を正規化）");
  note(sl, [
    "横軸はレーザの入射角、縦軸は各角度帯におけるクラス割合です。",
    "ガラスや鏡の直接returnは、面に近い角度で多くなります。",
    "中程度の角度ではReflectionやObstacle behindの割合が相対的に上がります。",
    "単一視点では材質と角度が混同されるため、複数視点収集が重要です。",
  ]);
}

// 10. Benchmark setup
{
  const sl = base("ベンチマークはLiDAR幾何とRGB外観を分け、return追加と再学習を検証する", "Benchmark");
  card(sl, 0.65, 1.05, 4.15, 2.05, "LiDAR segmentation", "MinkowskiNet\nSPVCNN\nCylinder3D", T.blue, T.pale2);
  pill(sl, "XYZI（4ch）", 0.95, 3.38, 1.6, T.gray, "F0F2F4");
  tx(sl, "vs", { x: 2.66, y: 3.38, w: 0.5, h: 0.36, fontSize: F.bodySmall, bold: true, color: T.muted, align: "center" });
  pill(sl, "XYZIR（5ch）", 3.18, 3.38, 1.65, T.blue, T.pale);
  card(sl, 5.20, 1.05, 4.15, 2.05, "RGB segmentation", "GlassSemNet / EBLNet\nHetNet / SATNet", T.orange, T.pale2);
  pill(sl, "外部pretrained", 5.47, 3.38, 1.72, T.gray, "F0F2F4");
  tx(sl, "vs", { x: 7.29, y: 3.38, w: 0.5, h: 0.36, fontSize: F.bodySmall, bold: true, color: T.muted, align: "center" });
  pill(sl, "3DRefで再学習", 7.82, 3.38, 1.55, T.orange, T.orangePale);
  rect(sl, 0.65, 4.05, 8.70, 0.68, T.pale, T.blue, true);
  tx(sl, "共通指標：class-wise IoU と mean IoU（mIoU）", { x: 0.9, y: 4.18, w: 8.2, h: 0.3, fontSize: F.bodySmall, bold: true, color: T.navy, align: "center" });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Sec. 5.1");
  note(sl, [
    "LiDAR側は同じPCSeg基盤で3モデルを比較します。",
    "入力をXYZIからXYZIRへ変え、return channelの寄与をアブレーションします。",
    "RGB側は既存データで学習済みのモデルと、3DRefで再学習したモデルを比較します。",
    "評価指標はクラス別IoUと平均mIoUです。",
  ]);
}

// 11. Lidar results
{
  const sl = base("Cylinder3Dが最高mIoUだが、return追加の平均改善は小さい", "Results");
  addChart(sl, C.bar, [
    { name: "XYZI", labels: ["Minkowski", "SPVCNN", "Cylinder3D"], values: [81.55, 81.85, 83.72] },
    { name: "XYZIR", labels: ["Minkowski", "SPVCNN", "Cylinder3D"], values: [82.11, 82.14, 83.92] },
  ], {
    x: 0.55, y: 1.03, w: 6.05, h: 3.85,
    barDir: "col",
    grouping: "clustered",
    valAxisMinVal: 75,
    valAxisMaxVal: 86,
    valAxisMajorUnit: 2,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0.00",
    chartColors: [T.gray, T.blue],
  });
  stat(sl, 6.82, 1.18, 2.45, 1.2, "83.92", "Cylinder3D + return", T.blue);
  rect(sl, 6.82, 2.60, 2.45, 0.9, T.pale, T.blue, true);
  tx(sl, "+0.35 pt", { x: 7.00, y: 2.72, w: 2.09, h: 0.34, fontSize: 25, bold: true, color: T.blue, align: "center" });
  tx(sl, "3手法の平均改善", { x: 7.00, y: 3.08, w: 2.09, h: 0.20, fontSize: 11.5, color: T.sub, align: "center" });
  card(sl, 6.82, 3.72, 2.45, 0.98, "なぜ小さい？", "Ouster 2nd: < 0.4%\n定義はセンサ依存", T.orange, T.orangePale, { bodySize: 13.5 });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Table 5; 改善値は表から算出");
  note(sl, [
    "総合mIoUではCylinder3Dが最良で、return付き83.92です。",
    "return追加はMinkowskiで約0.57、SPVCNNで約0.30、Cylinder3Dで約0.20ポイントです。",
    "論文本文は約0.5%改善と要約していますが、表からの単純平均は約0.35ポイントです。",
    "改善が限定的なのは、Ousterでsecond returnが極端に少なく、センサごとにreturnの意味も違うためです。",
  ]);
}

// 12. RGB results
{
  const sl = base("3DRefで再学習するとRGB反射検出は20ポイント超改善する", "Results");
  addChart(sl, C.bar, [
    { name: "既存データで学習", labels: ["EBLNet / Glass", "SATNet / Mirror"], values: [60.49, 49.46] },
    { name: "3DRefで再学習", labels: ["EBLNet / Glass", "SATNet / Mirror"], values: [86.71, 82.47] },
  ], {
    x: 0.55, y: 1.03, w: 6.18, h: 3.85,
    barDir: "col",
    grouping: "clustered",
    valAxisMinVal: 0,
    valAxisMaxVal: 100,
    valAxisMajorUnit: 20,
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelFormatCode: "0.00",
    chartColors: [T.gray, T.orange],
  });
  rect(sl, 6.95, 1.18, 2.42, 1.05, T.orangePale, T.orange, true);
  tx(sl, "+26.22", { x: 7.10, y: 1.31, w: 2.12, h: 0.43, fontSize: 29, bold: true, color: T.orange, align: "center" });
  tx(sl, "EBLNet / Glass", { x: 7.10, y: 1.79, w: 2.12, h: 0.22, fontSize: 12, color: T.sub, align: "center" });
  rect(sl, 6.95, 2.47, 2.42, 1.05, T.pale, T.blue, true);
  tx(sl, "+33.01", { x: 7.10, y: 2.60, w: 2.12, h: 0.43, fontSize: 29, bold: true, color: T.blue, align: "center" });
  tx(sl, "SATNet / Mirror", { x: 7.10, y: 3.08, w: 2.12, h: 0.22, fontSize: 12, color: T.sub, align: "center" });
  card(sl, 6.95, 3.75, 2.42, 0.95, "結論", "既存データとの差が大きい", T.green, "E7F4F2", { bodySize: 13.5 });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Table 6; 差分は表から算出");
  note(sl, [
    "RGBでは効果が非常に大きく、EBLNetのglassは約26ポイント改善します。",
    "SATNetのmirrorは約33ポイント改善します。",
    "これは新モデルの優位性ではなく、学習データのドメイン適合の効果です。",
    "既存のガラス・鏡データだけでは、3DRefの多様な反射環境へ一般化しにくいことを示します。",
  ]);
}

// 13. Critical assessment
{
  const sl = base("3DRefは再現可能な基盤を作ったが、融合評価と一般化には余地が残る", "Discussion");
  card(sl, 0.62, 1.08, 4.22, 1.50, "強み ①：整合したGround Truth", "同一メッシュから点群・画像ラベルを生成し、モダリティ間の意味を統一", T.blue, T.pale);
  card(sl, 0.62, 2.82, 4.22, 1.50, "強み ②：センサ差を含む実データ", "3種LiDAR・複数return・多様な室内反射物を一つの評価系で比較", T.green, "E7F4F2");
  card(sl, 5.15, 1.08, 4.22, 1.50, "限界 ①：環境の偏り", "単一施設の3シーケンス中心。屋外・動的環境・材質多様性は未検証", T.orange, T.orangePale);
  card(sl, 5.15, 2.82, 4.22, 1.50, "限界 ②：真のfusion未評価", "RGBとLiDARを別々にベンチマーク。相補性を使うfusion modelは将来課題", T.purple, "F2EAF5");
  rect(sl, 1.62, 4.40, 6.76, 0.42, T.navy, T.navy, true);
  tx(sl, "次の一手：sensor-aware fusion ＋ cross-domain評価", { x: 1.82, y: 4.47, w: 6.36, h: 0.24, fontSize: 16.5, bold: true, color: T.white, align: "center" });
  cite(sl, "Assessment based on Zhao & Schwertfeger (2024), Sec. 6");
  note(sl, [
    "強みは、ラベルの整合性とセンサ多様性です。データセット研究として実用価値があります。",
    "一方で、収集場所は限定的で、屋外や動的物体への一般化は分かりません。",
    "さらに論文はLiDARとRGBを別々に評価しており、タイトルが示唆するmulti-modal fusion自体は未実施です。",
    "今後はセンサ固有のreturn定義を扱うfusionと、施設をまたぐcross-domain評価が重要です。",
  ]);
}

// 14. References
{
  const sl = base("本発表で参照した主要文献", "References");
  const refs = [
    "Zhao, X. & Schwertfeger, S. (2024). 3DRef: 3D Dataset and Benchmark for Reflection Detection in RGB and Lidar Data. 3DV 2024.",
    "Choy, C., Gwak, J. & Savarese, S. (2019). 4D Spatio-Temporal ConvNets: Minkowski Convolutional Neural Networks. CVPR.",
    "Tang, H. et al. (2020). Searching Efficient 3D Architectures with Sparse Point-Voxel Convolution. ECCV.",
    "Zhu, X. et al. (2021). Cylindrical and Asymmetrical 3D Convolution Networks for LiDAR Segmentation. CVPR.",
    "He, H. et al. (2021). Enhanced Boundary Learning for Glass-like Object Segmentation. ICCV.",
    "Huang, T. et al. (2023). Symmetry-Aware Transformer-Based Mirror Detection. AAAI.",
  ];
  const rowH = 0.66;
  refs.forEach((r, i) => {
    const y = 1.02 + i * rowH;
    tx(sl, r, { x: 0.72, y, w: 8.56, h: 0.52, fontSize: 14.2, color: T.text, valign: "mid" });
  });
  note(sl, [
    "主要な引用文献です。",
    "3DRef本論文に加えて、LiDARの3モデルとRGBの代表的な2モデルを示しています。",
    "詳細な先行研究一覧は原論文のReferencesを参照してください。",
  ]);
}

// 15. Conclusions
{
  const sl = pptx.addSlide();
  sl.background = { color: T.navy };
  tx(sl, "3DRefが示した3つの要点", { x: 0.62, y: 0.42, w: 8.75, h: 0.6, fontSize: 28, bold: true, color: T.white });
  const items = [
    ["1", "データ基盤", "51,823サンプルと3DメッシュGTで、反射現象を統一的に学習可能にした"],
    ["2", "multi-return", "後段returnはReflection / Obstacle behindに集中し、識別の有力な手掛かりになる"],
    ["3", "benchmark", "LiDAR改善は小さい一方、RGB再学習は26–33ポイント改善しdomain gapを示した"],
  ];
  items.forEach((d, i) => {
    const y = 1.30 + i * 1.10;
    rect(sl, 0.72, y, 0.62, 0.62, i === 2 ? T.orange : T.blue, i === 2 ? T.orange : T.blue, true);
    tx(sl, d[0], { x: 0.72, y, w: 0.62, h: 0.62, fontSize: 22, bold: true, color: T.white, align: "center" });
    tx(sl, d[1], { x: 1.58, y: y - 0.01, w: 1.65, h: 0.32, fontSize: 17.5, bold: true, color: "A9C8E5" });
    tx(sl, d[2], { x: 3.30, y: y - 0.02, w: 5.87, h: 0.68, fontSize: 17.5, color: T.white, valign: "mid" });
  });
  rect(sl, 0.72, 4.54, 8.56, 0.52, "102A49", "102A49", true);
  tx(sl, "Dataset / code:  http://3dref.github.io   |   質問・議論をお願いします", { x: 0.92, y: 4.64, w: 8.16, h: 0.26, fontSize: 14.5, color: "C8D9EA", align: "center" });
  note(sl, [
    "まとめです。",
    "第一に、3DRefは大規模で整合した反射検出データ基盤を提供しました。",
    "第二に、後段returnが反射・透過現象を表す有効な情報であると定量化しました。",
    "第三に、RGBでは大きなdomain gapがあり、データセット内再学習の価値が高いことを示しました。",
    "今後はRGB-LiDAR融合と施設外一般化が重要です。",
  ]);
}

// 16. Appendix A
{
  const sl = base("補足：公開データ構造は学習・評価・再現実験を一通り支える", "Appendix A");
  const items = [
    ["Raw", "pose / image / labeled mesh / textured mesh / calibration"],
    ["RGB", "glass / mirror / other / all のmask、train / test split"],
    ["Point clouds", "SemanticKITTI形式：XYZI（4ch）とXYZIR（5ch）の点群・ラベル"],
    ["Scripts", "ray tracing / statistics / evaluation"],
    ["Networks", "EBLNet / PCSeg / SATNet のcode・weight"],
  ];
  items.forEach((d, i) => {
    const y = 1.04 + i * 0.76;
    rect(sl, 0.75, y, 1.60, 0.54, i % 2 ? "E7F4F2" : T.pale, i % 2 ? T.green : T.blue, true);
    tx(sl, d[0], { x: 0.75, y, w: 1.60, h: 0.54, fontSize: 17, bold: true, color: T.navy, align: "center" });
    tx(sl, d[1], { x: 2.62, y, w: 6.55, h: 0.54, fontSize: 16.5, color: T.sub });
  });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Sec. 4.4");
  note(sl, [
    "補足としてデータ構造を示します。",
    "Rawには再処理可能な姿勢・較正・メッシュが含まれます。",
    "SemanticKITTI形式で既存の点群セグメンテーション実装に接続できます。",
    "処理スクリプトと学習済みネットワークも公開され、再現実験を始めやすい構成です。",
  ]);
}

// 17. Appendix B
{
  const sl = base("補足：センサ別ではLivoxが最高で、Ousterはglass・mirrorに弱い", "Appendix B");
  addTable(sl, [
    ["Cylinder3D + return", "Total", "Glass", "Mirror", "Reflection", "Obstacle"],
    ["Ouster", { text: "79.71", options: { color: T.orange } }, "58.27", "59.35", { text: "90.51", options: { bold: true, color: T.blue } }, "84.18"],
    ["Hesai", "84.98", { text: "78.69", options: { bold: true, color: T.blue } }, { text: "88.46", options: { bold: true, color: T.blue } }, "86.54", "78.88"],
    ["Livox", { text: "85.96", options: { bold: true, color: T.green } }, "72.64", "85.73", "89.37", { text: "86.59", options: { bold: true, color: T.green } }],
  ], { x: 0.62, y: 1.20, w: 8.76, colW: [2.10, 1.2, 1.2, 1.2, 1.55, 1.51], rowH: 0.58 });
  card(sl, 0.72, 3.92, 2.62, 0.88, "Ouster", "2nd returnが少ない", T.orange, T.orangePale, { bodySize: 14.0 });
  card(sl, 3.69, 3.92, 2.62, 0.88, "Hesai", "glass / mirrorが最良", T.blue, T.pale, { bodySize: 14.0 });
  card(sl, 6.66, 3.92, 2.62, 0.88, "Livox", "総合 / Obstacleが最良", T.green, "E7F4F2", { bodySize: 14.0 });
  cite(sl, "Source: Zhao & Schwertfeger (2024), Table 5（mIoU）");
  note(sl, [
    "センサ別の詳細です。全体ではLivoxが85.96で最良です。",
    "HesaiはglassとmirrorのIoUが高く、LivoxはObstacle behindで強いです。",
    "Ousterは反射点自体のIoUは高い一方、glass・mirror表面の検出が弱いです。",
    "モデル性能だけでなく、return方式とラベル分布の違いを併せて読む必要があります。",
  ]);
}

// 18. Appendix C
{
  const sl = base("補足：想定質問は『GT誤差』『比較の公平性』『融合の欠如』に集約される", "Appendix C");
  card(sl, 0.62, 1.03, 8.76, 1.05, "Q1. メッシュGTは本当に正確か？", "A. 穴や反射由来の破綻は手動修正するが、mesh pose・ray casting誤差は残り得る。GT品質の定量評価は限定的。", T.blue, T.pale2);
  card(sl, 0.62, 2.32, 8.76, 1.05, "Q2. return channelの比較は公平か？", "A. 同じ4ch/5ch入力で比較する一方、Ouster・Hesai・Livoxでreturnの定義と頻度が異なる。sensor-awareな設計が必要。", T.orange, T.pale2);
  card(sl, 0.62, 3.61, 8.76, 1.05, "Q3. なぜRGBとLiDARを融合しないのか？", "A. 本論文はデータセットと単一モダリティbaselineの確立が目的。融合は明示的な将来課題で、3DRefの次の研究機会。", T.green, T.pale2);
  cite(sl, "Discussion based on Zhao & Schwertfeger (2024), Sec. 4–6");
  note(sl, [
    "質疑で出やすい3点をまとめています。",
    "第一はGT品質で、手動修正を含むため完全自動・完全無誤差ではありません。",
    "第二はセンサ間比較で、returnの意味が統一されていません。",
    "第三は融合で、現状はdataset readinessを示した段階です。",
  ]);
}

async function writeDeck() {
  const maxSlides = Number(process.env.MAX_SLIDES || 0);
  if (maxSlides > 0) pptx._slides = pptx._slides.slice(0, maxSlides);
  const outputName = process.env.OUT_NAME || "セミナー202060723.pptx";
  await pptx.writeFile({ fileName: path.join(__dirname, outputName) });
}

writeDeck().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
