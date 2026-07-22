const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "大阪大学 修士課程セミナー";
pptx.subject = "3DRef論文紹介（教授向け詳細版）";
pptx.title = "3DRef: 3D Dataset and Benchmark for Reflection Detection in RGB and Lidar Data";
pptx.company = "Osaka University";
pptx.lang = "ja-JP";
pptx.theme = { headFontFace: "Yu Gothic", bodyFontFace: "Yu Gothic", lang: "ja-JP" };

const SW = 10, SH = 5.625, M = 0.5, TOP = 0.88, BOTTOM = 5.15, CW = 9.0;
const T = {
  navy: "17365D", navy2: "102A49", blue: "2E75B6", blue2: "5B9BD5",
  pale: "EAF2F8", pale2: "F5F7FA", orange: "ED7D31", orangePale: "FCE4D6",
  green: "2A9D8F", greenPale: "E7F4F2", purple: "8064A2", red: "C94C4C",
  text: "20242A", sub: "4D5966", muted: "6F7B86", line: "D5DCE3",
  white: "FFFFFF", gray: "8C8C8C", lightGray: "EFF2F5"
};
const F = { cover: 29, coverSub: 15, title: 21, body: 17, small: 14.5, label: 12.5,
  cardTitle: 16, cardBody: 14, stat: 31, tableHead: 11.5, tableCell: 10.5, cite: 8.5 };
const A = path.join(__dirname, "3dref_assets");
const FIG = {
  overview: path.join(A, "fig-000.png"), platform: path.join(A, "fig-002.jpg"),
  reflective: path.join(A, "fig-003.jpg"), seq1: path.join(A, "fig-004.jpg"),
  seq2: path.join(A, "fig-005.jpg"), seq3: path.join(A, "fig-006.jpg"),
  angle: path.join(A, "fig-007.png")
};
const S = pptx.ShapeType, C = pptx.ChartType;
const shadow = () => ({ type: "outer", color: "000000", blur: 2, angle: 45, distance: 1, opacity: 0.11 });

function tx(sl, text, o={}) {
  sl.addText(text, { fontFace: "Yu Gothic", color: T.text, margin: 0, shrinkText: true,
    valign: "mid", breakLine: false, ...o });
}
function shape(sl, type, x, y, w, h, fill, line=fill, extra={}) {
  sl.addShape(type, { x,y,w,h, fill:{color:fill}, line:{color:line,pt:line===fill?0.2:0.8}, ...extra });
}
function rect(sl,x,y,w,h,fill,line=fill,extra={}) { shape(sl,S.rect,x,y,w,h,fill,line,extra); }
function line(sl,x,y,w,h,color=T.line,pt=1,dash="solid") {
  sl.addShape(S.line,{x,y,w,h,line:{color,pt,dashType:dash,beginArrowType:"none",endArrowType:"none"}});
}
function arrow(sl,x,y,w,h,color=T.blue,pt=2) {
  sl.addShape(S.line,{x,y,w,h,line:{color,pt,beginArrowType:"none",endArrowType:"triangle"}});
}
function base(title, section="") {
  const sl=pptx.addSlide(); sl.background={color:T.white}; rect(sl,0,0,SW,0.68,T.navy);
  tx(sl,title,{x:M,y:0.07,w:8.15,h:0.52,fontSize:F.title,bold:true,color:T.white,valign:"mid"});
  if(section) tx(sl,section,{x:8.55,y:0.12,w:0.92,h:0.4,fontSize:F.cite+1,color:"C8D9EA",align:"right"});
  return sl;
}
function cite(sl,text){ tx(sl,text,{x:M,y:4.96,w:CW,h:0.14,fontSize:F.cite,color:T.muted,valign:"mid"}); }
function notes(sl, lines){ sl.addNotes(lines); }
function fitImage(sl,p,x,y,w,h,iw,ih,alt=""){
  const r=iw/ih; let dw=w,dh=dw/r; if(dh>h){dh=h;dw=dh*r;}
  sl.addImage({path:p,x:x+(w-dw)/2,y:y+(h-dh)/2,w:dw,h:dh,altText:alt});
  return {x:x+(w-dw)/2,y:y+(h-dh)/2,w:dw,h:dh};
}
function card(sl,x,y,w,h,title,body,accent=T.blue,fill=T.pale2,o={}){
  sl.addShape(S.rect,{x,y,w,h,fill:{color:fill},line:{color:T.line,pt:0.7},shadow:shadow()});
  rect(sl,x,y,0.055,h,accent);
  tx(sl,title,{x:x+0.18,y:y+0.08,w:w-0.30,h:o.titleH||0.34,fontSize:o.titleSize||F.cardTitle,bold:true,color:T.navy,valign:"top"});
  tx(sl,body,{x:x+0.18,y:y+(o.bodyY||0.46),w:w-0.30,h:h-(o.bodyY||0.46)-0.10,fontSize:o.bodySize||F.cardBody,color:T.sub,valign:"top",breakLine:true,margin:[1,0,2,0]});
}
function pill(sl,text,x,y,w,color=T.blue,fill=T.pale){
  shape(sl,S.roundRect,x,y,w,0.32,fill,color,{rectRadius:0.05});
  tx(sl,text,{x,y,w,h:0.32,fontSize:F.label,bold:true,color,align:"center"});
}
function callout(sl,value,label,x,y,w,h,color=T.blue){
  shape(sl,S.roundRect,x,y,w,h,T.white,T.line,{rectRadius:0.05,shadow:shadow()});
  tx(sl,value,{x:x+0.08,y:y+0.12,w:w-0.16,h:0.55,fontSize:F.stat,bold:true,color,align:"center"});
  tx(sl,label,{x:x+0.1,y:y+0.72,w:w-0.2,h:h-0.8,fontSize:F.label,color:T.sub,align:"center",valign:"top"});
}
function takeaway(sl,text,color=T.green){
  rect(sl,0.62,4.48,8.76,0.38,"F8FAFB",color);
  rect(sl,0.62,4.48,0.07,0.38,color);
  tx(sl,text,{x:0.82,y:4.51,w:8.35,h:0.30,fontSize:F.small,bold:true,color:T.navy});
}
function grid(sl, headers, rows, x,y,w,h,widths, opts={}){
  const rh=h/(rows.length+1); let cx=x;
  headers.forEach((hd,i)=>{const cw=w*widths[i]; rect(sl,cx,y,cw,rh,T.navy,T.white);
    tx(sl,hd,{x:cx+0.03,y:y+0.02,w:cw-0.06,h:rh-0.04,fontSize:opts.headSize||F.tableHead,bold:true,color:T.white,align:"center"}); cx+=cw;});
  rows.forEach((row,ri)=>{cx=x; row.forEach((cell,ci)=>{const cw=w*widths[ci]; const obj=typeof cell==="object"?cell:{text:String(cell)};
    const fill=obj.fill||(ri%2===0?T.white:T.pale2); rect(sl,cx,y+rh*(ri+1),cw,rh,fill,T.line);
    tx(sl,obj.text,{x:cx+0.03,y:y+rh*(ri+1)+0.02,w:cw-0.06,h:rh-0.04,fontSize:obj.size||opts.cellSize||F.tableCell,
      bold:!!obj.bold,color:obj.color||T.text,align:obj.align||(ci===0?"left":"center"),valign:"mid"}); cx+=cw;});});
}
function chart(sl,type,data,o){
  sl.addChart(type,data,{chartColors:[T.blue,T.orange,T.green,T.purple],showTitle:false,showLegend:true,legendPos:"b",
    legendFontFace:"Yu Gothic",legendFontSize:9.5,legendColor:T.sub,catAxisLabelFontFace:"Yu Gothic",catAxisLabelFontSize:10,
    catAxisLabelColor:T.sub,valAxisLabelFontFace:"Yu Gothic",valAxisLabelFontSize:9.5,valAxisLabelColor:T.sub,
    valGridLine:{color:"DDE3E8",pt:0.5},catGridLine:{color:T.white,transparency:100},chartArea:{fill:{color:T.white},line:{color:T.white,transparency:100}},
    plotArea:{fill:{color:T.white},line:{color:T.white,transparency:100}},showValue:true,dataLabelColor:T.text,dataLabelFontFace:"Yu Gothic",
    dataLabelFontSize:9.5,...o});
}
function numbered(sl,n,title,body,x,y,w,h,color){
  card(sl,x,y,w,h,title,body,color,T.pale2,{bodySize:13.2});
  shape(sl,S.ellipse,x+w-0.34,y-0.12,0.32,0.32,color,color);
  tx(sl,String(n),{x:x+w-0.34,y:y-0.12,w:0.32,h:0.32,fontSize:13,bold:true,color:T.white,align:"center"});
}

// 1 Title
{
  const sl=pptx.addSlide(); sl.background={color:T.navy}; rect(sl,0,4.70,SW,0.925,T.navy2);
  tx(sl,"3DRef",{x:0.66,y:0.48,w:2.1,h:0.55,fontSize:34,bold:true,color:T.blue2});
  tx(sl,"反射面を含む環境のための\nRGB・LiDAR 3Dデータセットとベンチマーク",{x:0.66,y:1.15,w:8.7,h:1.42,fontSize:F.cover,bold:true,color:T.white,valign:"top",breakLine:true});
  tx(sl,"Xiting Zhao, Sören Schwertfeger — 3DV 2024",{x:0.68,y:2.82,w:8.5,h:0.34,fontSize:F.coverSub,color:"C8D9EA"});
  pill(sl,"論文紹介・詳細版",0.68,3.42,1.85,T.white,"284B70");
  tx(sl,"大阪大学 修士課程セミナー  |  2026年7月23日",{x:0.68,y:4.91,w:8.7,h:0.28,fontSize:14,color:"C8D9EA"});
  notes(sl,["本発表では、3DV 2024の3DRef論文を、データセット設計と評価結果の両面から説明します。","中心となる問いは、反射・透過によってLiDARの幾何仮定が破れる環境を、3Dでどのように正解付けするかです。","論文の表に記載された数値を優先し、本文との不一致がある箇所は明示します。"]);
}

// 2 Problem
{
  const sl=base("反射面では「最初の実物表面から返る」というLiDAR仮定が破れる","Problem");
  const xs=[0.62,2.86,5.10,7.34], colors=[T.green,T.blue,T.orange,T.red];
  const titles=["① 表面で反射","② 透過後の障害物","③ 多重反射・虚像","④ 無返答・欠測"];
  const bodies=["反射材そのものから\n有効な点が返る","ガラスを通過し\n背後の実物を観測","鏡面経由で仮想位置に\n点が生成される","吸収・散乱・視野外で\n点が得られない"];
  xs.forEach((x,i)=>{card(sl,x,1.10,1.98,2.10,titles[i],bodies[i],colors[i],i===2?T.orangePale:T.pale2,{titleSize:14.3,bodySize:13.3});
    arrow(sl,x+0.32,3.55,1.30,0,colors[i],2.5); shape(sl,S.ellipse,x+0.78,3.39,0.30,0.30,colors[i],colors[i]);});
  tx(sl,"同一スキャン内で4現象が混在 → phantom wall・穴・誤った障害物",{x:0.72,y:4.08,w:8.56,h:0.32,fontSize:F.body,bold:true,color:T.navy,align:"center"});
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 1");
  notes(sl,["通常のLiDAR処理は、レーザが最初に当たった実表面から戻ることを暗黙に仮定します。","反射・透明材料では、表面点、透過後の障害物、虚像点、欠測のいずれも起こり得ます。","したがって『反射材の領域』だけでなく、点が実物か虚像か、透明材の背後かまで区別する必要があります。","この誤観測は地図上のphantom wallや穴となり、自己位置推定と経路計画を同時に損ないます。"]);
}

// 3 Contributions
{
  const sl=base("3DRefの貢献は、新規モデルよりも学習・比較を可能にする基盤の構築にある","Contributions");
  rect(sl,0.72,0.98,8.56,0.82,T.pale,T.blue,{shadow:shadow()});
  tx(sl,"研究課題：多様な反射現象をRGBと3D LiDARで精密にラベル化し、既存手法を同一基準で比較できるか",{x:0.95,y:1.13,w:8.1,h:0.45,fontSize:F.body,bold:true,color:T.navy,align:"center"});
  numbered(sl,1,"51,800超の整列データ","48,024点群＋3,799 RGB。3 LiDAR、3シーケンスの屋内データ。",0.72,2.15,2.68,2.05,T.blue);
  numbered(sl,2,"3D meshを正解の核に","手修正した色付きmeshからray castingで点群ラベルと画像maskを生成。",3.66,2.15,2.68,2.05,T.green);
  numbered(sl,3,"単一モダリティの基準値","LiDAR 3手法、RGB 4手法を評価し、return番号の寄与も検証。",6.60,2.15,2.68,2.05,T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Abstract; Sec. 1");
  notes(sl,["論文の主貢献は新しい検出ネットワークではなく、データ・注釈・評価基盤です。","第一に、3種類のLiDARとRGBを時刻・座標系で整列した大規模屋内データを提供します。","第二に、3D meshを共通の正解源として点群と画像のラベルを生成します。","第三に、既存のLiDAR/RGB手法をベンチマークし、multi-return情報が有効かを数値で検証します。"]);
}

// 4 Related datasets
{
  const sl=base("既存データは2D・単一材料が中心で、3D反射ラベルが欠けていた","Related work");
  grid(sl,["Dataset","Modal.","Samples","Target"],[
    ["GDD","RGB","3,900","Glass"],["GSD","RGB","4,102","Glass"],["MSD","RGB","4,018","Mirror"],
    ["Mirror3D","RGB-D","5,894","Mirror"],["TROSD","RGB-D","11,060","Mirror & Glass"],
    [{text:"3DRef",bold:true,color:T.blue},{text:"LiDAR, RGB",bold:true},{text:"48,024 / 3,799",bold:true},{text:"All reflective",bold:true,color:T.blue}]
  ],0.62,1.02,5.55,3.25,[0.27,0.19,0.22,0.32],{cellSize:10.2});
  card(sl,6.44,1.06,2.92,1.30,"差分 1｜3D点ラベル","主要LiDARデータセットには、反射面と虚像点の3Dラベルがない。",T.blue,T.pale,{bodySize:13.2});
  card(sl,6.44,2.55,2.92,1.30,"差分 2｜現象の分離","glass/mirrorだけでなく、reflection pointと背後障害物を区別。",T.orange,T.orangePale,{bodySize:13.2});
  takeaway(sl,"3DRefは「反射物体の2D領域」から「LiDARで生じた3D現象」へ対象を広げる。",T.green);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 1; Sec. 2");
  notes(sl,["Table 1では、従来データの多くがRGBまたはRGB-Dのglass/mirror segmentationです。","3DRefの48,024と3,799はそれぞれ点群とRGB画像で、単純に同一モダリティの標本数として比較しない点に注意します。","本論文の新規性は、材料カテゴリに加えて、虚像点と透明材背後の実障害物を3Dで区別することです。"]);
}

// 5 Multi-return hypothesis
{
  const sl=base("multi-returnは中間反射と背後物体を分ける手掛かりになり得る","Hypothesis");
  tx(sl,"1本のビームから複数のechoを返すLiDAR",{x:0.72,y:0.98,w:3.5,h:0.30,fontSize:F.small,bold:true,color:T.navy});
  shape(sl,S.ellipse,0.85,2.20,0.35,0.35,T.navy,T.navy); tx(sl,"LiDAR",{x:0.62,y:2.62,w:0.82,h:0.25,fontSize:F.label,bold:true,color:T.navy,align:"center"});
  line(sl,2.45,1.25,0,2.85,T.blue,3); tx(sl,"glass",{x:2.04,y:4.17,w:0.82,h:0.26,fontSize:F.label,bold:true,color:T.blue,align:"center"});
  line(sl,5.30,1.25,0,2.85,T.orange,3); tx(sl,"obstacle",{x:4.76,y:4.17,w:1.08,h:0.26,fontSize:F.label,bold:true,color:T.orange,align:"center"});
  arrow(sl,1.20,2.36,1.22,-0.72,T.blue,2); arrow(sl,1.20,2.36,4.08,0.62,T.orange,2);
  pill(sl,"1st return",1.50,1.35,1.25,T.blue,T.pale); pill(sl,"2nd / 3rd return",3.80,3.18,1.75,T.orange,T.orangePale);
  card(sl,6.20,1.15,3.05,1.25,"検証する仮説","return番号を入力特徴に加えると、reflection / obstacle behindの識別が改善する。",T.green,T.greenPale,{bodySize:13.2});
  card(sl,6.20,2.72,3.05,1.25,"ただし注意","return生成規則と密度は機種ごとに異なるため、return番号は普遍的特徴とは限らない。",T.orange,T.orangePale,{bodySize:13.2});
  takeaway(sl,"後段のTable 4・5で、クラス選択性とmIoU改善を別々に確認する。",T.blue);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 1, 3.3, 4.1");
  notes(sl,["multi-return LiDARは、1本の発射に対して複数のechoを記録できます。","透明面や多重反射では、早いreturnと遅いreturnが異なる物理経路に対応する可能性があります。","論文は、return番号を5番目の入力特徴として加える実験を行います。","ただしセンサごとにreturn selectionが異なるため、改善量とセンサ依存性を分けて読む必要があります。"]);
}

// 6 Platform
{
  const sl=base("3種類のLiDARを同一プラットフォームに搭載し、return特性を比較する","Acquisition");
  fitImage(sl,FIG.platform,0.55,0.96,3.45,3.80,2983,3617,"3DRef data collection platform");
  grid(sl,["Sensor","Type / Return","Key specification"],[
    ["Ouster OS0-128","spinning / strongest, 2nd","128 ch, 90° V-FoV, ≤5.2M pts/s"],
    ["Hesai QT64","spinning / first, last","64 ch, 104.2° V-FoV"],
    ["Livox Avia","solid-state / 1st–3rd",">70° FoV, ≤240k pts/s"],
    ["Insta360","dual fisheye RGB","6K, raw 3072×3072 / side"]
  ],4.18,1.08,5.18,2.62,[0.25,0.31,0.44],{cellSize:9.6,headSize:10.5});
  card(sl,4.18,3.91,5.18,0.76,"同期","PTP hardware synchronization：センサ間時刻差を1 ms以内に制御",T.green,T.greenPale,{bodyY:0.39,titleH:0.26,titleSize:13.5,bodySize:12.5});
  cite(sl,"Source: Zhao & Schwertfeger (2024), Fig. 2; Table 2; Sec. 3.1");
  notes(sl,["収集装置は手持ち型で、Ouster、Hesai、Livoxと360度カメラ、IMUを搭載します。","Ousterはstrongest/second strongest、Hesaiはfirst/last、Livoxは最大triple echoであり、同じreturn番号でも選択規則は異なります。","PTPで1ミリ秒以内に同期します。これにより動きながら取得してもセンサ間対応を保ちます。","Ouster・カメラ・IMUはPolar scanner側に統合され、ground-truth mesh作成にも使われます。"]);
}

// 7 Calibration
{
  const sl=base("factory calibration＋FARO整合＋MA-LIOで時空間整列を成立させる","Calibration");
  const steps=[
    ["① Factory calibration","Ouster・camera・IMUの内部／外部校正"],
    ["② Initial extrinsic","3 LiDARをFARO ground-truth fieldへregistration"],
    ["③ PTP sync","hardware時刻同期（≤1 ms）"],
    ["④ MA-LIO","連続時間B-splineでmotion distortion補正"],
    ["⑤ Global alignment","時変extrinsicと残留timestampをオンライン補正"]
  ];
  steps.forEach((s,i)=>{const x=0.62+i*1.78; card(sl,x,1.22,1.55,2.16,s[0],s[1],[T.blue,T.green,T.purple,T.orange,T.red][i],T.pale2,{titleSize:12.4,bodySize:11.5,titleH:0.45,bodyY:0.58}); if(i<4) arrow(sl,x+1.57,2.28,0.18,0,T.muted,1.5);});
  rect(sl,1.08,3.74,7.84,0.54,T.pale,T.blue,{shadow:shadow()});
  tx(sl,"出力：undistorted・global座標系に整列した3 LiDAR点群＋カメラ姿勢",{x:1.30,y:3.86,w:7.40,h:0.30,fontSize:F.body,bold:true,color:T.navy,align:"center"});
  takeaway(sl,"注釈精度はmeshだけでなく、pose・extrinsic・timestampの整合精度にも依存する。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 3.1; MA-LIO description");
  notes(sl,["3Dラベル生成では、LiDAR原点からmeshへrayを飛ばすため、姿勢誤差は直接ラベル誤差になります。","初期外部校正は、3台のLiDAR点群をFARO ground-truth fieldへregistrationして求めます。","MA-LIOは連続時間B-spline軌跡を用い、動き歪みを補正すると同時に、時変extrinsicと残留時刻ずれをオンライン推定します。","つまり、本データセットの品質は幾何再構成とセンサ整列の二段階に支えられています。"]);
}

// 8 Sequences
{
  const sl=base("3シーケンスは反射材・視点・入射角の多様性を意図して収集された","Sequences");
  const imgs=[[FIG.seq1,1496,1122],[FIG.seq2,1496,1121],[FIG.seq3,1496,1121]];
  const heads=["Seq.1｜鏡・床から天井までの窓","Seq.2｜会議室・全面ガラス","Seq.3｜手すり・掲示・TV・白板"];
  const counts=["3,732 frames/LiDAR ＋ 541 RGB","4,702 frames/LiDAR ＋ 1,716 RGB","7,574 frames/LiDAR ＋ 1,542 RGB"];
  for(let i=0;i<3;i++){const x=0.55+i*3.0; fitImage(sl,imgs[i][0],x,1.02,2.75,2.08,imgs[i][1],imgs[i][2],heads[i]);
    tx(sl,heads[i],{x,y:3.20,w:2.75,h:0.38,fontSize:12.5,bold:true,color:T.navy,align:"center"});
    pill(sl,counts[i],x+0.15,3.70,2.45,i===2?T.orange:T.blue,i===2?T.orangePale:T.pale);}
  takeaway(sl,"同じ場所を複数回・複数方向から通過し、距離と入射角を変えて観測する。",T.green);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Figs. 4–6; Sec. 3.1, 3.3");
  notes(sl,["収集者は手持ち装置で屋内を歩き、同一反射面を複数の視点・距離・入射角から観測します。","Sequence 1は鏡と大きな窓、Sequence 2は全面ガラスの会議室、Sequence 3は白板・TV・ガラス手すりなど多様な材質を含みます。","図中の赤線は軌跡です。論文の可視化では0.1 m subsamplingを行い、天井を除去しています。"]);
}

// 9 Mesh GT
{
  const sl=base("正解の中心は、反射部を手修正・色分けした高精度3D meshである","Ground truth");
  fitImage(sl,FIG.overview,0.58,0.93,4.10,3.76,891,1049,"3DRef overview and labeled mesh");
  card(sl,4.98,1.02,4.28,0.90,"① Polar scannerでtextured mesh生成","Ouster＋camera＋IMUの計測から、環境の幾何とtextureを再構成。",T.blue,T.pale,{bodySize:12.5,titleSize:14});
  card(sl,4.98,2.10,4.28,0.90,"② 反射・ガラス由来の穴を手修正","meshの欠損や誤形状を人手でcleaningし、正解幾何を補完。",T.orange,T.orangePale,{bodySize:12.5,titleSize:14});
  card(sl,4.98,3.18,4.28,0.90,"③ texture上で反射材料を色分け","glass / mirror / other reflectiveを手作業で領域指定。",T.green,T.greenPale,{bodySize:12.5,titleSize:14});
  takeaway(sl,"点ごとの人手注釈ではなく、1つの3D scene modelを全modalitiesの正解源にする。",T.blue);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Fig. 1; Sec. 3.2");
  notes(sl,["注釈の起点は、Polar scannerから生成したtexture付き3D meshです。","ガラスや鏡によってmeshに穴や誤形状が生じるため、著者はmeshを手作業でcleaningし、欠損を修正します。","さらにtexture上でglass、mirror、other reflectiveを色分けします。","これにより、点群と画像の各フレームを独立に塗るのではなく、一貫した3D scene modelからラベルを投影できます。"]);
}

// 10 Ray casting
{
  const sl=base("LiDAR点は『センサ原点→観測点』のrayとmeshの交点で自動ラベル化する","Point labeling");
  shape(sl,S.ellipse,0.72,2.08,0.38,0.38,T.navy,T.navy); tx(sl,"LiDAR pose",{x:0.52,y:2.55,w:0.78,h:0.28,fontSize:F.label,bold:true,color:T.navy,align:"center"});
  line(sl,4.15,1.02,0,3.12,T.blue,3); tx(sl,"labeled mesh",{x:3.55,y:4.22,w:1.20,h:0.28,fontSize:F.label,bold:true,color:T.blue,align:"center"});
  shape(sl,S.ellipse,6.08,2.72,0.28,0.28,T.orange,T.orange); tx(sl,"measured point",{x:5.54,y:3.10,w:1.36,h:0.28,fontSize:F.label,bold:true,color:T.orange,align:"center"});
  arrow(sl,1.08,2.26,5.02,0.60,T.orange,2.3); shape(sl,S.ellipse,4.02,2.58,0.26,0.26,T.green,T.green);
  pill(sl,"first mesh intersection",3.26,2.92,1.78,T.green,T.greenPale);
  card(sl,6.88,1.08,2.36,1.20,"入力","点座標・LiDAR pose・色付きmesh",T.blue,T.pale,{bodySize:12.5});
  card(sl,6.88,2.56,2.36,1.20,"出力","交差面の材料＋点の実在性に基づくclass ID",T.green,T.greenPale,{bodySize:12.5});
  takeaway(sl,"Open3D ray castingにより、全フレームへ一貫した3Dラベルを展開する。",T.blue);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 3.2");
  notes(sl,["各点について、LiDARの観測姿勢をmesh座標系へ変換し、センサ原点から観測点方向へrayを飛ばします。","rayが最初に交差したmesh面の色・種類を参照してラベルを割り当てます。","この処理はOpen3D ray castingで自動化されています。","境界付近ではpose・mesh・calibrationの誤差がラベルに影響するため、完全な自動正解ではない点には注意が必要です。"]);
}

// 11 RGB masks
{
  const sl=base("RGB maskも同じmeshをcamera viewへrenderして生成する","Image labeling");
  const boxes=[["Camera pose + intrinsics",0.65,T.blue],["Labeled 3D mesh",2.62,T.green],["View rendering",4.59,T.purple],["Class masks",6.56,T.orange]];
  boxes.forEach((b,i)=>{card(sl,b[1],1.32,1.55,1.25,b[0],i===0?"各frameの姿勢・内部パラメータ":i===1?"材料classをtextureへ符号化":i===2?"camera視点へ投影": "glass / mirror / other / all",b[2],T.pale2,{titleSize:12.6,bodySize:11.3,titleH:0.44,bodyY:0.55}); if(i<3) arrow(sl,b[1]+1.58,1.95,0.34,0,T.muted,1.6);});
  card(sl,1.10,3.10,3.55,1.00,"魚眼画像の扱い","既存RGB segmentation手法と互換にするため、fisheyeをpinhole画像へundistort。",T.orange,T.orangePale,{bodySize:13});
  card(sl,5.35,3.10,3.55,1.00,"modalities間の一貫性","点群と画像が同じ3D mesh由来のため、class定義と空間位置を共有。",T.green,T.greenPale,{bodySize:13});
  takeaway(sl,"本論文の本質は、meshを介してRGBとLiDARの正解を同じ世界座標に結びつけた点。",T.blue);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 3.2");
  notes(sl,["RGBでは、各カメラ姿勢と内部パラメータを用いて、色付きmeshを画像平面へrenderします。","render結果からglass、mirror、other reflective、および全反射材のmaskを抽出します。","360度魚眼画像は、既存のpinhole前提ネットワークへ入力するためundistortされます。","LiDARとRGBが同じmeshを正解源とすることが、将来のsensor fusionにとって重要です。"]);
}

// 12 Labels
{
  const sl=base("7個のIDのうち、0を除く6クラスが意味のある評価対象となる","Label taxonomy");
  const rows=[
    ["0","Unlabeled","mesh外・正解なし",T.gray],
    ["1","Normal","通常の実表面",T.navy],
    ["2","Glass","ガラス面",T.blue],
    ["3","Mirror","鏡面",T.purple],
    ["4","Other reflective","アクリル、白板、TV、monitor、glazed tile等",T.green],
    ["5","Reflection point","反射経路で生じた虚像点（実在表面ではない）",T.orange],
    ["6","Obstacle behind glass","透明ガラスを透過して得た背後障害物",T.red]
  ];
  rows.forEach((r,i)=>{const y=0.96+i*0.47; shape(sl,S.ellipse,0.75,y+0.06,0.28,0.28,r[3],r[3]); tx(sl,r[0],{x:0.75,y:y+0.06,w:0.28,h:0.28,fontSize:11,bold:true,color:T.white,align:"center"});
    tx(sl,r[1],{x:1.22,y,w:2.05,h:0.38,fontSize:14,bold:true,color:r[3]}); tx(sl,r[2],{x:3.38,y,w:5.72,h:0.38,fontSize:13.2,color:T.sub});
    if(i<6) line(sl,1.20,y+0.43,7.92,0,T.line,0.6);});
  takeaway(sl,"“6 classes”と説明する場合も、配布ラベルにはID 0（Unlabeled）が含まれる。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 3.2; dataset label definition");
  notes(sl,["配布ラベルは0から6までの7 IDです。","0はmesh外などのUnlabeledで、通常は評価対象から外します。したがって意味クラスはNormalを含む6クラスです。","重要なのは、材料クラス2〜4と、点の生成現象を表す5〜6を区別していることです。","Reflection pointは虚像点、Obstacle behind glassは実在する背後物体という幾何学的に異なる意味を持ちます。"]);
}

// 13 Size and structure
{
  const sl=base("48,024点群と3,799画像を、raw・RGB・SemanticKITTI形式で公開する","Dataset scale");
  callout(sl,"48,024","labeled point clouds\n= 16,008 × 3 LiDAR",0.72,1.03,2.45,1.52,T.blue);
  callout(sl,"3,799","labeled RGB images",3.42,1.03,2.20,1.52,T.green);
  callout(sl,"3","indoor sequences",5.87,1.03,1.45,1.52,T.orange);
  callout(sl,"7 IDs","0 unlabeled + 6 classes",7.57,1.03,1.72,1.52,T.purple);
  grid(sl,["Package","Contents","Purpose"],[
    ["Raw","poses, images, meshes, calibration, labeled clouds","再現・再処理"],
    ["RGB","glass / mirror / other / all masks, train/test","2D benchmark"],
    ["SemanticKITTI","XYZI (4ch), XYZIR (5ch)","3D segmentation"],
    ["Scripts / Networks","ray tracing, stats, eval, code/weights","再現性"]
  ],0.72,2.86,8.56,1.42,[0.20,0.52,0.28],{cellSize:10.2});
  takeaway(sl,"点群数は3 LiDARの同時系列を含むため、独立scene数ではなくframe×sensor数として解釈する。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 3.3; project release structure");
  notes(sl,["各LiDARについて16,008 frame、合計48,024点群です。RGBは3,799画像です。","3台は同じ軌跡を同時観測するため、48,024を独立な環境数として解釈してはいけません。","SemanticKITTI互換形式では、XYZIの4チャネルとreturn番号を追加したXYZIRの5チャネルを提供します。","rawデータ、mesh、calibration、ray tracingと評価scriptも公開対象です。"]);
}

// 14 Sensor imbalance
{
  const sl=base("class分布はLiDAR機種とsequenceに大きく依存する","Dataset analysis");
  chart(sl,C.bar,[
    {name:"Normal",labels:["Ouster S1","Hesai S1","Livox S1","Ouster S3","Hesai S3","Livox S3"],values:[84.19,76.67,52.23,87.05,76.97,55.78]},
    {name:"Reflection",labels:["Ouster S1","Hesai S1","Livox S1","Ouster S3","Hesai S3","Livox S3"],values:[9.31,10.77,29.22,1.67,3.41,14.34]},
    {name:"Obstacle",labels:["Ouster S1","Hesai S1","Livox S1","Ouster S3","Hesai S3","Livox S3"],values:[0.66,1.54,7.62,3.64,5.46,19.28]}
  ],{x:0.58,y:1.02,w:6.15,h:3.26,barDir:"col",catAxisLabelRotate:315,valAxisMinVal:0,valAxisMaxVal:100,valAxisMajorUnit:20,showLegend:true,showValue:false});
  card(sl,6.95,1.04,2.36,1.18,"Livox Seq.1","Reflection 29.22%\nObstacle 7.62%",T.orange,T.orangePale,{bodySize:13.5});
  card(sl,6.95,2.48,2.36,1.18,"Livox Seq.3","Reflection 14.34%\nObstacle 19.28%",T.red,"F9E7E7",{bodySize:13.5});
  takeaway(sl,"モデル比較だけでなく、sensor-specific return policyとclass priorの影響を考える必要がある。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 3（抜粋）");
  notes(sl,["Table 3の代表値を可視化しています。Normalが多数派ですが、LivoxではReflectionとObstacleの比率が大きくなります。","Sequence 1のLivoxはReflection 29.22%、Sequence 3ではObstacle behind glassが19.28%です。","同じ環境でもreturn selectionやsampling patternによりラベル分布が変わります。","したがって、センサ別成績は単純なハードウェア性能だけでなく、入力分布の差も含みます。"]);
}

// 15 Return distribution
{
  const sl=base("late returnはReflection / Obstacleに選択的だが、完全な識別子ではない","Multi-return");
  grid(sl,["Class","1st return","2nd return","3rd return"],[
    ["Normal","99.02","0.97","0.01"],["Glass","99.61","0.39","0.00"],["Mirror","98.74","1.26","0.00"],
    ["Other reflective","99.82","0.18","0.00"],
    [{text:"Reflection",bold:true,color:T.orange},{text:"77.16",bold:true},{text:"22.07",bold:true,color:T.orange},{text:"0.77",bold:true,color:T.orange}],
    [{text:"Obstacle",bold:true,color:T.red},{text:"62.20",bold:true},{text:"33.38",bold:true,color:T.red},{text:"4.43",bold:true,color:T.red}]
  ],0.72,1.02,5.62,3.18,[0.34,0.22,0.22,0.22],{cellSize:11.3});
  card(sl,6.62,1.10,2.66,1.28,"2nd+3rdの比率","Normal 0.98%\nReflection 22.84%\nObstacle 37.81%",T.green,T.greenPale,{bodySize:14});
  card(sl,6.62,2.68,2.66,1.28,"解釈","late returnは反射現象に濃縮されるが、1stにも多数残る。",T.orange,T.orangePale,{bodySize:13.4});
  takeaway(sl,"return番号は有用な補助特徴だが、番号だけでclassを決定できない。",T.blue);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 4");
  notes(sl,["Table 4は各クラスの点が何番目のreturnで観測されたかを示します。","Normal、Glass、Mirror、Other reflectiveは98.7%以上がfirst returnです。","一方、Reflectionは22.84%、Obstacleは37.81%がsecondまたはthird returnです。","ただしReflectionの77.16%、Obstacleの62.20%はfirst returnなので、return番号は決定規則ではなく補助特徴です。"]);
}

// 16 incidence angle
{
  const sl=base("入射角により直接反射・透過・多重反射の割合が変化する","Incidence angle");
  fitImage(sl,FIG.angle,0.72,1.00,5.10,3.62,800,637,"Incident-angle distribution by class");
  card(sl,6.10,1.06,3.15,0.98,"計算方法","mesh法線とbeam方向から入射角を計算し、角度bin内を100%に正規化。",T.blue,T.pale,{bodySize:12.5});
  card(sl,6.10,2.22,3.15,0.98,"観察","glass / mirrorの直接returnは低角度で多く、角度増加とともに減少。",T.green,T.greenPale,{bodySize:12.5});
  card(sl,6.10,3.38,3.15,0.98,"含意","材料・角度・視点が結合するため、複数方向からの収集が不可欠。",T.orange,T.orangePale,{bodySize:12.5});
  cite(sl,"Source: Zhao & Schwertfeger (2024), Fig. 7; Sec. 3.3");
  notes(sl,["入射角はmesh法線とLiDAR beam方向から計算します。図は各角度bin内のclass構成を100%に正規化したものです。","Normalはgrazing方向まで増加した後、極端な角度で減少します。","GlassとMirrorの直接returnは比較的小さい角度で多く、角度が大きくなると減少します。","この結果は、単一視点だけでは材料特性と角度効果を分離しにくいことを示します。"]);
}

// 17 benchmark setup
{
  const sl=base("ベンチマークはLiDAR segmentationとRGB segmentationを別々に評価する","Benchmark setup");
  card(sl,0.65,1.00,4.15,2.82,"LiDAR｜PCSeg codebase","Models：MinkowskiNet / SPVCNN / Cylinder3D\nInput A：XYZI（4ch）\nInput B：XYZIR（5ch, return番号追加）\nMetric：class-wise IoU / mIoU",T.blue,T.pale,{bodySize:15});
  card(sl,5.20,1.00,4.15,2.82,"RGB｜既存glass / mirror手法","Glass：GlassSemNet / EBLNet\nMirror：HetNet / SATNet\nPretrained modelを3DRefへ直接適用\nSATNet・EBLNetは3DRefで再学習",T.green,T.greenPale,{bodySize:15});
  pill(sl,"Hardware：dual NVIDIA RTX 3090（24 GB ×2）",2.47,4.05,5.05,T.purple,"EEE9F3");
  takeaway(sl,"この論文はmultimodal fusionを実験していない。RGBとLiDARの単独baselineを提示する。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 4");
  notes(sl,["LiDARではPCSeg codebase上の3つのsemantic segmentation手法を使用します。","return番号なしのXYZIと、return番号を加えたXYZIRを比較します。","RGBではglass用とmirror用の既存モデルを、pretrainedのまま評価し、一部を3DRefで再学習します。","論文はRGB-LiDAR fusionを将来課題とし、現在のベンチマークは単一モダリティです。"]);
}

// 18 LiDAR overall
{
  const sl=base("LiDARではCylinder3Dが最高、return追加の全体改善は小さい","LiDAR results");
  grid(sl,["Method","Return","mIoU","Normal","Glass","Mirror","Other","Reflect.","Obstacle"],[
    ["Minkowski","No","81.55","96.52","69.40","76.51","84.96","84.45","77.45"],
    ["Minkowski","Yes","82.11","96.75","70.27","74.38","84.95","86.44","79.90"],
    ["SPVCNN","No","81.85","96.53","69.77","77.77","84.15","84.72","78.16"],
    ["SPVCNN","Yes","82.14","96.67","69.42","77.60","86.42","84.89","77.86"],
    [{text:"Cylinder3D",bold:true,color:T.blue},"No",{text:"83.72",bold:true},"96.84","69.78","79.87","83.76","88.58","83.49"],
    [{text:"Cylinder3D",bold:true,color:T.blue},{text:"Yes",bold:true,color:T.green},{text:"83.92",bold:true,color:T.green},"96.96","70.02","79.68","83.58",{text:"89.28",bold:true},{text:"83.99",bold:true}]
  ],0.52,1.00,8.96,2.85,[0.16,0.09,0.10,0.11,0.10,0.10,0.11,0.11,0.12],{cellSize:8.6,headSize:9.2});
  callout(sl,"83.92","best mIoU\nCylinder3D + return",0.82,4.00,2.08,0.78,T.green);
  card(sl,3.20,3.98,2.75,0.82,"Normalは容易","全手法で96.5%以上。class imbalanceの影響も考慮。",T.blue,T.pale,{titleSize:13,bodyY:0.38,bodySize:10.7});
  card(sl,6.25,3.98,2.75,0.82,"Glassが難しい","全体評価では約69–70 IoUで頭打ち。",T.orange,T.orangePale,{titleSize:13,bodyY:0.38,bodySize:10.7});
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 5 (All sensors)");
  notes(sl,["Table 5のAll sensors部分を再掲しています。","最高mIoUはCylinder3Dにreturn番号を加えた83.9188です。","Normalは96%以上ですが、Glassは約70で、材料面の識別が相対的に難しいです。","return追加は全体mIoUを改善するものの、全クラス・全モデルで一様に改善するわけではありません。"]);
}

// 19 delta
{
  const sl=base("return番号のmIoU改善は+0.20〜+0.57 pointで、クラス別には負の効果もある","Return ablation");
  chart(sl,C.bar,[{name:"ΔmIoU (Yes − No)",labels:["MinkowskiNet","SPVCNN","Cylinder3D"],values:[0.5652,0.2951,0.1992]}],
    {x:0.72,y:1.08,w:5.18,h:2.92,barDir:"col",showLegend:false,valAxisMinVal:0,valAxisMaxVal:0.7,valAxisMajorUnit:0.1,dataLabelPosition:"outEnd",dataLabelFormatCode:"0.00",showValue:true});
  callout(sl,"+0.35","table-derived mean\npercentage points",6.20,1.13,2.88,1.35,T.green);
  card(sl,6.20,2.73,2.88,1.18,"非一様な効果","Minkowski mirror −2.13、SPVCNN glass/obstacleも低下。",T.orange,T.orangePale,{bodySize:12.8});
  takeaway(sl,"本文の“0.5% across methods”は概括表現。Table 5の単純平均は+0.353 point。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 5; Sec. 4.1（差分は表から計算）");
  notes(sl,["return有無のmIoU差は、MinkowskiNetで+0.5652、SPVCNNで+0.2951、Cylinder3Dで+0.1992 pointです。","3手法の単純平均は+0.3532 pointです。本文の『across methodsで0.5%改善』は丸めまたは概括的表現と考えられます。","クラス別にはMinkowskiNetのMirrorが約2.13 point低下するなど、効果は非一様です。","したがって、return番号の情報量は確認できるものの、強い決定的改善とは言えません。"]);
}

// 20 sensor results
{
  const sl=base("同じCylinder3Dでもsensor別mIoUは79.71〜85.96と差がある","Sensor-wise results");
  grid(sl,["Sensor","mIoU","Normal","Glass","Mirror","Other","Reflect.","Obstacle"],[
    ["Ouster","79.71","97.90","58.27","59.35","88.08",{text:"90.51",bold:true,color:T.blue},"84.18"],
    ["Hesai","84.98","96.51",{text:"78.69",bold:true,color:T.green},{text:"88.46",bold:true,color:T.green},"80.81","86.54","78.88"],
    [{text:"Livox",bold:true,color:T.orange},{text:"85.96",bold:true,color:T.orange},"93.59","72.64","85.73","87.85","89.37",{text:"86.59",bold:true,color:T.orange}]
  ],0.72,1.08,8.56,1.72,[0.16,0.11,0.12,0.12,0.12,0.12,0.13,0.12],{cellSize:10.5});
  card(sl,0.82,3.13,2.52,1.02,"Ouster","Normal/Reflectionは高いが、Glass・Mirrorが低い。",T.blue,T.pale,{bodySize:12.4});
  card(sl,3.73,3.13,2.52,1.02,"Hesai","Glass・Mirrorが最良。Obstacleは3機種中で低い。",T.green,T.greenPale,{bodySize:12.4});
  card(sl,6.64,3.13,2.52,1.02,"Livox","最高mIoU。late returnが多く、Obstacleも高い。",T.orange,T.orangePale,{bodySize:12.4});
  takeaway(sl,"sensor差はreturn特性・点密度・class priorが混在した結果で、純粋な機種優劣ではない。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 5 (Cylinder3D + return)");
  notes(sl,["Cylinder3D＋returnをsensor別に見ると、Ouster 79.71、Hesai 84.98、Livox 85.96です。","OusterはGlassとMirrorが低く、Hesaiは両クラスが高い一方、Obstacleは低めです。","Livoxはlate return比率が高く、Obstacle IoUも86.59です。","ただし点密度、FoV、return rule、class分布が同時に変わるため、センサの因果的比較ではありません。"]);
}

// 21 RGB
{
  const sl=base("RGBモデルは既存データから3DRefへ移すと大きく低下し、再学習で回復する","RGB results");
  grid(sl,["Method","Training data","Original test","3DRef target","3DRef mIoU"],[
    ["GlassSemNet","GDD","90.80","Glass","53.69"],
    ["HetNet","PMD","69.00","Mirror","44.05"],
    ["SATNet","RGBD","78.42","Mirror","49.46"],
    ["EBLNet","GDD","88.72","Glass","60.49"],
    ["EBLNet","MSD","80.33","Mirror","57.61"],
    [{text:"SATNet*",bold:true,color:T.green},"3DRef","—","Mirror / All",{text:"82.47 / 68.81",bold:true,color:T.green}],
    [{text:"EBLNet*",bold:true,color:T.blue},"3DRef","—","Glass / All",{text:"86.71 / 87.60",bold:true,color:T.blue}]
  ],0.72,1.00,8.56,3.42,[0.18,0.18,0.18,0.22,0.24],{cellSize:10.5});
  takeaway(sl,"最大の知見は、既存データで高得点でも3DRefへ直接generalizeしないというdomain gap。",T.orange);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 6; * = retrained on 3DRef");
  notes(sl,["pretrained modelを3DRefへ直接適用すると、GlassSemNetは90.80から53.69、SATNetは78.42から49.46へ低下します。","これは屋内環境、撮像条件、材料の多様性、mask生成法の差によるdomain gapを示します。","3DRefで再学習すると、SATNet Mirror 82.47、EBLNet Glass 86.71まで回復します。","All reflectiveでは表上SATNet 68.81、EBLNet 87.60です。"]);
}

// 22 evidence reading
{
  const sl=base("結果は有用だが、論文本文と表の不一致を分けて読む必要がある","Critical reading");
  card(sl,0.68,1.05,4.10,1.48,"Table 5：return改善量","表からの差分平均は+0.353 point。本文は“0.5% across methods”と記述。丸め方が明示されない。",T.orange,T.orangePale,{bodySize:13.4});
  card(sl,5.22,1.05,4.10,1.48,"Table 6：87.60の帰属","表ではEBLNet*の3DRef-All=87.60。本文は“Retrained SATNet ... 87.6%”と記述し、手法名が不一致。",T.red,"F9E7E7",{bodySize:13.4});
  card(sl,0.68,2.90,4.10,1.22,"論文が支持する主張","3D labelsを構築できる／baselineを提示できる／late returnが反射現象に偏る。",T.green,T.greenPale,{bodySize:13.4});
  card(sl,5.22,2.90,4.10,1.22,"まだ支持しない主張","RGB-LiDAR fusionの有効性／未知環境での汎化／sensor差の因果説明。",T.blue,T.pale,{bodySize:13.4});
  takeaway(sl,"教授向け説明では、著者の結論と表から直接読める証拠を明確に区別する。",T.navy);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Sec. 4.1–4.2; Tables 5–6");
  notes(sl,["ここでは批判ではなく、証拠の読み分けを行います。","return改善量について、本文の0.5%とTable 5からの単純平均0.353 pointは一致しません。","RGBについて、87.60は表ではEBLNetのAllですが、本文はSATNetと記述しています。手法名の誤記の可能性があります。","したがって発表では、具体値は表を一次情報として提示し、本文表現は別に紹介します。"]);
}

// 23 limitations future
{
  const sl=base("3DRefは基盤として強い一方、環境・注釈・評価設計に限界が残る","Limitations");
  card(sl,0.65,1.02,2.62,2.82,"① Coverage","屋内3 sequenceが中心。屋外、水面、雨天、車載速度、多様な材質・経年変化は未評価。",T.orange,T.orangePale,{bodySize:14});
  card(sl,3.69,1.02,2.62,2.82,"② Ground truth","mesh修正は人手依存。ray castingは境界、薄い面、pose/extrinsic誤差に敏感。",T.red,"F9E7E7",{bodySize:14});
  card(sl,6.73,1.02,2.62,2.82,"③ Benchmark","単一モダリティ・segmentation中心。汎化、navigation改善、fusionの実証は未実施。",T.blue,T.pale,{bodySize:14});
  pill(sl,"著者のfuture work：センサ・材質・環境拡張／RGB-LiDAR fusion／depth・polarization／self・semi-supervised",1.05,4.08,7.90,T.green,T.greenPale);
  takeaway(sl,"次の研究課題は、より多様な条件での汎化と、検出がmapping/navigationをどれだけ改善するかの評価。",T.green);
  cite(sl,"Source: Zhao & Schwertfeger (2024), Conclusion; limitations are presentation analysis");
  notes(sl,["論文の明示的な将来課題は、センサ・材質・環境の拡張、multimodal fusion、depth/polarization、self/semi-supervised learningです。","追加で、3 sequenceの屋内データから未知環境への汎化をどこまで主張できるかは限定的です。","mesh修正とray castingは高効率ですが、正解がpose・calibration・mesh品質に依存します。","最終的には、segmentation精度だけでなく、SLAMやnavigationの安全性改善へつながるかを評価する必要があります。"]);
}

// 24 references
{
  const sl=base("参考文献","References");
  const refs=[
    "[1] X. Zhao and S. Schwertfeger, “3DRef: 3D Dataset and Benchmark for Reflection Detection in RGB and Lidar Data,” 3DV, 2024. arXiv:2403.06538.",
    "[2] Y. Mei et al., “Don't Hit Me! Glass Detection in Real-world Scenes,” CVPR, 2020. (GDD / GlassSemNet)",
    "[3] H. Lin et al., “Mirror Detection via Semantic-aware Context Contrast,” AAAI, 2020. (MSD)",
    "[4] J. Xu et al., “Mirror3D: Depth Refinement for Mirror Surfaces,” CVPR, 2021.",
    "[5] L. Fan et al., “EBLNet: Edge-aware Boundary Localization Network for Glass-like Object Segmentation,” ICCV, 2021.",
    "[6] R. Cheng et al., “SATNet: Symmetrical Attention Transfer Network for Dichromatic Image Segmentation,” ACM MM, 2022.",
    "[7] Y. Zhang et al., “PCSeg: An Open Source Point Cloud Segmentation Codebase,” 2023.",
    "Dataset / code: http://3dref.github.io"
  ];
  refs.forEach((r,i)=>{tx(sl,r,{x:0.72,y:0.96+i*0.47,w:8.56,h:0.38,fontSize:12.2,color:i===0?T.navy:T.sub,bold:i===0,valign:"top"}); if(i<7) line(sl,0.72,1.36+i*0.47,8.56,0,T.line,0.5);});
  notes(sl,["主論文と、発表で名前を挙げた代表的なRGBおよびpoint-cloud segmentation関連文献です。","データとコードは3dref.github.ioで公開されています。","質疑では、Table 1に掲載された全関連データセットとTable 5・6の元論文を必要に応じて参照します。"]);
}

// 25 conclusion
{
  const sl=base("結論：3DRefは反射を『材料』と『点の生成現象』の両方で学習可能にした","Conclusion");
  numbered(sl,1,"3D正解の設計","修正済みtextured meshを共通基盤に、点群とRGB maskを一貫生成。",0.72,1.05,2.68,2.55,T.blue);
  numbered(sl,2,"multi-returnの証拠","late returnはReflection / Obstacleに偏る。ただしmIoU改善は平均+0.35 point。",3.66,1.05,2.68,2.55,T.green);
  numbered(sl,3,"benchmarkの意味","既存RGBモデルのdomain gapを示し、LiDAR/RGB研究の再現可能な基準値を提供。",6.60,1.05,2.68,2.55,T.orange);
  rect(sl,0.72,3.94,8.56,0.54,T.navy,T.navy,{shadow:shadow()});
  tx(sl,"残る問い：この検出結果を使うと、反射環境でのSLAM・mapping・navigationはどこまで安全になるか？",{x:0.96,y:4.05,w:8.08,h:0.32,fontSize:F.body,bold:true,color:T.white,align:"center"});
  cite(sl,"Summary based on Zhao & Schwertfeger (2024)");
  notes(sl,["第一の結論は、3D meshを共通正解源にした注釈設計がデータセットの核だという点です。","第二に、late returnは反射現象へ選択的ですが、return番号追加のmIoU改善は限定的で、補助特徴と理解すべきです。","第三に、既存RGBモデルは3DRefへ直接汎化せず、データセット固有の再学習が重要です。","本研究を次へ進めるには、検出精度がSLAMとnavigationの失敗を実際にどれだけ減らすかを測る必要があります。"]);
}

// 26 appendix Table 3
{
  const sl=base("付録A｜sensor・sequence別のclass分布（%）","Appendix");
  grid(sl,["Seq","Sensor","Normal","Glass","Mirror","Other","Reflect.","Obstacle","Return 1/2/3"],[
    ["1","Ouster","84.19","0.91","0.61","2.85","9.31","0.66","99.72/0.28/—"],
    ["1","Hesai","76.67","3.16","3.41","2.60","10.77","1.54","92.60/7.40/—"],
    ["1","Livox","52.23","4.73","1.72","2.42","29.22","7.62","86.30/11.63/2.07"],
    ["2","Ouster","87.75","2.59","0.04","3.23","2.41","2.45","99.71/0.29/—"],
    ["2","Hesai","78.64","7.48","0.28","3.14","3.79","4.16","93.82/6.17/—"],
    ["2","Livox","68.71","6.34","0.13","1.73","11.60","9.48","92.66/6.90/0.42"],
    ["3","Ouster","87.05","2.17","—","2.76","1.67","3.64","99.62/0.38/—"],
    ["3","Hesai","76.97","7.76","—","2.14","3.41","5.46","92.13/7.87/—"],
    ["3","Livox","55.78","5.33","—","1.29","14.34","19.28","83.49/14.33/2.17"]
  ],0.42,0.90,9.16,3.78,[0.055,0.105,0.105,0.085,0.085,0.09,0.10,0.105,0.27],{cellSize:8.1,headSize:8.7});
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 3. Unlabeled share is omitted here, as in the selected columns.");
  notes(sl,["Table 3の主要列を付録として一覧化しています。","Ousterのsecond returnは全sequenceで0.4%未満、Livoxはsecond/thirdが最大16.5%程度です。","MirrorはSequence 3には含まれません。","クラス分布とreturn分布がsensor・sequenceの両方で変化することを確認できます。"]);
}

// 27 appendix class deltas
{
  const sl=base("付録B｜return追加によるclass-wise IoU差（point）","Appendix");
  grid(sl,["Method","ΔmIoU","ΔNormal","ΔGlass","ΔMirror","ΔOther","ΔReflection","ΔObstacle"],[
    ["Minkowski","+0.565","+0.229","+0.875",{text:"−2.134",color:T.red,bold:true},"−0.015",{text:"+1.984",color:T.green,bold:true},{text:"+2.452",color:T.green,bold:true}],
    ["SPVCNN","+0.295","+0.143",{text:"−0.354",color:T.red},"−0.171",{text:"+2.271",color:T.green,bold:true},"+0.175","−0.294"],
    ["Cylinder3D","+0.199","+0.120","+0.235","−0.189","−0.175",{text:"+0.703",color:T.green,bold:true},"+0.502"]
  ],0.62,1.06,8.76,1.86,[0.18,0.115,0.115,0.115,0.115,0.115,0.13,0.115],{cellSize:10.4});
  card(sl,0.72,3.28,4.05,1.00,"改善が安定するclassはない","Reflectionは3手法すべて改善するが、他classは符号が変わる。",T.orange,T.orangePale,{bodySize:13});
  card(sl,5.23,3.28,4.05,1.00,"平均値だけでは不十分","return番号はmodelがどう利用するかに依存し、特定classを悪化させ得る。",T.blue,T.pale,{bodySize:13});
  cite(sl,"Source: Zhao & Schwertfeger (2024), Table 5. Differences calculated from reported values.");
  notes(sl,["Table 5のreturn有無からclass-wise差分を計算したものです。","Reflectionは3手法で正ですが、Glass、Mirror、Other、Obstacleは手法により正負が変わります。","MinkowskiNetではMirrorが2.134 point低下する一方、Obstacleは2.452 point改善します。","return featureを加える際は、全体mIoUだけでなくclass-wise trade-offを確認すべきです。"]);
}

// 28 appendix reproducibility
{
  const sl=base("付録C｜再現性を確認するときのチェックポイント","Appendix");
  const items=[
    ["Data unit","48,024は16,008 timestamps×3 LiDAR。独立scene数ではない。",T.blue],
    ["Input","XYZIとXYZIRを区別し、return番号の符号化を確認する。",T.green],
    ["Labels","ID 0はUnlabeled。mIoUの対象classとignore設定を確認する。",T.orange],
    ["Split","RGB train/testとLiDAR sequence/sensorの分割単位を確認する。",T.purple],
    ["Geometry","mesh・pose・extrinsic・timestamp補正がray labelへ直結する。",T.red],
    ["Claims","本文の概括値ではなくTable 5・6の値とmethod名を参照する。",T.navy]
  ];
  items.forEach((it,i)=>{const col=i%2,row=Math.floor(i/2); card(sl,0.68+col*4.55,0.98+row*1.23,4.10,1.00,it[0],it[1],it[2],T.pale2,{bodySize:12.5,titleSize:14});});
  takeaway(sl,"再現実験では『同じframeの別sensorがtrain/testを跨いでいないか』も確認すべき。",T.orange);
  cite(sl,"Based on dataset structure and experimental description in Zhao & Schwertfeger (2024)");
  notes(sl,["再現性確認で特に重要な点を整理しています。","同時刻の3センサ点群は相関が高いため、分割単位によっては情報漏洩に近い状況が起こり得ます。論文と公開scriptのsplit定義を確認すべきです。","Unlabeledのignore設定、return番号の符号化、mesh・poseの座標系も結果に直結します。","また、本文と表で不一致がある箇所はTable 5・6の値を基準に記録します。"]);
}

const out=process.env.OUT_NAME || path.join(__dirname,"セミナー202060723.pptx");
pptx.writeFile({fileName:out});
