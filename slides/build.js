// AI 시대의 서·논술형 평가 — 연수 슬라이드 (일반 내용 버전)
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/3a682cee-e22e-4645-a187-2d23bdfad056_67af8d6b-9f22-43a6-9cd5-4b0e2008c511/pptx/scripts/apply_theme.js");
const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "연두 서논술형", headFontFace: "맑은 고딕", bodyFontFace: "맑은 고딕",
  colors: { dk1:"1F2A17", lt1:"FFFFFF", dk2:"35611C", lt2:"F1F8E4", accent1:"9CCC3D", accent2:"5F9A28", accent3:"D7EBAF",
    accent4:"E8792B", accent5:"FBE3D0", accent6:"6B7566", hlink:"35611C", folHlink:"5F9A28" }
};
const H = THEME.colors;
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "AI 활용 서·논술형 평가";
pres.author = "김수경";
const C = pres.SchemeColor;
const FOOT = "AI 활용 서·논술형 평가";
const SRC = "교육부·17개 시·도교육청·한국교육과정평가원 서·논술형 평가 연수 자료를 재구성";

/* ---------- 레이아웃 ---------- */
pres.defineSlideMaster({ title:"COVER", background:{ color:H.lt1 }, objects:[
  { placeholder:{ options:{ name:"title", type:"title", align:"left", x:0.6, y:1.45, w:5.3, h:1.75, fontSize:44, bold:true, color:C.text1, valign:"top", margin:0 }, text:"" } },
  { placeholder:{ options:{ name:"body", type:"body", x:0.6, y:3.4, w:5.3, h:0.4, fontSize:16, color:C.accent6, valign:"top", margin:0 }, text:"" } } ]});
pres.defineSlideMaster({ title:"SECTION", background:{ color:H.dk2 }, objects:[
  { placeholder:{ options:{ name:"title", type:"title", align:"left", x:3.3, y:1.5, w:6.1, h:1.1, fontSize:34, bold:true, color:C.background1, valign:"top", margin:0 }, text:"" } },
  { placeholder:{ options:{ name:"body", type:"body", x:3.3, y:2.75, w:6.1, h:2.0, fontSize:16, color:C.accent3, valign:"top", margin:0, paraSpaceAfter:8 }, text:"" } } ]});
const frame = (title, bg) => pres.defineSlideMaster({ title, background:{ color:bg }, objects:[
  { placeholder:{ options:{ name:"title", type:"title", align:"left", x:0.5, y:0.35, w:9.0, h:0.7, fontSize:28, bold:true, color:C.text1, valign:"middle", margin:0 }, text:"" } },
  { text:{ text:FOOT, options:{ x:0.5, y:5.2, w:5, h:0.25, fontSize:10, color:C.accent6, margin:0 } } } ],
  slideNumber:{ x:9.0, y:5.2, w:0.5, h:0.25, fontSize:10, color:H.accent6, align:"right" } });
frame("CONTENT", H.lt1);
frame("PRACTICE", H.lt2);

/* ---------- 도우미 ---------- */
let sec = "";
const section = t => { sec = t; pres.addSection({ title:t }); };
const slide = (master, title, notes) => { const s = pres.addSlide({ masterName:master, sectionTitle:sec }); s.addText(title, { placeholder:"title" }); if (notes) s.addNotes(notes); return s; };
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox:true, margin:0, valign:"top", fontSize:14, color:C.text1 }, o));
const box = (s, x, y, w, h, fill, name) => s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{ color:fill }, line:{ color:fill, width:0 }, rectRadius:0.08, objectName:name || "카드" });
const num = (s, n, x, y, d = 0.42, fill = C.accent2, color = C.background1) => s.addText(String(n), { shape:pres.ShapeType.ellipse, x, y, w:d, h:d, fill:{ color:fill }, line:{ color:fill, width:0 }, color, bold:true, fontSize:d > 0.5 ? 20 : 14, align:"center", valign:"middle", margin:0, objectName:"번호" });
const card = (s, x, y, w, h, head, body, o = {}) => {
  box(s, x, y, w, h, o.fill || C.background2);
  T(s, head, { x:x + 0.2, y:y + 0.18, w:w - 0.4, h:0.35, fontSize:o.hs || 16, bold:true, color:o.hc || C.text2 });
  if (body) T(s, body, { x:x + 0.2, y:y + 0.62, w:w - 0.4, h:h - 0.75, fontSize:14, color:o.bc || C.text1, paraSpaceAfter:4 });
};
const bullets = (s, items, o) => T(s, items.map((t, i) => ({ text:t, options:{ bullet:true, breakLine:i < items.length - 1 } })), Object.assign({ paraSpaceAfter:8 }, o));
const arrow = (s, x, y) => s.addShape(pres.ShapeType.rightArrow, { x, y, w:0.26, h:0.26, fill:{ color:C.accent1 }, line:{ color:C.accent1, width:0 }, objectName:"화살표" });
const source = s => T(s, SRC, { x:0.5, y:4.9, w:9, h:0.22, fontSize:10, color:C.accent6 });
const callout = (s, text, y = 4.25, o = {}) => { box(s, 0.5, y, 9, 0.55, o.fill || C.text2, "강조"); T(s, text, { x:0.75, y, w:8.5, h:0.55, fontSize:16, bold:true, color:o.color || C.background1, valign:"middle" }); };
const table = (s, rows, colW, y, o = {}) => s.addTable(rows.map((r, ri) => r.map((c, ci) => ({ text:c, options:{
    bold:ri === 0 || ci === 0, color:ri === 0 ? H.lt1 : H.dk1, fill:{ color:ri === 0 ? H.dk2 : ci === 0 ? H.accent3 : ri % 2 ? H.lt1 : H.lt2 },
    align:ri === 0 ? "center" : "left", valign:"middle", fontSize:14, margin:[4, 8, 4, 8] } }))),
  { x:0.5, y, w:colW.reduce((a, b) => a + b, 0), colW, border:{ type:"solid", color:H.accent3, pt:0.75 }, rowH:o.rowH, objectName:"표" });
const divider = (n, title, lines, notes) => { const s = pres.addSlide({ masterName:"SECTION", sectionTitle:sec });
  s.addText(String(n), { isTextBox:true, x:0.6, y:1.1, w:2.4, h:2.6, fontSize:150, bold:true, color:C.accent1, margin:0, valign:"middle", objectName:"번호" });
  s.addText(title, { placeholder:"title" });
  s.addText(lines.map((t, i) => ({ text:t, options:{ bullet:true, breakLine:i < lines.length - 1 } })), { placeholder:"body" });
  if (notes) s.addNotes(notes); };
const practice = (n, title, goal, steps, check, notes) => { const s = slide("PRACTICE", `실습 ${n}  ${title}`, notes);
  T(s, goal, { x:0.5, y:1.15, w:9, h:0.4, fontSize:16, color:C.text2, bold:true });
  steps.forEach(([h, b], i) => { const y = 1.75 + i * 0.62; num(s, i + 1, 0.5, y); T(s, h, { x:1.1, y, w:2.6, h:0.42, fontSize:16, bold:true, valign:"middle" }); T(s, b, { x:3.75, y, w:5.75, h:0.42, fontSize:14, valign:"middle" }); });
  box(s, 0.5, 4.3, 9, 0.6, C.background1, "점검 질문"); T(s, [{ text:"점검 질문  ", options:{ bold:true, color:C.accent4 } }, { text:check }], { x:0.75, y:4.3, w:8.5, h:0.6, fontSize:14, valign:"middle" }); };

/* ================= 도입 ================= */
section("도입");
{ const s = pres.addSlide({ masterName:"COVER", sectionTitle:sec });
  const E = (x, y, d, fill, tr) => s.addShape(pres.ShapeType.ellipse, { x, y, w:d, h:d, fill:{ color:fill, transparency:tr || 0 }, line:{ color:fill, width:0, transparency:100 }, objectName:"배경 원" });
  E(5.5, -1.0, 6.4, C.background2); E(6.9, 2.3, 4.6, C.accent3, 35); E(4.85, 4.35, 0.9, C.accent1, 15); E(9.0, 0.35, 0.5, C.accent1);
  // 답안지와 채점기준표를 닮은 그림
  s.addShape(pres.ShapeType.roundRect, { x:6.35, y:0.95, w:2.75, h:3.75, fill:{ color:C.background1 }, line:{ color:H.accent3, width:0.75 }, rectRadius:0.1, shadow:{ type:"outer", color:H.dk2, opacity:0.18, blur:10, offset:3, angle:90 }, objectName:"답안지" });
  const bar = (x, y, w, fill, h = 0.11) => s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{ color:fill }, line:{ color:fill, width:0 }, rectRadius:0.05, objectName:"줄" });
  bar(6.6, 1.25, 1.1, C.accent2, 0.16);
  [2.2, 2.0, 2.25, 1.6].forEach((w, i) => bar(6.6, 1.65 + i * 0.25, w, C.accent3));
  [[1.45, C.accent2], [1.15, C.accent1], [1.35, C.accent2]].forEach(([w, f], i) => { const y = 2.95 + i * 0.5;
    s.addText("✓", { shape:pres.ShapeType.ellipse, x:6.6, y, w:0.32, h:0.32, fill:{ color:f }, line:{ color:f, width:0 }, color:C.background1, bold:true, fontSize:12, align:"center", valign:"middle", margin:0, objectName:"점검" });
    bar(7.08, y + 0.1, w, C.accent3); });
  s.addShape(pres.ShapeType.roundRect, { x:5.45, y:3.25, w:1.75, h:1.15, fill:{ color:C.text2 }, line:{ color:C.text2, width:0 }, rectRadius:0.1, shadow:{ type:"outer", color:H.dk2, opacity:0.2, blur:8, offset:2, angle:90 }, objectName:"수준 카드" });
  [["상", 1.0, C.accent1], ["중", 0.7, C.accent3], ["하", 0.4, C.background1]].forEach(([t, w, f], i) => { const y = 3.4 + i * 0.3;
    T(s, t, { x:5.6, y, w:0.25, h:0.24, fontSize:11, bold:true, color:C.background1, valign:"middle" }); bar(5.9, y + 0.06, w, f, 0.12); });
  s.addText("2026 중등 AI·디지털 기반 역량강화 직무연수 [심화] 7기", { shape:pres.ShapeType.roundRect, rectRadius:0.15, x:0.6, y:0.65, w:4.75, h:0.36, fill:{ color:C.background2 }, line:{ color:C.background2, width:0 }, fontSize:12, bold:true, color:C.text2, align:"center", valign:"middle", margin:0, objectName:"연수명" });
  s.addText([{ text:"AI 활용", options:{ breakLine:true } }, { text:"서·논술형 평가", options:{ color:C.text2 } }], { placeholder:"title" });
  s.addText("평가 목적, 채점기준, 쓰기 과정과 피드백", { placeholder:"body" });
  T(s, "2026. 12. 3.(목) 17:00~19:50", { x:0.6, y:4.45, w:4.6, h:0.3, fontSize:14, color:C.accent6 });
  T(s, "수도여자고등학교 교사 김수경", { x:0.6, y:4.75, w:4.6, h:0.32, fontSize:16, bold:true, color:C.text1 }); }

{ const s = slide("CONTENT", "오늘의 흐름", "");
  [["1","무엇을 평가할까",["평가 목적 재정립","결과에서 과정으로","성취기준에서 과제로"]],
   ["2","어떻게 채점할까",["채점기준 개발","변별력 점검","타당성 점검"]],
   ["3","과정을 어떻게 볼까",["쓰기 과정 수업","AI 활용 허용 범위","AI 활용 피드백"]]].forEach(([k, h, items], i) => { const x = 0.5 + i * 3.1;
    box(s, x, 1.3, 2.8, 3.4, C.background2); num(s, k, x + 0.2, 1.5, 0.55);
    T(s, h, { x:x + 0.2, y:2.25, w:2.4, h:0.4, fontSize:19, bold:true, color:C.text2 }); bullets(s, items, { x:x + 0.2, y:2.9, w:2.4, h:1.6, fontSize:15 }); }); }

{ const s = slide("CONTENT", "AI가 글을 써 주는 교실", "");
  s.addImage({ path:__dirname + "/ebs.png", x:0.5, y:1.3, w:3.9, h:2.1, altText:"EBS 다큐프라임 생각의 멸종 1부 장면. 자막: AI와 함께 푼 수학 문제, 실력도 늘었을까?", objectName:"EBS 장면" });
  T(s, "출처: EBS 다큐프라임 〈생각의 멸종〉 1부", { x:0.5, y:3.46, w:3.9, h:0.22, fontSize:10, color:C.accent6 });
  box(s, 0.5, 3.8, 3.9, 1.0, C.text2); T(s, "우리는 지금 무엇을\n평가하고 있을까", { x:0.75, y:3.8, w:3.4, h:1.0, fontSize:20, bold:true, color:C.background1, valign:"middle" });
  [["AI와 함께하면 과제는 풀린다","수학 문제도, 글쓰기도 AI가 그럴듯한 답을 내놓습니다. 실력도 함께 늘었을까요?"],
   ["결과물만으로는 알 수 없다","잘 쓴 글이 학생의 생각인지, AI의 문장인지 구별하기 어렵습니다."],
   ["생각이 드러나는 평가가 필요하다","학생이 무엇을 이해하고 어떻게 판단했는지 볼 수 있어야 합니다."]].forEach(([h, b], i) => { const y = 1.3 + i * 1.2;
    num(s, i + 1, 4.75, y + 0.05); T(s, h, { x:5.35, y, w:4.15, h:0.4, fontSize:17, bold:true, color:C.text2 }); T(s, b, { x:5.35, y:y + 0.43, w:4.15, h:0.65 }); }); }

/* ================= 1차시 ================= */
section("1 평가 목적 재정립");
divider(1, "평가 목적 재정립", ["서·논술형 평가는 왜 하는가","결과 평가에서 과정 평가로","성취기준에서 과제와 채점기준으로"]);

{ const s = slide("CONTENT", "서·논술형 평가는 왜 하는가", "2022 개정 교육과정 총론: 수행평가를 내실화하고 서술형과 논술형 평가의 비중을 확대한다.");
  [["생각을 언어로 드러낸다","자신의 생각을 논리적인 글로 표현하며 스스로 돌아볼 기회를 얻습니다."],
   ["고차원적 사고를 요구한다","정답이 정해지지 않은 물음에 분석, 종합, 평가로 답해야 합니다."],
   ["과정을 종합적으로 본다","사고 과정, 논리적 표현, 문제 해결 방식을 함께 파악할 수 있습니다."],
   ["구체적인 피드백이 된다","학생이 쓴 글에서 이해 정도와 오개념을 읽어 낼 수 있습니다."]].forEach(([h, b], i) => card(s, 0.5 + (i % 2) * 4.65, 1.3 + Math.floor(i / 2) * 1.75, 4.35, 1.55, h, b));
  source(s); }

{ const s = slide("CONTENT", "결과 평가에서 과정 평가로", "표의 마지막 줄이 핵심. 과정 평가는 AI 대응책이기 전에 원래 서·논술형 평가가 가려던 방향입니다.");
  table(s, [["","결과 평가","과정 평가"],["평가 대상","완성된 글 한 편","글이 만들어지는 과정과 완성된 글"],["평가 시점","제출한 뒤 한 번","단계마다 여러 번"],
    ["평가 근거","최종 답안","계획서, 조사 기록, 초고, 고친 흔적, 최종 답안"],["피드백","점수를 알려 줌","쓰는 도중에 받고 고쳐 씀"],["AI의 영향","누가 썼는지 확인하기 어려움","생각의 흔적이 근거로 남음"]], [1.6, 3.2, 4.2], 1.3, { rowH:0.55 }); }

{ const s = slide("CONTENT", "과정이 곧 평가 근거가 된다", "단계마다 무엇이 남는지 미리 정해 두면, 최종 글만 보지 않아도 됩니다.");
  const st = [["질문과 계획","계획서"],["자료 조사","출처와\n해석 기록"],["초고 쓰기","초고"],["피드백과\n고쳐 쓰기","고친 점 기록"],["최종 글과\n성찰","자기 점검표"]];
  st.forEach(([h, e], i) => { const x = 0.5 + i * 1.84; box(s, x, 1.4, 1.58, 1.15, C.text2); T(s, h, { x, y:1.4, w:1.58, h:1.15, fontSize:15, bold:true, color:C.background1, align:"center", valign:"middle" });
    if (i < 4) arrow(s, x + 1.58, 1.84);
    box(s, x, 2.85, 1.58, 1.05, C.background2); T(s, e, { x, y:2.85, w:1.58, h:1.05, align:"center", valign:"middle" }); });
  T(s, "학생이 하는 일", { x:0.5, y:1.12, w:3, h:0.25, fontSize:11, bold:true, color:C.accent2 }); T(s, "남는 평가 근거", { x:0.5, y:2.6, w:3, h:0.25, fontSize:11, bold:true, color:C.accent2 });
  callout(s, "질문에서 최종 글까지 이어지는 기록이 학생의 생각을 보여 줍니다"); }

{ const s = slide("CONTENT", "AI에 열어 둘 일, 학생 몫으로 남길 일", "선생님들이 먼저 골라 보게 한 뒤 팁을 짚습니다.");
  T(s, "아래 활동 가운데 어떤 일을 AI에 열어 두고, 어떤 일을 학생 몫으로 남길까요?", { x:0.5, y:1.15, w:9, h:0.4, fontSize:16, bold:true, color:C.text2 });
  const pick = (t, x, y, w, col, hex) => s.addText(t, { shape:pres.ShapeType.roundRect, rectRadius:0.12, x, y, w, h:0.28, fill:{ color:C.background1 }, line:{ color:hex, width:0.75 }, fontSize:11, bold:true, color:col, align:"center", valign:"middle", margin:0, objectName:"선택" });
  ["검색어 만들기","자료의 해석","입장 정하기와 판단","어려운 용어 풀이","근거를 들어 논증하기","자료가 있는 기관 찾기","다 쓴 글의 논리 점검","자기 경험과 연결하기"].forEach((t, i) => { const x = 0.5 + (i % 2) * 4.65, y = 1.7 + Math.floor(i / 2) * 0.62;
    box(s, x, y, 4.35, 0.5, C.background2); T(s, t, { x:x + 0.2, y, w:2.6, h:0.5, fontSize:15, valign:"middle" });
    pick("AI", x + 2.95, y + 0.11, 0.55, C.accent2, H.accent2); pick("학생", x + 3.6, y + 0.11, 0.6, C.accent4, H.accent4); });
  callout(s, [{ text:"팁   ", options:{ color:C.accent1 } }, { text:"평가하려는 능력은 학생 몫으로 남깁니다" }], 4.3); }

{ const s = slide("CONTENT", "성취기준에서 채점기준까지 한 줄로", "다섯 칸이 같은 평가 요소로 이어지는지가 타당도의 출발점입니다.");
  [["성취기준","무엇을 배워야\n하는가"],["평가 요소","무엇을\n볼 것인가"],["과제","어떤 글을\n쓰게 할 것인가"],["채점기준","어떤 수준으로\n판단할 것인가"],["피드백","다음에 무엇을\n할 것인가"]].forEach(([h, q], i) => { const x = 0.5 + i * 1.84;
    box(s, x, 1.5, 1.58, 0.75, i === 1 ? C.accent4 : C.text2); T(s, h, { x, y:1.5, w:1.58, h:0.75, fontSize:17, bold:true, color:C.background1, align:"center", valign:"middle" });
    if (i < 4) arrow(s, x + 1.58, 1.74); T(s, q, { x, y:2.45, w:1.58, h:0.9, align:"center" }); });
  box(s, 0.5, 3.55, 9, 1.25, C.background2); T(s, [{ text:"평가 요소가 축입니다. ", options:{ bold:true, color:C.accent4 } }, { text:"과제의 발문과 조건, 채점기준의 채점 요소, 피드백의 관점이 모두 같은 평가 요소에서 나와야 가르친 것을 평가하게 됩니다." }], { x:0.75, y:3.55, w:8.5, h:1.25, fontSize:15, valign:"middle" }); }

{ const s = slide("CONTENT", "성취기준에서 평가 요소 뽑기", "성취기준의 모든 내용을 평가할 필요는 없습니다. 서·논술형으로 볼 핵심만 고릅니다.");
  [["지식·이해","알고 이해해야 할 내용"],["과정·기능","교과 고유의 사고와 탐구"],["가치·태도","교과 활동으로 기르는 가치와 태도"]].forEach(([h, b], i) => { const y = 1.3 + i * 1.15; box(s, 0.5, y, 3.6, 1.0, C.background2);
    T(s, h, { x:0.7, y:y + 0.12, w:3.2, h:0.35, fontSize:16, bold:true, color:C.text2 }); T(s, b, { x:0.7, y:y + 0.5, w:3.2, h:0.4 }); });
  T(s, "평가 요소를 정할 때", { x:4.5, y:1.3, w:5, h:0.35, fontSize:17, bold:true, color:C.text2 });
  bullets(s, ["모든 내용을 넣지 않고 핵심만 고릅니다","학생이 하는 행동이 드러나는 명사형 어구로 씁니다 (예: 쟁점 분석하기)","요소끼리 겹치거나 모호하지 않게 구분합니다","과제의 발문과 조건에 그대로 반영합니다"], { x:4.5, y:1.8, w:5, h:2.7 });
  source(s); }

{ const s = slide("CONTENT", "서·논술형 과제의 구조", "");
  [["발문","무엇을 어떻게 쓸지 알려 줍니다. 설명, 비교, 분석, 논하시오 같은 반응 지시어를 분명하게 씁니다."],
   ["자료","교사가 주는 제시문과 통계, 또는 학생이 직접 조사한 자료를 씁니다."],
   ["조건","내용, 형식, 분량 조건을 3~5개로 제한합니다. 정답을 암시하지 않게 씁니다."]].forEach(([h, b], i) => card(s, 0.5 + i * 2.3, 1.3, 2.1, 3.3, h, b));
  card(s, 7.4, 1.3, 2.1, 3.3, "과정 산출물", "계획서, 조사 기록, 초고처럼 단계마다 낼 것을 미리 정합니다.", { fill:C.accent5, hc:C.accent4 });
  source(s); }

{ const s = slide("CONTENT", "학생의 생각이 드러나는 과제", "AI를 못 쓰게 하는 과제가 아니라, AI를 써도 학생의 판단이 남는 과제입니다.");
  [["수업 맥락과 연결한다","수업 중 토론, 모둠 활동, 자기 경험을 근거로 쓰게 합니다."],
   ["직접 조사한 자료를 쓴다","학생이 찾은 사례와 출처, 그에 대한 해석을 요구합니다."],
   ["조건을 바꾸거나 반론하게 한다","상황이 달라지면 입장이 어떻게 바뀌는지, 반대 입장에 어떻게 답할지 묻습니다."],
   ["선택한 이유를 설명하게 한다","왜 이 근거를 골랐고 무엇을 버렸는지 쓰게 합니다."]].forEach(([h, b], i) => { const x = 0.5 + (i % 2) * 4.65, y = 1.3 + Math.floor(i / 2) * 1.75;
    box(s, x, y, 4.35, 1.55, C.background2); num(s, i + 1, x + 0.2, y + 0.2); T(s, h, { x:x + 0.8, y:y + 0.2, w:3.4, h:0.42, fontSize:16, bold:true, color:C.text2, valign:"middle" }); T(s, b, { x:x + 0.8, y:y + 0.72, w:3.4, h:0.75 }); }); }

/* ================= 2차시 ================= */
section("2 채점기준 개발과 점검");
divider(2, "채점기준 개발과 점검", ["채점기준의 구성과 유형","채점기준 작성 지침","변별력과 타당성 점검"]);

{ const s = slide("CONTENT", "채점기준은 세 단계로 만든다", "예시답안을 함께 쓰면 수준을 나누기가 쉬워집니다.");
  [["채점 요소 추출","평가 요소에서 채점할 요소를 가져옵니다. 한 과제에 여러 개일 수 있습니다."],
   ["수준 설정","요소마다 수준별 수행 특성을 씁니다. 성취기준 범위를 벗어나지 않게 합니다."],
   ["예시답안 작성","모범답안이 아니라 학생이 실제로 쓸 수준과 어휘로 씁니다. 유사답안도 함께 적습니다."]].forEach(([h, b], i) => { const x = 0.5 + i * 3.1;
    box(s, x, 1.3, 2.8, 3.3, C.background2); num(s, i + 1, x + 0.2, 1.5, 0.55); T(s, h, { x:x + 0.2, y:2.2, w:2.4, h:0.4, fontSize:18, bold:true, color:C.text2 }); T(s, b, { x:x + 0.2, y:2.7, w:2.4, h:1.8 });
    if (i < 2) arrow(s, x + 2.82, 2.8); });
  source(s); }

{ const s = slide("CONTENT", "총체적 채점과 분석적 채점", "");
  table(s, [["","총체적 채점","분석적 채점"],["방법","답안 전체를 보고 하나의 점수를 줌","채점 요소마다 점수를 주고 합산함"],["장점","채점이 빠름","구체적인 피드백이 가능하고 일관성을 지키기 쉬움"],
    ["단점","세부 피드백이 어렵고 채점자마다 달라지기 쉬움","개발과 채점에 시간이 듦"],["알맞은 경우","요소로 나누기 어려운 글","과정과 피드백이 중요한 평가"]], [1.6, 3.7, 3.7], 1.3, { rowH:0.62 });
  source(s); }

{ const s = slide("CONTENT", "총체적 채점기준 예시", "논술형 과제의 일반적인 예. 다음 장의 분석적 채점기준과 같은 과제입니다.");
  table(s, [["수준","채점기준"],["상 (3점)","입장이 분명하고 일관되며, 주장과 이어지는 근거를 둘 이상 들고, 자료를 출처와 함께 제시해 스스로 해석함"],
    ["중 (2점)","입장은 있으나 일부에서 흔들리고, 근거가 부족하거나 자료의 해석이 충분하지 않음"],
    ["하 (1점)","입장이 드러나지 않거나, 근거와 자료 없이 주장만 함"]], [1.6, 7.4], 1.25, { rowH:0.72 });
  T(s, "주장, 근거, 자료 활용을 따로 나누지 않고 답안 전체를 종합해 한 번에 판단합니다.", { x:0.5, y:4.35, w:9, h:0.4, color:C.accent6 }); }

{ const s = slide("CONTENT", "분석적 채점기준 예시", "같은 과제를 채점 요소별로 나눈 예. 마지막 요소가 AI 시대에 덧붙는 부분입니다.");
  table(s, [["채점 요소","상","중","하"],["주장","입장이 분명하고 글 전체에서 일관됨","입장은 있으나 일부에서 흔들림","입장이 드러나지 않음"],
    ["근거","주장과 이어지는 근거를 둘 이상 제시함","근거가 하나이거나 연결이 약함","근거 없이 주장만 함"],
    ["자료 활용","자료를 출처와 함께 제시하고 스스로 해석함","자료는 있으나 해석이 없음","자료를 쓰지 않음"],
    ["정직한 자료 활용","출처와 AI 활용을 모두 표시함","표시가 일부 빠짐","표시 없이 옮겨 씀"]], [1.65, 2.75, 2.45, 2.15], 1.25, { rowH:0.68 }); }

{ const s = slide("CONTENT", "채점기준 작성 지침", "가채점 뒤에 고치는 것까지가 채점기준 개발입니다.");
  [["구체적으로","누가 채점해도 같은 점수가 나오게 씁니다."],["겹치지 않게","수준 사이에 걸치는 부분이 없게 합니다."],["빠짐없이","어떤 답안이든 한 수준에 들어가게 합니다."],
   ["부분점수 인정","예시답안 밖의 답을 모두 0점 처리하지 않습니다."],["무관한 감점 금지","평가 목표가 아닌 필체, 맞춤법으로 깎지 않습니다."],["학생에게 공개","채점기준이 학습의 기준이 되게 합니다."]].forEach(([h, b], i) => card(s, 0.5 + (i % 3) * 3.1, 1.3 + Math.floor(i / 3) * 1.75, 2.8, 1.55, h, b));
  source(s); }

{ const s = slide("CONTENT", "변별력 점검: 수준을 구분하는가", "");
  box(s, 0.5, 1.3, 3.5, 3.5, C.text2); T(s, "가상 답안으로\n미리 채점해 보기", { x:0.75, y:1.6, w:3.0, h:1.0, fontSize:22, bold:true, color:C.background1 });
  T(s, "실제 답안을 받기 전에 채점기준의 빈틈을 찾습니다. 수준별 가상 답안은 AI에게 만들게 할 수 있습니다.", { x:0.75, y:2.85, w:3.0, h:1.7, color:C.accent3 });
  [["수준별 가상 답안 준비","상, 중, 하와 두 수준 사이에 걸친 답안"],["채점기준으로 직접 채점","의도한 수준은 가리고 채점"],["점수가 갈리는지 확인","애매한 답안을 기준으로 판정할 수 있는가"],["채점기준 보완","빠진 수준이나 모호한 표현 고치기"]].forEach(([h, b], i) => { const y = 1.3 + i * 0.9;
    num(s, i + 1, 4.35, y + 0.05); T(s, h, { x:4.95, y, w:4.55, h:0.38, fontSize:16, bold:true, color:C.text2 }); T(s, b, { x:4.95, y:y + 0.4, w:4.55, h:0.38 }); }); }

{ const s = slide("CONTENT", "타당성 점검: 보려던 것을 보는가", "네 번째 질문이 AI 시대에 새로 생긴 타당성 질문입니다.");
  [["채점 요소가 성취기준의 평가 요소와 일치하는가", false],["높은 점수가 글의 길이나 말솜씨가 아니라 평가 요소의 수행에서 나오는가", false],["수업에서 가르친 것을 평가하는가", false],["AI가 쓴 글이 그대로 높은 점수를 받는 구조는 아닌가", true]].forEach(([q, hot], i) => { const y = 1.3 + i * 0.88;
    box(s, 0.5, y, 9, 0.73, hot ? C.accent5 : C.background2); T(s, "Q" + (i + 1), { x:0.7, y, w:0.7, h:0.73, fontSize:18, bold:true, color:hot ? C.accent4 : C.accent2, valign:"middle" }); T(s, q, { x:1.45, y, w:7.85, h:0.73, fontSize:16, valign:"middle" }); }); }

{ const s = slide("CONTENT", "채점할 때 생기는 오류", "가채점으로 기준을 맞추고, 동료와 같은 답안을 채점해 보면 줄일 수 있습니다.");
  [["후광 효과","학생에 대한 선입견이 점수에 영향을 줌"],["집중화 경향","중간 점수로 몰려 변별이 안 됨"],["표준의 오류","채점자마다 기준이 다름"],
   ["대비의 오류","잘 아는 내용은 박하게, 낯선 내용은 후하게 줌"],["근접의 오류","처음이나 나중 답안의 인상에 치우침"],["논리의 오류","모순된 판단이 채점에 섞임"]].forEach(([h, b], i) => card(s, 0.5 + (i % 3) * 3.1, 1.3 + Math.floor(i / 3) * 1.75, 2.8, 1.55, h, b));
  source(s); }

/* ================= 3차시 ================= */
section("3 쓰기 과정과 피드백");
divider(3, "쓰기 과정과 피드백", ["쓰기 과정 수업의 흐름","AI 활용 허용 범위 안내","AI 활용 피드백"]);

{ const s = slide("CONTENT", "쓰기 과정 수업의 흐름", "4차시 예시. 채점기준을 첫 시간에 공개하는 것이 중요합니다.");
  [["1차시","과제와 채점기준 안내","개요 짜기","개요"],["2차시","초고 쓰기","수업 시간에 직접 쓰기","초고"],["3차시","피드백과 고쳐 쓰기","교사, 동료, AI 피드백","고친 점 기록"],["4차시","제출과 성찰","채점기준표로 자기 점검","최종 글, 점검표"]].forEach(([k, h, b, e], i) => { const x = 0.5 + i * 2.3;
    box(s, x, 1.3, 2.1, 2.3, C.background2); T(s, k, { x:x + 0.18, y:1.45, w:1.75, h:0.3, fontSize:13, bold:true, color:C.accent2 }); T(s, h, { x:x + 0.18, y:1.8, w:1.75, h:0.75, fontSize:16, bold:true, color:C.text2 }); T(s, b, { x:x + 0.18, y:2.65, w:1.75, h:0.8 });
    box(s, x, 3.75, 2.1, 0.85, C.accent3); T(s, [{ text:"남는 근거\n", options:{ fontSize:11, bold:true, color:C.text2 } }, { text:e }], { x:x + 0.18, y:3.75, w:1.75, h:0.85, valign:"middle" }); }); }

{ const s = slide("CONTENT", "단계별 AI 활용 허용 범위 (예시)", "학교와 과제에 맞게 고쳐 쓰는 예시입니다. 기준은 평가하려는 능력을 학생이 수행하는가입니다.");
  table(s, [["쓰기 단계","AI 사용","조건"],["주제와 질문 정하기","사용 안 함","스스로 정하기"],["자료 조사","사용 가능","AI 활용 표시, 수치는 원래 출처에서 확인"],["해석과 입장 정하기","사용 안 함","내 말로 쓰기"],["초고 쓰기","사용 안 함","수업 시간에 쓰기"],["고쳐 쓰기","조건부 사용","논리가 끊긴 곳을 지적받는 데만 쓰기"]], [2.6, 1.9, 4.5], 1.3, { rowH:0.55 }); }

{ const s = slide("CONTENT", "학생용 AI 활용 안내문 (예시)", "안내문의 규칙이 채점기준에 들어 있어야 학생이 지킵니다.");
  box(s, 0.5, 1.3, 4.9, 3.5, C.background2); T(s, "이번 서·논술형 평가의 AI 활용 약속", { x:0.75, y:1.48, w:4.4, h:0.35, fontSize:16, bold:true, color:C.text2 });
  bullets(s, ["자료 조사에는 AI를 써도 됩니다","AI가 알려 준 수치는 원래 출처에서 확인합니다","AI를 쓴 곳에는 무엇을 물었는지 표시합니다","해석과 주장은 내 말로 씁니다","제출 전에 채점기준표로 스스로 점검합니다"], { x:0.75, y:1.98, w:4.4, h:2.7, paraSpaceAfter:6 });
  T(s, "안내문에 담을 것", { x:5.8, y:1.3, w:3.7, h:0.35, fontSize:17, bold:true, color:C.text2 });
  [["쓸 수 있는 단계와 없는 단계","학생이 헷갈리지 않게"],["AI를 썼을 때의 표시 방법","어디에, 무엇을 물었는지"],["채점기준과의 연결","정직한 자료 활용을 채점 요소로"]].forEach(([h, b], i) => { const y = 1.85 + i * 0.98; num(s, i + 1, 5.8, y + 0.03); T(s, h, { x:6.4, y, w:3.1, h:0.38, fontSize:15, bold:true }); T(s, b, { x:6.4, y:y + 0.4, w:3.1, h:0.38, color:C.accent6 }); }); }

{ const s = slide("CONTENT", "좋은 피드백의 요소", "피드백은 점수 통지가 아니라 다음에 무엇을 할지 알려 주는 것입니다.");
  [["명확성","학생이 이해할 수 있는 말로"],["긍정적 태도","노력과 잘한 점을 먼저 인정"],["목표 참조","목표와 현재 위치, 남은 거리를 제시"],["과정 중심","정답 여부보다 사고 과정과 근거 제시"],
   ["자기조절 유도","스스로 돌아보고 고치게"],["구체성","잘했어요 대신 무엇을 어떻게"],["개인화","학생의 수준과 오답 유형에 맞게"]].forEach(([h, b], i) => { const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4, x = 0.5 + col * 4.65, y = 1.3 + row * 0.85;
    box(s, x, y, 4.35, 0.72, C.background2); T(s, h, { x:x + 0.2, y, w:1.45, h:0.72, fontSize:15, bold:true, color:C.text2, valign:"middle" }); T(s, b, { x:x + 1.65, y, w:2.6, h:0.72, valign:"middle" }); });
  box(s, 5.15, 3.85, 4.35, 0.72, C.text2); T(s, "고쳐 쓸 기회가 있어야 피드백이 됩니다", { x:5.35, y:3.85, w:3.95, h:0.72, fontSize:15, bold:true, color:C.background1, valign:"middle" });
  source(s); }

{ const s = slide("CONTENT", "AI 활용 피드백: 초안은 AI, 확정은 교사", "AI에게 점수를 매기게 하지 않습니다. 채점기준에 비춘 피드백 초안까지만 맡깁니다.");
  [["채점기준과\n익명 답안","C"],["AI가 쓴\n피드백 초안","C"],["교사가 점검하고\n고쳐 확정","A"],["학생이 읽고\n고쳐 쓰기","C"]].forEach(([h, k], i) => { const x = 0.5 + i * 2.3; box(s, x, 1.3, 2.02, 0.95, k === "A" ? C.accent4 : C.text2);
    T(s, h, { x, y:1.3, w:2.02, h:0.95, fontSize:15, bold:true, color:C.background1, align:"center", valign:"middle" }); if (i < 3) arrow(s, x + 2.03, 1.65); });
  card(s, 0.5, 2.5, 4.35, 2.3, "교사가 점검할 것", null); bullets(s, ["답안에 없는 내용을 말하지 않는가","근거 없는 칭찬이 지나치지 않은가","대신 고쳐 써 주지 않는가","학생의 입장을 한쪽으로 몰아가지 않는가"], { x:0.7, y:3.1, w:3.95, h:1.65, paraSpaceAfter:4 });
  card(s, 5.15, 2.5, 4.35, 2.3, "지킬 것", null, { fill:C.accent5, hc:C.accent4 }); bullets(s, ["채점과 점수는 교사가 합니다","이름, 학번 같은 개인정보는 넣지 않습니다","학교와 교육청이 허용한 도구를 씁니다"], { x:5.35, y:3.1, w:3.95, h:1.65, paraSpaceAfter:4 }); }

{ const s = pres.addSlide({ masterName:"SECTION", sectionTitle:sec });
  T(s, "정리", { x:0.6, y:0.55, w:6, h:0.35, fontSize:14, bold:true, color:C.accent1 });
  T(s, "AI 활용 서·논술형 평가", { x:0.6, y:0.95, w:8.8, h:0.7, fontSize:32, bold:true, color:C.background1 });
  [["과정을 평가한다","완성된 글과 함께 글이 만들어지는 과정을 근거로 삼습니다"],["평가 요소로 잇는다","성취기준, 과제, 채점기준, 피드백이 한 줄로 이어집니다"],["채점기준을 점검한다","수준을 구분하는지, 보려던 것을 보는지 확인합니다"],["피드백으로 돌려준다","고쳐 쓸 기회와 함께 다음에 할 일을 알려 줍니다"]].forEach(([h, b], i) => { const x = 0.6 + (i % 2) * 4.5, y = 2.0 + Math.floor(i / 2) * 1.3;
    num(s, i + 1, x, y + 0.04, 0.42, C.accent1, C.text1); T(s, h, { x:x + 0.6, y, w:3.7, h:0.38, fontSize:18, bold:true, color:C.background1 }); T(s, b, { x:x + 0.6, y:y + 0.42, w:3.7, h:0.65, color:C.accent3 }); }); }

(async () => { await pres.writeFile({ fileName:OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
