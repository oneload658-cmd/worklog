const pptxgen = require("pptxgenjs");
const { execFileSync } = require("child_process");
const { applyTheme } = require("/root/.claude/skills/synced/06c47b33-d640-428c-a3a5-a42ae98094c8_2ae10ec8-84d2-43dd-8c58-6c8a9c68ff68/pptx/scripts/apply_theme.js");
const A = "/home/user/worklog/act-match/assets/";
const OUT = "/home/user/worklog/act-match/ACT_MATCH_GLOBAL_재설계안_v4_WEUM.pptx";

// ── text measurement (Pretendard metrics) for highlighter bands
const FONTS = { xb: "/root/.fonts/Pretendard-ExtraBold.otf", bk: "/root/.fonts/Pretendard-Black.otf" };
function measure(text, pt, w = "xb") {
  const out = execFileSync("python3", ["-c",
    "import sys;from PIL import ImageFont;f=ImageFont.truetype(sys.argv[1],int(float(sys.argv[2])*20));print(f.getlength(sys.argv[3])/20/72)",
    FONTS[w], String(pt), text]).toString();
  return parseFloat(out);
}

const THEME = {
  name: "ACT MATCH Script",
  headFontFace: "Pretendard ExtraBold",
  bodyFontFace: "Pretendard",
  colors: {
    dk1: "141414", lt1: "FFFFFF", dk2: "2B2B2B", lt2: "F3F3F3",
    accent1: "FFE14D", accent2: "D92B2B", accent3: "7A7A7A", accent4: "DADADA",
    accent5: "1F1F1F", accent6: "FFF6C2", hlink: "D92B2B", folHlink: "7A7A7A",
  },
};
const HEX = THEME.colors;
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "ACT MATCH GLOBAL 재설계안 v4.0";
pres.author = "WEUM CORPORATION";
pres.company = "WEUM CORPORATION";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6, CW = W - 2 * M;
const MONO = "Courier New";
const FOOT = "ACT MATCH GLOBAL 재설계안 v4.0  ·  내부 검토용  ·  CONFIDENTIAL  ·  © 2026 WEUM CORPORATION";
const TITLE_Y = 0.82, TITLE_PT = 30;

function layout(name, dark) {
  pres.defineSlideMaster({
    title: name,
    background: { color: dark ? C.text1 : C.background1 },
    objects: [
      { image: { x: M, y: 6.95, w: 0.24, h: 0.24, path: A + (dark ? "mark_dark.png" : "mark_light.png") } },
      { text: { text: FOOT, options: { x: M + 0.34, y: 6.95, w: 9, h: 0.24, fontSize: 9, color: C.accent3, margin: 0, valign: "middle" } } },
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.4, w: 10.5, h: 0.3, fontSize: 12, fontFace: MONO, color: C.accent3, margin: 0, bold: true }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: TITLE_Y, w: CW, h: 0.6, fontSize: TITLE_PT, color: dark ? C.background1 : C.text1, margin: 0, valign: "top", align: "left", fontFace: "Pretendard ExtraBold" }, text: "" } },
    ],
    slideNumber: { x: W - M - 0.6, y: 6.95, w: 0.6, h: 0.24, fontSize: 9, color: C.accent3, align: "right", fontFace: MONO },
  });
}
layout("LIGHT", false);
layout("DARK", true);
pres.defineSlideMaster({ title: "COVER", background: { color: C.text1 }, objects: [] });

let sceneNo = 1;
// header: screenplay slug + title with highlighter band behind `hl`
function head(s, slug, title, hl, kind, dark) {
  sceneNo++;
  s.addText(`S#${String(sceneNo).padStart(2, "0")}.  ${slug}`, { placeholder: "kicker" });
  if (hl) {
    const i = title.indexOf(hl);
    if (i < 0) throw new Error("hl not in title: " + title);
    const full = measure(title, TITLE_PT);
    if (full > CW - 0.9) throw new Error("title too long: " + title + " " + full);
    const x0 = M + measure(title.slice(0, i), TITLE_PT);
    const w = measure(hl, TITLE_PT);
    s.addShape(pres.ShapeType.rect, { x: x0 - 0.04, y: TITLE_Y + 0.27, w: w + 0.08, h: 0.24, fill: { color: dark ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "highlighter" });
  }
  s.addText(title, { placeholder: "title" });
  if (kind) tag(s, W - M - 0.62, 0.41, kind);
}
// 사실 = ink stamp, 추론 = highlighter stamp
function tag(s, x, y, kind) {
  const fact = kind === "사실";
  s.addText(kind, {
    x, y, w: 0.62, h: 0.28, shape: pres.ShapeType.roundRect, rectRadius: 0.05, isTextBox: true,
    fill: { color: fact ? C.text1 : C.accent1 }, color: fact ? C.background1 : C.text1,
    line: { color: HEX.dk1, width: 0.75 }, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0,
  });
}
function card(s, x, y, w, h, fill, line) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.06, fill: { color: fill || C.background2 }, line: line ? { color: HEX.dk1, width: 1 } : { type: "none" } });
}
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 14 }, o)); }
function num(s, n, x, y, d, fill, color) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill || C.text1 }, line: { type: "none" } });
  T(s, String(n), { x, y, w: d, h: d, fontSize: d * 30, bold: true, color: color || C.background1, align: "center", valign: "middle" });
}
const L = (master, sec) => pres.addSlide({ masterName: master, sectionTitle: sec });
const bullets = (arr, o) => arr.map((t, k) => ({ text: t, options: Object.assign({ bullet: true, breakLine: k < arr.length - 1 }, o || {}) }));

// ════════════════════════════════════════ v4 BODY
const sec = (t) => pres.addSection({ title: t });
function rowsList(s, items, x, y, w, step, o = {}) {
  items.forEach((it, i) => {
    const yy = y + i * step;
    if (o.num !== false) num(s, i + 1, x, yy + 0.02, 0.4, o.fill || C.text1, o.color || C.background1);
    T(s, it[0], { x: x + 0.6, y: yy, w: w - 0.6, h: 0.35, fontSize: o.h || 15, bold: true, color: o.dark ? C.background1 : C.text1 });
    T(s, it[1], { x: x + 0.6, y: yy + 0.37, w: w - 0.6, h: step - 0.42, fontSize: o.b || 12.5, color: o.dark ? C.accent4 : C.accent3, lineSpacingMultiple: 1.2 });
  });
}

// ═════════ 1. COVER
sec("표지");
{
  const s = L("COVER", "표지");
  s.addImage({ path: A + "lockup_dark.png", x: M, y: 0.55, w: 3.75, h: 1.0 });
  T(s, "S#01.  EXT. 인천공항 출국장 — DAWN", { x: M, y: 2.15, w: 9, h: 0.35, fontSize: 15, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "한국 배우를,", { x: M, y: 2.65, w: 11, h: 1.0, fontSize: 60, fontFace: "Pretendard Black", color: C.background1 });
  const w2 = measure("세계의 장면으로.", 60, "bk");
  s.addShape(pres.ShapeType.rect, { x: M - 0.08, y: 3.73, w: w2 + 0.2, h: 0.98, fill: { color: C.accent1 }, line: { type: "none" } });
  T(s, "세계의 장면으로.", { x: M, y: 3.67, w: 11, h: 1.0, fontSize: 60, fontFace: "Pretendard Black", color: C.text1 });
  T(s, "ACT MATCH 냉정한 사업성 재검토 · 플필 해부 · 글로벌 재설계안 v4.0", { x: M, y: 5.1, w: 11.5, h: 0.4, fontSize: 18, color: C.accent4 });
  T(s, "2026. 11  ·  WEUM CORPORATION  ·  내부 검토용  ·  CONFIDENTIAL", { x: M, y: 6.8, w: 9, h: 0.3, fontSize: 11, fontFace: MONO, color: C.accent3 });
}

// ═════════ 2. VERDICT
sec("냉정한 판정");
{
  const s = L("LIGHT", "냉정한 판정");
  head(s, "INT. 투자심사실 — DAY", "결론부터: v3 그대로라면 투자받기 어렵다", "투자받기 어렵다", "추론");
  s.addChart(pres.ChartType.bar, [
    { name: "v3 (국내 매칭 앱)", labels: ["실현 가능성", "사업성", "투자 매력", "글로벌 확장성"], values: [45, 30, 25, 40] },
    { name: "v4 (글로벌 재설계)", labels: ["실현 가능성", "사업성", "투자 매력", "글로벌 확장성"], values: [75, 65, 60, 75] },
  ], {
    x: M, y: 1.7, w: 7.0, h: 4.95, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 55, catAxisOrientation: "maxMin",
    chartColors: [HEX.accent4, HEX.dk1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
    showTitle: true, title: "100점 만점 평가 (추론)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valAxisMaxVal: 100, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: true, legendPos: "t", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.accent3,
  });
  const X = 7.95, w = CW - 7.35;
  card(s, X, 1.75, w, 4.9, C.text1);
  T(s, "한 줄 판정", { x: X + 0.35, y: 2.0, w: w - 0.7, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "‘또 하나의 캐스팅 앱’은\n이미 늦었다.\n‘한국 배우 수출 엔진’은\n아직 비어 있다.", { x: X + 0.35, y: 2.45, w: w - 0.7, h: 2.2, fontSize: 21, fontFace: "Pretendard ExtraBold", color: C.background1, lineSpacingMultiple: 1.2 });
  T(s, bullets(["국내 매칭은 자금 받은 경쟁사 3곳 이상이 선점", "WEUM의 무기는 기술이 아니라 사람·콘텐츠·광고", "그 무기는 ‘해외 수출’에서 가장 비싸게 팔린다"]), { x: X + 0.35, y: 4.8, w: w - 0.7, h: 1.7, fontSize: 13, color: C.accent4, paraSpaceAfter: 6 });
}

// ═════════ 3. FATAL FLAWS
{
  const s = L("LIGHT", "냉정한 판정");
  head(s, "INT. 투자심사실 — CONTINUOUS", "v3가 넘지 못하는 다섯 개의 벽", "다섯 개의 벽", "사실");
  const F = [
    ["선점된 시장", "플필(배우 약 2만·기업 2천여, Pre-A 투자), CASTLINK(프로필 3.4만+, LA 법인), 캐스틱(미국 진출 목표)이 이미 같은 일을 한다"],
    ["작은 시장 상한", "업계 선두권 CASTLINK의 2025년 매출이 약 16.7억 원 (2024년 5.8억, 적자). 국내 캐스팅 플랫폼만으로는 큰 매출이 어렵다"],
    ["기능 차별의 소멸", "‘같은 대본 연기 챌린지’조차 CASTLINK가 전 세계 대상 8주 챌린지로 예고. 기능은 몇 주면 복제된다"],
    ["냉시동 문제", "우리에게는 배우 DB·연출자 네트워크·데이터가 0. 양면 플랫폼은 공급과 수요를 동시에 0에서 시작해야 한다"],
    ["차가운 투자 환경", "‘콘텐츠 투자금 급감’ 보도, 벤처투자 절반 이상이 AI로 쏠림, 초기 단계 투자 감소. 트랙션 없는 후발 플랫폼은 심사 대상이 아니다"],
  ];
  F.forEach((f, i) => {
    const y = 1.75 + i * 0.98;
    card(s, M, y, CW, 0.86, i % 2 ? C.background1 : C.background2);
    T(s, "0" + (i + 1), { x: M + 0.3, y, w: 0.7, h: 0.86, fontSize: 22, fontFace: "Pretendard Black", color: C.accent2, valign: "middle" });
    T(s, f[0], { x: M + 1.05, y, w: 2.6, h: 0.86, fontSize: 16, bold: true, valign: "middle" });
    T(s, f[1], { x: M + 3.7, y, w: CW - 4.0, h: 0.86, fontSize: 13, color: C.accent3, valign: "middle", lineSpacingMultiple: 1.2 });
  });
  T(s, "※ 수치·사례는 사실, 각 항목의 결론 문장은 추론", { x: M, y: 6.58, w: CW, h: 0.28, fontSize: 10.5, italic: true, color: C.accent3 });
}

// ═════════ 4. UNFAIR ADVANTAGES
{
  const s = L("LIGHT", "냉정한 판정");
  head(s, "INT. WEUM 본사 — NIGHT", "그래도 WEUM에만 있는 무기는 네 가지다", "네 가지", "추론");
  const U = [
    ["LEGEND", "레전드 배우", "나문희 배우. 플랫폼은 복제되지만 사람은 복제되지 않는다. 콘텐츠·신뢰·화제성을 한 번에 만든다."],
    ["LICENSE", "매니지먼트 기반", "배우 엔터 대표 회사. 출연 계약·정산·보호를 실제로 해본 조직은 플랫폼 스타트업에 없다."],
    ["BRAND $", "10년 광고·기획", "브랜드 예산을 끌어오는 능력. 캐스팅 시장에서 가장 큰 돈은 광고에 있다."],
    ["STORY", "콘텐츠 기획", "‘발견 → 작품’을 보여주는 스토리텔링. 배우 발굴을 볼거리로 만드는 힘."],
  ];
  const cw = (CW - 3 * 0.3) / 4;
  U.forEach((u, i) => {
    const X = M + i * (cw + 0.3), dark = i === 0;
    card(s, X, 1.75, cw, 3.9, dark ? C.text1 : C.background2);
    T(s, u[0], { x: X + 0.3, y: 2.0, w: cw - 0.6, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: dark ? C.accent1 : C.accent3 });
    T(s, u[1], { x: X + 0.3, y: 2.45, w: cw - 0.6, h: 0.5, fontSize: 20, fontFace: "Pretendard Black", color: dark ? C.background1 : C.text1 });
    T(s, u[2], { x: X + 0.3, y: 3.15, w: cw - 0.6, h: 2.3, fontSize: 13.5, color: dark ? C.accent4 : C.accent3, lineSpacingMultiple: 1.35 });
  });
  card(s, M, 5.9, CW, 0.75, C.accent1);
  T(s, "결론: 기술로 싸우지 말고, 사람·콘텐츠·광고 예산으로 싸운다. 그 전장은 국내가 아니라 ‘국경’이다.", { x: M + 0.35, y: 5.9, w: CW - 0.7, h: 0.75, fontSize: 16, bold: true, valign: "middle" });
}

// ═════════ 5. PLFIL PROFILE
sec("플필 해부");
{
  const s = L("LIGHT", "플필 해부");
  head(s, "INT. 플필 사무실 — DAY", "플필 해부 ①  회사는 작고, 사업은 넓다", "회사는 작고, 사업은 넓다", "사실");
  const kpi = [["2만 명", "신인 연예인 회원"], ["2천여 곳", "이용 기업"], ["8명", "임직원 (2026.6)"], ["2022", "법인 설립"]];
  const kw = (CW - 3 * 0.25) / 4;
  kpi.forEach((k, i) => {
    const X = M + i * (kw + 0.25);
    card(s, X, 1.75, kw, 1.35, i === 2 ? C.accent6 : C.background2);
    T(s, k[0], { x: X + 0.3, y: 1.9, w: kw - 0.6, h: 0.6, fontSize: 28, fontFace: "Pretendard Black" });
    T(s, k[1], { x: X + 0.3, y: 2.55, w: kw - 0.6, h: 0.35, fontSize: 13, color: C.accent3 });
  });
  const biz = [
    ["캐스팅 플랫폼", "배우 DB 키워드 검색(사투리·외국어 등), 캐스팅 공고, 프로필 공개 범위 설정"],
    ["플필 플러스 (배우 유료)", "월 29,800원. 현직자 프로필·연기영상 피드백 70% 할인, 공고 주 1회 문자. ‘캐스팅 연락 약 2배’(회사 주장)"],
    ["에이전트 플필 (2024~)", "미국형 에이전트 기반 전속 매니지먼트. 배우·스포츠 선수 ‘글로벌 매니지먼트’ 표방"],
    ["스튜디오 플필", "자체 숏폼 드라마. 쇼츠 2편으로 조회수 1,000만 돌파, 2025.2 기준 5편 제작 완료"],
    ["초상권 IP 중개", "컴투스 게임 캐릭터에 배우 얼굴 스캔·음성 캐스팅 중개. ‘AI 시대 휴먼 IP 인프라’ 비전"],
    ["투자 · 해외", "프라이머(2024.1), 땡스벤처스·킹슬리벤처스(2025 말). 해외 인재 이메일 가입 개방"],
  ];
  const hw = (CW - 0.3) / 2;
  biz.forEach((b, i) => {
    const X = M + (i % 2) * (hw + 0.3), y = 3.35 + Math.floor(i / 2) * 1.1;
    T(s, b[0], { x: X, y, w: hw, h: 0.33, fontSize: 14.5, bold: true });
    T(s, b[1], { x: X, y: y + 0.36, w: hw, h: 0.65, fontSize: 12, color: C.accent3, lineSpacingMultiple: 1.2 });
  });
}

// ═════════ 6. PLFIL MONEY FLOW
{
  const s = L("LIGHT", "플필 해부");
  head(s, "INT. 플필 사무실 — NIGHT", "플필 해부 ②  돈은 누가 내는가", "누가 내는가", "추론");
  // flow diagram: actors -> platform <- companies ; platform -> agent/studio/IP
  const box = (x, y, w, h, t, sub, fill, dark) => {
    card(s, x, y, w, h, fill, !dark && fill === C.background1);
    T(s, t, { x: x + 0.2, y: y + 0.18, w: w - 0.4, h: 0.4, fontSize: 16, bold: true, align: "center", color: dark ? C.background1 : C.text1 });
    T(s, sub, { x: x + 0.2, y: y + 0.6, w: w - 0.4, h: h - 0.7, fontSize: 12, align: "center", color: dark ? C.accent4 : C.accent3, lineSpacingMultiple: 1.2 });
  };
  box(M, 2.0, 3.2, 1.6, "배우 2만 명", "플필 플러스 월 29,800원\n피드백·프로필 투어 대행", C.accent1, false);
  box(M + 4.45, 2.0, 3.2, 1.6, "플필 플랫폼", "DB 검색 · 공고 보드\n(8명 조직)", C.text1, true);
  box(M + 8.9, 2.0, 3.23, 1.6, "기업 2천여 곳", "공고 게시·배우 검색\n(기업 과금 구조 미공개)", C.background1, false);
  s.addShape(pres.ShapeType.line, { x: M + 3.25, y: 2.8, w: 1.15, h: 0, line: { color: HEX.dk1, width: 2, endArrowType: "triangle" } });
  s.addShape(pres.ShapeType.line, { x: M + 7.7, y: 2.8, w: 1.15, h: 0, line: { color: HEX.dk1, width: 2, beginArrowType: "triangle" } });
  T(s, "₩ 구독", { x: M + 3.25, y: 2.4, w: 1.15, h: 0.3, fontSize: 11, bold: true, align: "center", color: C.accent2 });
  T(s, "₩ ?", { x: M + 7.7, y: 2.4, w: 1.15, h: 0.3, fontSize: 11, bold: true, align: "center", color: C.accent3 });
  const sub = [["에이전트 플필", "전속 배우 수수료"], ["스튜디오 플필", "숏폼 IP"], ["초상권 중개", "게임·광고 IP"]];
  sub.forEach((x, i) => {
    const X = M + 2.3 + i * 2.75;
    s.addShape(pres.ShapeType.line, { x: M + 6.05, y: 3.65, w: X + 1.15 - (M + 6.05), h: 0.55, line: { color: HEX.accent3, width: 1, dashType: "dash" } });
    card(s, X, 4.25, 2.3, 0.95, C.background2);
    T(s, x[0], { x: X, y: 4.35, w: 2.3, h: 0.35, fontSize: 13.5, bold: true, align: "center" });
    T(s, x[1], { x: X, y: 4.72, w: 2.3, h: 0.3, fontSize: 11.5, align: "center", color: C.accent3 });
  });
  card(s, M, 5.5, CW, 1.15, C.text1);
  T(s, [
    { text: "해석  ", options: { bold: true, color: C.accent1 } },
    { text: "확인되는 반복 매출은 ‘배우가 내는 구독료’다. 수익원을 매니지먼트·제작·IP로 넓히는 중이며, 이는 플랫폼·소속사·제작사가 한 몸이 되는 구조다.", options: { color: C.background1 } },
  ], { x: M + 0.35, y: 5.5, w: CW - 0.7, h: 1.15, fontSize: 14.5, valign: "middle", lineSpacingMultiple: 1.25 });
}

// ═════════ 7. STRENGTHS TO ABSORB
{
  const s = L("LIGHT", "플필 해부");
  head(s, "INT. 플필 사무실 — DAY", "플필 해부 ③  배워서 가져올 장점 6가지", "가져올 장점", "추론");
  const P = [
    ["키워드 검색", "사투리·외국어 같은 ‘특기 태그’ 검색", "→ 해외용 태그: 영어·일본어 레벨, 무술, 비자 상태"],
    ["피드백 서비스", "배우의 성장 욕구를 돈으로 바꾼 상품", "→ 해외 오디션 코칭·셀프테이프 피드백"],
    ["미국형 에이전트", "플랫폼 + 전속 에이전트의 결합", "→ WEUM은 이미 매니지먼트, 해외 협력 에이전트 추가"],
    ["자체 숏폼 제작", "발굴 → 출연까지 직접 보여줌", "→ LEGEND SCRIPT로 같은 증명, 더 큰 화제"],
    ["초상권 동의 중개", "AI 시대 얼굴·목소리 권리 사업", "→ 해외 광고·게임의 한국 배우 IP 라이선싱"],
    ["해외 가입 개방", "이메일만으로 외국 인재 가입", "→ 반대 방향(한국 배우 → 해외)을 우리가 연다"],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  P.forEach((p, i) => {
    const X = M + (i % 3) * (cw + 0.3), y = 1.75 + Math.floor(i / 3) * 2.5;
    card(s, X, y, cw, 2.3);
    num(s, i + 1, X + 0.3, y + 0.28, 0.42, C.accent1, C.text1);
    T(s, p[0], { x: X + 0.85, y: y + 0.3, w: cw - 1.1, h: 0.4, fontSize: 16.5, bold: true });
    T(s, p[1], { x: X + 0.3, y: y + 0.9, w: cw - 0.6, h: 0.5, fontSize: 13, color: C.accent3 });
    T(s, p[2], { x: X + 0.3, y: y + 1.45, w: cw - 0.6, h: 0.7, fontSize: 13, bold: true, color: C.accent2, lineSpacingMultiple: 1.2 });
  });
}

// ═════════ 8. WEAKNESSES TO FIX
{
  const s = L("LIGHT", "플필 해부");
  head(s, "INT. 편집실 — NIGHT", "플필 해부 ④  우리가 메울 빈틈 6가지", "메울 빈틈", "추론");
  const W6 = [
    ["‘더 내면 더 보인다’", "배우 구독 + ‘연락 2배’ 구조는 저소득 배우에게 역진적이고 공정성 의심을 부른다", "배우 기본 무료, 돈은 제작·브랜드가 낸다"],
    ["한 몸이 된 이해관계", "플랫폼·전속 에이전트·제작사가 한 회사. 노출 공정성 규칙 공개는 확인되지 않음", "방화벽 4규칙 공개 + 외부 점검"],
    ["해외 = 인바운드", "해외 전략이 ‘외국 인재 → K엔터’ 가입 개방. 한국 배우의 해외 진출은 비어 있다", "아웃바운드 ‘수출’ 전담"],
    ["한국식 프로필", "한국은 모델형 사진, 미국은 캐릭터형 헤드샷(AP 보도). 해외용 포맷이 따로 필요", "ACT PASSPORT 현지 포맷"],
    ["8명 조직의 한계", "셀프테이프 품질·통번역·비자·현지 계약 같은 오프라인 실무는 소규모 IT 조직에 무겁다", "스튜디오 + 글로벌 데스크"],
    ["연출자 정보 비대칭", "배우가 연출자의 현장 평판을 볼 방법이 없다", "SET REVIEW + VERIFIED"],
  ];
  W6.forEach((w, i) => {
    const y = 1.72 + i * 0.83;
    T(s, w[0], { x: M, y, w: 2.9, h: 0.72, fontSize: 14.5, bold: true, valign: "middle" });
    T(s, w[1], { x: M + 3.0, y, w: 6.0, h: 0.72, fontSize: 12.5, color: C.accent3, valign: "middle", lineSpacingMultiple: 1.15 });
    s.addText("→ " + w[2], { x: M + 9.2, y: y + 0.1, w: CW - 9.2, h: 0.52, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: C.accent6 }, color: C.text1, fontSize: 12.5, bold: true, valign: "middle", margin: [0, 8, 0, 8] });
    if (i < 5) s.addShape(pres.ShapeType.line, { x: M, y: y + 0.78, w: CW, h: 0, line: { color: HEX.accent4, width: 0.75 } });
  });
}

// ═════════ 9. COMPETITIVE MAP
{
  const s = L("LIGHT", "플필 해부");
  head(s, "EXT. 경쟁 지도 — DAY", "아무도 서 있지 않은 칸은 ‘해외 × 사람’이다", "‘해외 × 사람’", "추론");
  const gx = M + 0.6, gy = 1.85, gw = 7.4, gh = 4.6;
  s.addShape(pres.ShapeType.rect, { x: gx + gw / 2, y: gy, w: gw / 2, h: gh / 2, fill: { color: C.accent1 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.line, { x: gx, y: gy + gh / 2, w: gw, h: 0, line: { color: HEX.dk1, width: 1.25 } });
  s.addShape(pres.ShapeType.line, { x: gx + gw / 2, y: gy, w: 0, h: gh, line: { color: HEX.dk1, width: 1.25 } });
  T(s, "사람·콘텐츠 (에이전시·IP)", { x: gx, y: gy - 0.02, w: 3, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3 });
  T(s, "기술·DB (플랫폼)", { x: gx, y: gy + gh - 0.3, w: 3, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3 });
  T(s, "국내 →", { x: gx + 0.05, y: gy + gh / 2 + 0.05, w: 1.5, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3 });
  T(s, "→ 해외", { x: gx + gw - 1.25, y: gy + gh / 2 + 0.05, w: 1.2, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3, align: "right" });
  const dot = (x, y, label, big) => {
    s.addShape(pres.ShapeType.ellipse, { x: x - 0.11, y: y - 0.11, w: 0.22, h: 0.22, fill: { color: big ? C.accent2 : C.text1 }, line: { type: "none" } });
    T(s, label, { x: x + 0.18, y: y - 0.17, w: 2.6, h: 0.34, fontSize: big ? 15 : 12.5, bold: true, color: big ? C.accent2 : C.text1 });
  };
  dot(gx + 1.3, gy + 2.95, "플필");
  dot(gx + 0.9, gy + 1.65, "플필 에이전트·스튜디오");
  dot(gx + 3.95, gy + 3.1, "CASTLINK (→LA)");
  dot(gx + 1.0, gy + 4.05, "캐스틱");
  dot(gx + 5.0, gy + 3.85, "Backstage (미국)");
  dot(gx + 4.2, gy + 0.55, "KOFIC 배우 홍보 (공공·비상업)");
  dot(gx + 5.0, gy + 1.4, "WEUM ACT MATCH GLOBAL", true);
  const X = 9.05, w = CW - 8.45;
  T(s, "왜 이 칸인가", { x: X, y: 1.85, w, h: 0.4, fontSize: 17, fontFace: "Pretendard Black" });
  T(s, bullets([
    "플랫폼들은 DB를 키우는 경쟁 중. 해외에서 실제로 계약을 성사시키는 ‘사람의 일’은 비어 있다",
    "KOFIC 홍보는 비상업 공공사업 — 계약·정산까지 책임지는 주체가 없다",
    "WEUM은 매니지먼트 + 광고 + 레전드 배우로 이 칸에 바로 설 수 있다",
  ]), { x: X, y: 2.4, w, h: 3.2, fontSize: 13, paraSpaceAfter: 10, lineSpacingMultiple: 1.2 });
  T(s, "※ CASTLINK의 LA 법인 목적은 미확인 — 6개월 내 이 칸이 좁아질 수 있다", { x: X, y: 5.85, w, h: 0.7, fontSize: 11.5, italic: true, color: C.accent2 });
}

// ═════════ 10. DECONSTRUCT
sec("재설계");
{
  const s = L("LIGHT", "재설계");
  head(s, "INT. 작업실 — NIGHT", "해체: 버릴 것, 남길 것, 새로 넣을 것", "버릴 것, 남길 것, 새로 넣을 것", "추론");
  const cols = [
    ["KILL — 버린다", C.accent2, ["자체 스와이프 앱 대규모 개발 (5억)", "국내 B2B SaaS로 플필과 정면승부", "배우 DB를 0부터 쌓기", "AI 매칭 엔진 자체 개발"]],
    ["KEEP — 남긴다", C.text1, ["LEGEND SCRIPT (나문희·심영은)", "SAFE CAST 안전 원칙", "배우 데이터 헌장 (AI 동의)", "WEUM 방화벽 규칙"]],
    ["NEW — 새로 넣는다", C.accent1, ["GLOBAL DESK: 한국 배우 해외 수출 에이전시", "ACT PASSPORT: 3개 언어 연기 이력서", "3개 언어 대본 (한·일·영)", "플필·CASTLINK와 제휴, 경쟁 대신 공급 파트너"]],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  cols.forEach((c, i) => {
    const X = M + i * (cw + 0.3), y = 1.75;
    s.addText(c[0], { x: X, y, w: cw, h: 0.55, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: c[1] }, color: i === 2 ? C.text1 : C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0 });
    card(s, X, y + 0.65, cw, 4.25, C.background2);
    c[2].forEach((t, k) => {
      T(s, t, { x: X + 0.35, y: y + 0.95 + k * 0.95, w: cw - 0.7, h: 0.8, fontSize: 14.5, bold: i === 2, strike: i === 0 ? "sngStrike" : undefined, color: i === 0 ? C.accent3 : C.text1, lineSpacingMultiple: 1.2 });
    });
  });
}

// ═════════ 11. NEW CONCEPT (dark)
{
  const s = L("DARK", "재설계");
  head(s, "EXT. 세계 지도 위 — NIGHT", "K-ACTOR EXPORT ENGINE: 3개 층으로 다시 짠다", "3개 층", "추론", true);
  const Ly = [
    ["LAYER 1", "LEGEND SCRIPT GLOBAL", "콘텐츠로 끌어온다", "레전드가 건네는 3개 언어 대본. 화제 → 배우 유입 → 브랜드 스폰서"],
    ["LAYER 2", "ACT PASSPORT", "해외용으로 바꾼다", "한·일·영 프로필, 현지형 헤드샷, 3개 언어 셀프테이프, 검증 이력"],
    ["LAYER 3", "GLOBAL DESK", "계약까지 책임진다", "일본·동남아·미국 제작사와 브랜드의 요청 → 후보 패키징 → 계약·정산·비자"],
  ];
  Ly.forEach((l, i) => {
    const y = 1.8 + i * 1.62, last = i === 2;
    card(s, M, y, CW, 1.42, last ? C.accent1 : C.accent5);
    T(s, l[0], { x: M + 0.4, y: y + 0.22, w: 1.6, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: last ? C.text1 : C.accent1 });
    T(s, l[1], { x: M + 0.4, y: y + 0.55, w: 4.6, h: 0.6, fontSize: 24, fontFace: "Pretendard Black", color: last ? C.text1 : C.background1 });
    T(s, l[2], { x: M + 5.2, y: y + 0.22, w: 6.5, h: 0.4, fontSize: 17, bold: true, color: last ? C.text1 : C.accent1 });
    T(s, l[3], { x: M + 5.2, y: y + 0.68, w: 6.6, h: 0.6, fontSize: 14, color: last ? C.text1 : C.accent4 });
  });
}

// ═════════ 12. LAYER 1 LEGEND SCRIPT GLOBAL
{
  const s = L("LIGHT", "재설계");
  head(s, "INT. 촬영장 — NIGHT", "LAYER 1  한 장의 대본, 세 개의 언어", "세 개의 언어", "추론");
  const lines = [["KO", "미안해. 사실 나 그날 거기 있었어."], ["JP", "ごめん。実はあの日、そこにいたんだ。"], ["EN", "I'm sorry. I was there that day."]];
  card(s, M, 1.75, 5.6, 4.9, C.text1);
  T(s, "EP.001  ·  마지막 전화", { x: M + 0.4, y: 2.0, w: 4.8, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  lines.forEach((l, i) => {
    const y = 2.55 + i * 0.9;
    s.addText(l[0], { x: M + 0.4, y, w: 0.6, h: 0.42, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: C.accent1 }, color: C.text1, fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0, fontFace: MONO });
    T(s, l[1], { x: M + 1.2, y: y - 0.02, w: 4.1, h: 0.5, fontSize: 15.5, bold: true, color: C.background1, valign: "middle" });
  });
  T(s, "같은 장면을 한국 배우가 세 언어로 연기 →\n일본·미국 캐스팅 디렉터가 직접 본다", { x: M + 0.4, y: 5.4, w: 4.8, h: 0.9, fontSize: 13.5, color: C.accent4, lineSpacingMultiple: 1.25 });
  const X = 6.55, w = CW - 5.95;
  rowsList(s, [
    ["레전드 호스트", "나문희 배우가 대본을 고르고 시범 연기 (본인 동의 전제). EP.000 파일럿은 심영은 배우"],
    ["해외 심사위원", "일본·미국 현직 캐스팅 디렉터가 TOP 10 선정 — 국내 챌린지와의 결정적 차이"],
    ["진짜 결과", "TOP 3은 해외 제작사 실제 오디션 + ACT ORIGINALS 한·일 단편 출연"],
    ["돈이 되는 구조", "시즌 브랜드 스폰서 · 유튜브/OTT 유통 · 포맷 판매(일본·동남아 현지판)"],
  ], X, 1.8, w, 1.18, { fill: C.accent1, color: C.text1 });
}

// ═════════ 13. LAYER 2 PASSPORT
{
  const s = L("LIGHT", "재설계");
  head(s, "INT. 셀프테이프 스튜디오 — DAY", "LAYER 2  ACT PASSPORT — 세계가 읽는 연기 이력서", "세계가 읽는", "추론");
  // passport mock
  const px = M, py = 1.75, pw = 4.6, ph = 4.9;
  card(s, px, py, pw, ph, C.text1);
  T(s, "ACT PASSPORT", { x: px + 0.35, y: py + 0.25, w: 3, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  s.addShape(pres.ShapeType.rect, { x: px + 0.35, y: py + 0.7, w: 1.3, h: 1.6, fill: { color: "4A4A4A" }, line: { type: "none" } });
  T(s, "헤드샷\n(캐릭터형)", { x: px + 0.35, y: py + 1.2, w: 1.3, h: 0.6, fontSize: 10, align: "center", color: C.accent4 });
  T(s, "KIM ○○  |  김○○", { x: px + 1.85, y: py + 0.72, w: 2.6, h: 0.35, fontSize: 14, bold: true, color: C.background1 });
  T(s, "EN B2 · JP N2 · KO Native\n경상도 사투리 · 액션 · 수영", { x: px + 1.85, y: py + 1.12, w: 2.6, h: 0.65, fontSize: 11, color: C.accent4, lineSpacingMultiple: 1.2 });
  s.addText("✓ VERIFIED CREDITS 7", { x: px + 1.85, y: py + 1.9, w: 2.2, h: 0.32, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: C.accent1 }, color: C.text1, fontSize: 10.5, bold: true, align: "center", valign: "middle", margin: 0 });
  ["▶ SELF-TAPE  KO  ·  0:45", "▶ SELF-TAPE  JP  ·  0:45", "▶ SELF-TAPE  EN  ·  0:45"].forEach((t, i) => {
    s.addShape(pres.ShapeType.rect, { x: px + 0.35, y: py + 2.55 + i * 0.55, w: pw - 0.7, h: 0.45, fill: { color: "2C2C2C" }, line: { type: "none" } });
    T(s, t, { x: px + 0.5, y: py + 2.55 + i * 0.55, w: pw - 1, h: 0.45, fontSize: 11.5, fontFace: MONO, bold: true, color: C.background1, valign: "middle" });
  });
  T(s, "VISA READY: O-1B 서류 3/6 · 노조 자문 준비 중", { x: px + 0.35, y: py + 4.3, w: pw - 0.7, h: 0.35, fontSize: 11, bold: true, color: C.accent1 });
  const X = M + pw + 0.45, w = CW - pw - 0.45;
  rowsList(s, [
    ["3개 언어 프로필 · 현지형 헤드샷", "한국식 모델형 사진이 아니라 해외 캐스팅이 원하는 캐릭터형 (AP 보도된 차이)"],
    ["3개 언어 셀프테이프", "LEGEND SCRIPT 대본으로 촬영 → 해외 디렉터가 같은 조건에서 비교"],
    ["검증된 작품 이력", "출연 계약·크레딧을 WEUM이 확인한 ‘VERIFIED’ 표시"],
    ["비자·권리 준비도", "O-1B는 노조 자문 의견과 미국 측 고용주 청원 필요 (사실) — 필요 서류를 미리 쌓는다"],
  ], X, 1.8, w, 1.0);
  card(s, X, 5.95, w, 0.7, C.accent6);
  T(s, "가격: 기본 무료 · ‘패스포트 촬영 패키지’ 1회 30만 원 (구독 아님)", { x: X + 0.3, y: 5.95, w: w - 0.6, h: 0.7, fontSize: 13.5, bold: true, valign: "middle" });
}

// ═════════ 14. LAYER 3 GLOBAL DESK
{
  const s = L("LIGHT", "재설계");
  head(s, "INT. 글로벌 데스크 — 24H", "LAYER 3  GLOBAL DESK — 계약서에 사인할 때까지", "사인할 때까지", "추론");
  const st = [["REQUEST", "해외 제작사·브랜드가\n한국 배우 요청"], ["SHORTLIST", "PASSPORT에서\n후보 3~5명 패키징"], ["AUDITION", "원격 셀프테이프 →\n현지 콜백"], ["DEAL", "계약·통번역\n현지 협력 에이전트"], ["DELIVER", "비자·정산·세무\n현장 케어"]];
  const cw = 2.2, gap = (CW - 5 * cw) / 4;
  st.forEach((x, i) => {
    const X = M + i * (cw + gap), hot = i === 3;
    card(s, X, 1.8, cw, 1.85, hot ? C.accent1 : C.text1, hot);
    T(s, x[0], { x: X, y: 2.0, w: cw, h: 0.4, fontSize: 15, fontFace: MONO, bold: true, align: "center", color: hot ? C.text1 : C.accent1 });
    T(s, x[1], { x: X + 0.1, y: 2.5, w: cw - 0.2, h: 0.95, fontSize: 12.5, align: "center", color: hot ? C.text1 : C.background1, lineSpacingMultiple: 1.25 });
    if (i < 4) T(s, "›", { x: X + cw, y: 2.4, w: gap, h: 0.6, fontSize: 24, bold: true, color: C.accent3, align: "center", valign: "middle" });
  });
  const hw = (CW - 0.3) / 2;
  card(s, M, 3.95, hw, 2.7);
  T(s, "돈을 버는 방식", { x: M + 0.35, y: 4.15, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent2 });
  T(s, bullets(["해외 출연료 커미션 10~15% (전속·비전속 계약 구조별)", "제작사 패키징 수수료 (후보 선별·오디션 운영)", "해외 브랜드 한국 모델 캐스팅 대행료"]), { x: M + 0.35, y: 4.6, w: hw - 0.7, h: 1.9, fontSize: 13.5, paraSpaceAfter: 8 });
  const X2 = M + hw + 0.3;
  card(s, X2, 3.95, hw, 2.7, C.text1);
  T(s, "왜 플랫폼이 아니라 ‘데스크’인가", { x: X2 + 0.35, y: 4.15, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  T(s, bullets(["해외 계약은 클릭이 아니라 협상·비자·세무로 성사된다", "건당 금액이 커서 소수 성사로도 매출이 난다", "개발비 대신 사람에 투자 → 초기 손실 최소화"]), { x: X2 + 0.35, y: 4.6, w: hw - 0.7, h: 1.9, fontSize: 13.5, color: C.background1, paraSpaceAfter: 8 });
}

// ═════════ 15. PARTNER NOT COMPETE
{
  const s = L("LIGHT", "재설계");
  head(s, "INT. 미팅룸 — DAY", "플필과 싸우지 말고, 플필 위에 올라탄다", "플필 위에 올라탄다", "추론");
  const ow = 3.75, gap = (CW - 3 * ow) / 2;
  const O = [
    ["제휴안 A", "공급 파트너", "플필·CASTLINK 배우 중 해외 희망자를 PASSPORT로 전환. 성사 시 수익 공유", "추천"],
    ["제휴안 B", "콘텐츠 파트너", "LEGEND SCRIPT 참가 접수를 파트너 프로필로 받는다. 그들은 유입, 우리는 콘텐츠", ""],
    ["제휴안 C", "독자 노선", "제휴 없이 WEUM 네트워크만으로 출발. 느리지만 독립적", "차선"],
  ];
  O.forEach((o, i) => {
    const X = M + i * (ow + gap), rec = i === 0;
    card(s, X, 1.75, ow, 3.15, rec ? C.accent1 : C.background2, rec);
    T(s, o[0], { x: X + 0.35, y: 1.95, w: 2, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: rec ? C.text1 : C.accent3 });
    if (o[3]) s.addText(o[3], { x: X + ow - 1.15, y: 1.93, w: 0.8, h: 0.3, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: C.text1 }, color: C.background1, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
    T(s, o[1], { x: X + 0.35, y: 2.4, w: ow - 0.7, h: 0.5, fontSize: 22, fontFace: "Pretendard Black" });
    T(s, o[2], { x: X + 0.35, y: 3.1, w: ow - 0.7, h: 1.6, fontSize: 13.5, lineSpacingMultiple: 1.3 });
  });
  card(s, M, 5.15, CW, 1.5, C.text1);
  T(s, "플필이 얻는 것", { x: M + 0.35, y: 5.3, w: 3, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
  T(s, "회원에게 ‘해외 진출’이라는 새 가치, 성사 수익 공유", { x: M + 0.35, y: 5.65, w: 5.6, h: 0.8, fontSize: 13.5, color: C.background1 });
  T(s, "WEUM이 얻는 것", { x: M + 6.4, y: 5.3, w: 3, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
  T(s, "DB 구축 비용 0, 첫날부터 2만 명 규모의 후보 풀", { x: M + 6.4, y: 5.65, w: 5.4, h: 0.8, fontSize: 13.5, color: C.background1 });
}

// ═════════ 16. GLOBAL PRIORITY
sec("글로벌");
{
  const s = L("LIGHT", "글로벌");
  head(s, "EXT. 세계 지도 — DAY", "글로벌 순서: 일본 → 동남아 → 미국", "일본 → 동남아 → 미국", "사실");
  const G = [
    ["1순위", "일본", "가깝고, 이미 열렸다", ["채종협, 한국 배우 최초 TBS 황금시간대 주연 (2024)", "CJ ENM × 닛폰TV ‘메리 베리 러브’ 디즈니+ 72개국 톱10", "‘이제 실험 아닌 주류’ — 한일 합작 확대 보도"], C.accent1],
    ["2순위", "동남아 (태국)", "합작과 리메이크의 거점", ["콘진원, 2026.7 방콕에 제작사 12곳 피칭", "태국판 리메이크가 인도 등으로 재수출", "숏폼·OTT 현지 제작 증가"], C.background2],
    ["3순위", "미국", "가장 크지만 가장 높다", ["한국 배우의 할리우드 도전과 컨설팅 산업 증가 (AP, 2025.12)", "O-1B 비자: 노조 자문 + 고용주별 청원", "숏폼 앱 상위권 매출의 80%가 현지 제작 (2025)"], C.background2],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  G.forEach((g, i) => {
    const X = M + i * (cw + 0.3), first = i === 0;
    card(s, X, 1.75, cw, 4.9, g[4], first);
    T(s, g[0], { x: X + 0.35, y: 1.95, w: 2, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: first ? C.text1 : C.accent3 });
    T(s, g[1], { x: X + 0.35, y: 2.35, w: cw - 0.7, h: 0.6, fontSize: 28, fontFace: "Pretendard Black" });
    T(s, g[2], { x: X + 0.35, y: 3.0, w: cw - 0.7, h: 0.4, fontSize: 15, bold: true, color: first ? C.text1 : C.accent2 });
    T(s, bullets(g[3]), { x: X + 0.35, y: 3.6, w: cw - 0.7, h: 2.8, fontSize: 13, paraSpaceAfter: 8, lineSpacingMultiple: 1.2 });
  });
}

// ═════════ 17. GLOBAL BARRIERS
{
  const s = L("LIGHT", "글로벌");
  head(s, "INT. 출입국 심사대 — DAY", "국경의 벽 6개와 우리의 해법", "벽 6개", "추론");
  const hdr = { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: 13, valign: "middle" };
  const c = (t, o) => ({ text: t, options: Object.assign({ fontSize: 12.5, valign: "middle", color: C.text1 }, o || {}) });
  const rows = [
    [{ text: "벽", options: hdr }, { text: "실제 문제", options: hdr }, { text: "ACT MATCH GLOBAL 해법", options: Object.assign({}, hdr, { fill: { color: C.accent1 }, color: C.text1 }) }],
    [c("언어", { bold: true }), c("오디션 대본을 현지어로 연기해야 한다"), c("3개 언어 셀프테이프 + 언어 코치 제휴")],
    [c("비자", { bold: true }), c("미국 O-1B는 노조 자문·고용주별 청원 필요 (사실)"), c("이민 전문 변호사 제휴, PASSPORT에 서류 축적")],
    [c("노조·현지 규칙", { bold: true }), c("국가별 출연 계약·노조 규칙이 다르다"), c("국가별 현지 협력 에이전트와 공동 대리")],
    [c("프로필 문화", { bold: true }), c("한국은 모델형, 미국은 캐릭터형 헤드샷 (사실)"), c("현지형 헤드샷 촬영 패키지")],
    [c("고정관념", { bold: true }), c("아시아 배우 역할 편중 (AP 보도)"), c("장르·역할 다양한 장면 데이터로 반증")],
    [c("정산·세무", { bold: true }), c("해외 출연료 원천징수·환전·지급 지연"), c("WEUM 정산 대행 + 세무 파트너")],
  ];
  s.addTable(rows, { x: M, y: 1.75, w: CW, colW: [2.2, 5.1, 4.833], rowH: [0.5, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7], border: { type: "solid", pt: 0.75, color: HEX.accent4 }, margin: [0.05, 0.15, 0.05, 0.15] });
}

// ═════════ 18. INVESTOR VIEW
sec("투자");
{
  const s = L("LIGHT", "투자");
  head(s, "INT. VC 파트너 미팅 — DAY", "투자자는 이 다섯 가지를 묻는다", "다섯 가지", "추론");
  const Q = [
    ["시장이 큰가?", "국내 캐스팅 플랫폼은 작다 (선두권 매출 약 16.7억). 해외 출연·광고까지 넓혀야 의미 있는 크기", "△"],
    ["왜 너희인가?", "레전드 배우 + 매니지먼트 + 광고 10년. 기술 스타트업이 못 가진 조합", "○"],
    ["트랙션이 있나?", "현재 0. 이것이 가장 큰 약점 — 90일 검증으로 LOI·매출을 먼저 만든다", "×"],
    ["해자가 있나?", "사람(레전드)·해외 관계·검증 이력 데이터. 기능 해자는 없다", "○"],
    ["출구가 있나?", "해외 선례: Cast & Crew의 Backstage 인수(2022), Talent Systems의 Spotlight 인수 (사실)", "△"],
  ];
  Q.forEach((q, i) => {
    const y = 1.75 + i * 0.98;
    card(s, M, y, CW, 0.86, i % 2 ? C.background1 : C.background2);
    T(s, q[0], { x: M + 0.3, y, w: 2.6, h: 0.86, fontSize: 16, bold: true, valign: "middle" });
    T(s, q[1], { x: M + 3.0, y, w: CW - 4.4, h: 0.86, fontSize: 13, color: C.accent3, valign: "middle", lineSpacingMultiple: 1.2 });
    const col = q[2] === "×" ? C.accent2 : q[2] === "○" ? C.text1 : C.accent3;
    s.addShape(pres.ShapeType.ellipse, { x: M + CW - 1.0, y: y + 0.15, w: 0.56, h: 0.56, fill: { color: q[2] === "○" ? C.accent1 : C.background1 }, line: { color: HEX.dk1, width: 1 } });
    T(s, q[2], { x: M + CW - 1.0, y: y + 0.15, w: 0.56, h: 0.56, fontSize: 18, bold: true, align: "center", valign: "middle", color: col });
  });
}

// ═════════ 19. FUNDING PATH
{
  const s = L("LIGHT", "투자");
  head(s, "INT. VC 파트너 미팅 — CONTINUOUS", "지금 VC 투자는 어렵다. 순서를 바꾸면 가능하다", "순서를 바꾸면 가능하다", "추론");
  const st = [
    ["STEP 1", "0~6개월", "자체 자금 + 정부지원", "약 3억 원", "콘진원·KOFIC 해외진출 지원사업, WEUM 자체 투자. 목표: 첫 해외 계약 3건, 브랜드 스폰서 1건"],
    ["STEP 2", "6~12개월", "전략적 투자자", "3~5억 원", "일본 제작사·OTT, 광고대행사, 플랫폼 파트너. 돈보다 ‘수요’를 가져오는 투자"],
    ["STEP 3", "12~18개월", "VC 시드 / Pre-A", "10~20억 원", "해외 계약 20건+, 반복 매출 확인 후. 이때는 ‘트랙션 있는 엔터테크’로 심사"],
  ];
  st.forEach((x, i) => {
    const y = 1.75 + i * 1.6, last = i === 2;
    card(s, M, y, CW, 1.42, last ? C.text1 : C.background2);
    T(s, x[0], { x: M + 0.35, y: y + 0.2, w: 1.5, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: last ? C.accent1 : C.accent3 });
    T(s, x[1], { x: M + 0.35, y: y + 0.55, w: 1.8, h: 0.4, fontSize: 15, bold: true, color: last ? C.background1 : C.text1 });
    T(s, x[2], { x: M + 2.4, y: y + 0.2, w: 3.6, h: 0.45, fontSize: 19, fontFace: "Pretendard Black", color: last ? C.background1 : C.text1 });
    T(s, x[3], { x: M + 2.4, y: y + 0.72, w: 3.4, h: 0.45, fontSize: 17, bold: true, color: last ? C.accent1 : C.accent2 });
    T(s, x[4], { x: M + 6.1, y: y + 0.15, w: CW - 6.4, h: 1.12, fontSize: 13, color: last ? C.accent4 : C.accent3, valign: "middle", lineSpacingMultiple: 1.25 });
  });
}

// ═════════ 20. 3-YEAR P&L
{
  const s = L("LIGHT", "투자");
  head(s, "INT. 재무 회의실 — NIGHT", "보수적으로 잡아도 3년차 흑자 전환", "3년차 흑자 전환", "추론");
  s.addChart(pres.ChartType.bar, [
    { name: "매출", labels: ["1년차", "2년차", "3년차"], values: [2.5, 7.0, 15.0] },
    { name: "비용", labels: ["1년차", "2년차", "3년차"], values: [3.8, 7.5, 12.5] },
  ], {
    x: M, y: 1.7, w: 5.6, h: 4.95, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.accent1, HEX.dk1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0.0",
    showTitle: true, title: "매출 vs 비용 (억 원)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    catAxisLabelColor: HEX.accent3, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: true, legendPos: "t", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.accent3,
  });
  const X = 6.55, w = CW - 5.95;
  const tbl = [
    ["매출원 (억 원)", "1년차", "2년차", "3년차"],
    ["LEGEND 브랜드 스폰서", "1.5", "3.0", "5.0"],
    ["PASSPORT 촬영 패키지", "0.45", "1.2", "2.4"],
    ["GLOBAL DESK 커미션·패키징", "0.55", "2.0", "5.0"],
    ["포맷 판매·파트너 수익", "—", "0.8", "2.6"],
    ["매출 합계", "2.5", "7.0", "15.0"],
    ["영업손익", "−1.3", "−0.5", "+2.5"],
  ];
  s.addTable(tbl.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 12.5, bold: i === 0 || i >= 5 || j === 0, align: j ? "center" : "left", valign: "middle", color: i === 6 && j === 3 ? C.accent2 : C.text1, fill: { color: i === 0 ? HEX.accent1 : (i >= 5 ? HEX.lt2 : "FFFFFF") } } }))),
    { x: X, y: 1.8, w, colW: [2.75, 1.0, 1.0, w - 4.75], rowH: 0.44, border: { type: "solid", pt: 0.75, color: HEX.accent4 }, margin: [0.04, 0.1, 0.04, 0.1] });
  T(s, bullets(["DESK: 1년차 6건 → 3년차 50건, 건당 평균 출연료 5~6천만 원 × 15%", "PASSPORT: 1년차 150명 → 3년차 800명 × 30만 원", "비교: 국내 선두권 CASTLINK 2025년 매출 약 16.7억 (사실)"]), { x: X, y: 5.05, w, h: 1.6, fontSize: 12, color: C.accent3, paraSpaceAfter: 5 });
}

// ═════════ 21. SCENARIOS
{
  const s = L("LIGHT", "투자");
  head(s, "INT. 전략실 — NIGHT", "최악·기본·최선, 세 개의 시나리오", "세 개의 시나리오", "추론");
  const S3 = [
    ["BEAR · 최악", "3년차 매출 6억", "나문희 배우 불참, 브랜드 스폰서 0", "GLOBAL DESK만 운영 (일본 중심). 손실은 초기 3억 안에서 통제", C.background2],
    ["BASE · 기본", "3년차 매출 15억", "LEGEND 시즌 연 2회, 일본 계약 안정화", "3년차 흑자 +2.5억, Pre-A 유치 가능 구간", C.accent1],
    ["BULL · 최선", "3년차 매출 30억+", "일본 OTT 포맷 계약 + 미국 숏폼 공급 계약", "포맷 로열티가 붙으며 전략적 M&A 대상", C.text1],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  S3.forEach((x, i) => {
    const X = M + i * (cw + 0.3), dark = i === 2;
    card(s, X, 1.75, cw, 4.9, x[4], i === 1);
    T(s, x[0], { x: X + 0.35, y: 2.0, w: cw - 0.7, h: 0.35, fontSize: 13, fontFace: MONO, bold: true, color: dark ? C.accent1 : C.text1 });
    T(s, x[1], { x: X + 0.35, y: 2.5, w: cw - 0.7, h: 0.6, fontSize: 25, fontFace: "Pretendard Black", color: dark ? C.background1 : C.text1 });
    T(s, "조건", { x: X + 0.35, y: 3.4, w: 2, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: dark ? C.accent4 : C.accent3 });
    T(s, x[2], { x: X + 0.35, y: 3.72, w: cw - 0.7, h: 0.9, fontSize: 14, bold: true, color: dark ? C.background1 : C.text1, lineSpacingMultiple: 1.2 });
    T(s, "결과", { x: X + 0.35, y: 4.75, w: 2, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: dark ? C.accent4 : C.accent3 });
    T(s, x[3], { x: X + 0.35, y: 5.07, w: cw - 0.7, h: 1.3, fontSize: 13.5, color: dark ? C.accent4 : C.text1, lineSpacingMultiple: 1.25 });
  });
}

// ═════════ 22. 90-DAY VALIDATION
sec("실행");
{
  const s = L("LIGHT", "실행");
  head(s, "INT. 작전실 — DAY", "돈 쓰기 전에, 90일 동안 이것부터 증명한다", "90일", "추론");
  const V = [
    ["일본 수요", "일본 제작사·캐스팅 디렉터 10곳 인터뷰 → 한국 배우 요청 의향서(LOI) 3건", "LOI 3건"],
    ["브랜드 돈", "LEGEND SCRIPT 시즌 1 스폰서 사전 의향 1건 (WEUM 광고 네트워크)", "의향 1건"],
    ["배우 수요", "PASSPORT 촬영 패키지 사전예약 100명 (유료 예약금)", "100명"],
    ["파트너", "플필·CASTLINK 중 1곳과 공급 제휴 MOU", "MOU 1건"],
  ];
  V.forEach((v, i) => {
    const y = 1.75 + i * 1.0;
    card(s, M, y, 8.3, 0.86, C.background2);
    num(s, i + 1, M + 0.25, y + 0.2, 0.46, C.accent1, C.text1);
    T(s, v[0], { x: M + 0.95, y, w: 1.6, h: 0.86, fontSize: 15.5, bold: true, valign: "middle" });
    T(s, v[1], { x: M + 2.55, y, w: 4.6, h: 0.86, fontSize: 12.5, color: C.accent3, valign: "middle", lineSpacingMultiple: 1.15 });
    T(s, v[2], { x: M + 7.1, y, w: 1.1, h: 0.86, fontSize: 14, bold: true, color: C.accent2, valign: "middle", align: "right" });
  });
  const X = M + 8.6, w = CW - 8.6;
  card(s, X, 1.75, w, 3.86, C.text1);
  T(s, "판정 규칙", { x: X + 0.3, y: 1.95, w: w - 0.6, h: 0.35, fontSize: 14, bold: true, color: C.accent1 });
  T(s, [
    { text: "4개 중 3개 이상", options: { bold: true, color: C.accent1, breakLine: true, fontSize: 17 } },
    { text: "→ GO. STEP 1 집행", options: { color: C.background1, breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "2개 이하", options: { bold: true, color: C.accent1, breakLine: true, fontSize: 17 } },
    { text: "→ 일본 DESK만 소규모 운영, 나머지 보류", options: { color: C.background1 } },
  ], { x: X + 0.3, y: 2.45, w: w - 0.6, h: 3.0, fontSize: 13.5, lineSpacingMultiple: 1.3 });
  card(s, M, 5.85, CW, 0.8, C.accent6);
  T(s, "90일 검증 예산: 약 3,000만 원 (출장·통역·샘플 촬영). 실패해도 손실은 여기서 멈춘다.", { x: M + 0.35, y: 5.85, w: CW - 0.7, h: 0.8, fontSize: 15, bold: true, valign: "middle" });
}

// ═════════ 23. 12-MONTH ROADMAP
{
  const s = L("LIGHT", "실행");
  head(s, "INT. 작전실 — CONTINUOUS", "12개월 로드맵: 일본에서 증명하고 넓힌다", "일본에서 증명하고", "추론");
  const Q4 = [
    ["Q1", "VALIDATE", ["90일 검증 4개 지표", "EP.000 파일럿 (심영은)", "PASSPORT 촬영 샘플 30명"]],
    ["Q2", "LAUNCH JP", ["LEGEND SCRIPT 시즌 1 (한·일 2개 언어)", "일본 협력 에이전트 계약", "첫 일본 출연 계약"]],
    ["Q3", "SCALE", ["GLOBAL DESK 계약 누적 6건", "브랜드 스폰서 시즌 2", "태국 파일럿 제작사 3곳"]],
    ["Q4", "RAISE", ["성과 리포트 공개", "전략적 투자자 유치", "미국 숏폼 제작사 공급 테스트"]],
  ];
  const cw = (CW - 3 * 0.25) / 4;
  s.addShape(pres.ShapeType.line, { x: M, y: 2.2, w: CW, h: 0, line: { color: HEX.dk1, width: 1.5 } });
  Q4.forEach((q, i) => {
    const X = M + i * (cw + 0.25), hot = i === 1;
    s.addShape(pres.ShapeType.ellipse, { x: X, y: 2.06, w: 0.28, h: 0.28, fill: { color: hot ? C.accent1 : C.text1 }, line: { color: HEX.dk1, width: 1 } });
    T(s, q[0], { x: X + 0.4, y: 1.6, w: 1, h: 0.4, fontSize: 18, fontFace: "Pretendard Black" });
    card(s, X, 2.55, cw, 4.1, hot ? C.accent6 : C.background2);
    T(s, q[1], { x: X + 0.3, y: 2.75, w: cw - 0.6, h: 0.35, fontSize: 14, fontFace: MONO, bold: true, color: hot ? C.accent2 : C.text1 });
    T(s, bullets(q[2]), { x: X + 0.3, y: 3.3, w: cw - 0.6, h: 3.1, fontSize: 13.5, paraSpaceAfter: 10, lineSpacingMultiple: 1.2 });
  });
}

// ═════════ 24. TEAM
{
  const s = L("LIGHT", "실행");
  head(s, "INT. WEUM 본사 — DAY", "필요한 것은 개발자 10명이 아니라 핵심 4명이다", "핵심 4명", "추론");
  const P = [
    ["대표 (WEUM)", "브랜드·스폰서 영업, 레전드 배우 관계, 최종 의사결정", "보유"],
    ["글로벌 디렉터", "일본어 능통, 현지 제작사·에이전트 네트워크. 가장 먼저 채용", "채용"],
    ["콘텐츠 PD", "LEGEND SCRIPT 기획·제작, 셀프테이프 스튜디오 운영", "채용"],
    ["딜 코디네이터", "계약·통번역·정산·비자 실무", "채용"],
  ];
  const cw = (CW - 3 * 0.3) / 4;
  P.forEach((p, i) => {
    const X = M + i * (cw + 0.3), own = i === 0;
    card(s, X, 1.75, cw, 3.3, own ? C.text1 : C.background2);
    s.addText(p[2], { x: X + 0.3, y: 2.0, w: 0.8, h: 0.3, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: own ? C.accent1 : C.text1 }, color: own ? C.text1 : C.background1, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
    T(s, p[0], { x: X + 0.3, y: 2.5, w: cw - 0.6, h: 0.5, fontSize: 18, fontFace: "Pretendard Black", color: own ? C.background1 : C.text1 });
    T(s, p[1], { x: X + 0.3, y: 3.15, w: cw - 0.6, h: 1.7, fontSize: 13, color: own ? C.accent4 : C.accent3, lineSpacingMultiple: 1.3 });
  });
  T(s, "외부 파트너 (고정비 없이 건별 계약)", { x: M, y: 5.3, w: 8, h: 0.35, fontSize: 15, bold: true });
  T(s, "이민 전문 변호사  ·  일본·미국 협력 에이전트  ·  언어 코치  ·  세무사  ·  앱은 노코드 + 외주 최소 개발", { x: M, y: 5.75, w: CW, h: 0.7, fontSize: 14, color: C.accent3 });
}

// ═════════ 25. RISKS
{
  const s = L("LIGHT", "실행");
  head(s, "INT. 법무 검토 — DAY", "글로벌로 가면 리스크도 국경을 넘는다", "국경을 넘는다");
  const r = [
    ["기획업 등록 · 수수료", "사실", "기획업 미등록 시 2년 이하 징역 또는 2천만 원 이하 벌금 (보도). 문체부 엄정 조치 방침", "WEUM 등록 상태 확인 후 GLOBAL DESK를 WEUM 법인에서 운영"],
    ["해외 노동·비자", "사실", "미국 O-1B: 노조 자문 의견 + 고용주별 청원, 승인 전 근무 불가", "이민 변호사 제휴, 계약 전 비자 일정 확정"],
    ["초상권 · AI", "추론", "해외 광고·게임의 얼굴·목소리 활용 범위 분쟁", "데이터 헌장 + 작품별 개별 동의 + 사용 범위 계약서"],
    ["레전드 브랜드 보호", "추론", "콘텐츠 품질 사고가 나문희 배우 이미지에 직결", "출연 범위 사전 합의, 공동 검수권, 부정 이슈 즉시 중단 조항"],
  ];
  ["리스크", "근거", "대응"].forEach((h, i) => T(s, h, { x: [M, 4.15, 8.95][i], y: 1.72, w: 3, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3 }));
  r.forEach((x, i) => {
    const y = 2.1 + i * 1.13;
    card(s, M, y, CW, 1.0, i % 2 ? C.background1 : C.background2);
    T(s, x[0], { x: M + 0.3, y, w: 2.65, h: 1.0, fontSize: 15, bold: true, valign: "middle" });
    tag(s, 3.4, y + 0.36, x[1]);
    T(s, x[2], { x: 4.15, y, w: 4.55, h: 1.0, fontSize: 12.5, color: C.accent3, valign: "middle" });
    T(s, x[3], { x: 8.95, y, w: 3.7, h: 1.0, fontSize: 12.5, bold: true, valign: "middle" });
  });
}

// ═════════ 26. BRAND + COPY (dark)
sec("브랜딩");
{
  const s = L("DARK", "브랜딩");
  head(s, "INT. 카피룸 — NIGHT", "글로벌 카피 3개, 우선순위대로", "카피 3개", "추론", true);
  const cp = [
    ["1순위 · 브랜드 메인", "한국 배우를,\n세계의 장면으로.", "사업의 방향(수출)을 한 문장에 담았다. 배우에게는 꿈, 해외 제작사에게는 약속. 영문 ‘Korean actors, into the world’s scenes.’로도 그대로 통한다."],
    ["2순위 · LEGEND 캠페인", "한 장의 대본,\n세 개의 언어.", "3개 언어 대본이라는 차별점을 숫자로 보여준다. 숏폼 썸네일·포스터 헤드라인용."],
    ["3순위 · 배우 모집", "오디션장 밖에서,\n국경 밖에서\n캐스팅된다.", "v3의 1순위 카피를 확장해 연속성을 지킨다. 배우 대상 모집 광고·앱스토어용."],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  cp.forEach((c, i) => {
    const X = M + i * (cw + 0.3), first = i === 0;
    card(s, X, 1.75, cw, 4.9, first ? C.accent1 : C.accent5);
    T(s, c[0], { x: X + 0.4, y: 2.0, w: cw - 0.8, h: 0.3, fontSize: 12.5, fontFace: MONO, bold: true, color: first ? C.text1 : C.accent1 });
    T(s, c[1], { x: X + 0.4, y: 2.45, w: cw - 0.8, h: 1.75, fontSize: 25, fontFace: "Pretendard Black", color: first ? C.text1 : C.background1, lineSpacingMultiple: 1.15 });
    T(s, c[2], { x: X + 0.4, y: 4.35, w: cw - 0.8, h: 2.1, fontSize: 13, color: first ? C.text1 : C.accent4, lineSpacingMultiple: 1.35 });
  });
}

// ═════════ 27. LOGO MEANING
{
  const s = L("LIGHT", "브랜딩");
  head(s, "INT. 디자인 스튜디오 — DAY", "로고는 그대로, 의미는 넓힌다", "의미는 넓힌다", "추론");
  card(s, M, 1.75, 6.6, 2.7);
  s.addImage({ path: A + "lockup_light.png", x: M + 0.4, y: 2.3, w: 5.8, h: 1.55 });
  s.addImage({ path: A + "app_icon.png", x: M + 6.9, y: 1.75, w: 2.7, h: 2.7 });
  T(s, "브랜드 구조", { x: M + 9.95, y: 1.8, w: 2.2, h: 0.35, fontSize: 14, bold: true });
  T(s, "ACT MATCH GLOBAL\n├ LEGEND SCRIPT\n├ ACT PASSPORT\n└ GLOBAL DESK", { x: M + 9.95, y: 2.2, w: 2.3, h: 1.8, fontSize: 12, fontFace: MONO, color: C.accent3, lineSpacingMultiple: 1.3 });
  const mean = [["v3의 의미", "배우의 장면과 연출자의 장면이 겹치는 곳 = MATCH"], ["v4의 의미", "한국의 장면과 세계의 장면이 겹치는 곳 = 노란 교집합이 곧 ‘수출 성사’"], ["유지 이유", "브랜드 자산을 버리지 않고 사업 확장을 담는다. 상표 출원 시 9·35·41·42류 + 해외(일·미) 마드리드 출원 검토"]];
  mean.forEach((x, i) => {
    const y = 4.75 + i * 0.65;
    T(s, x[0], { x: M, y, w: 2.0, h: 0.5, fontSize: 15, bold: true, color: i === 1 ? C.accent2 : C.text1 });
    T(s, x[1], { x: M + 2.0, y, w: CW - 2.0, h: 0.6, fontSize: 13.5, color: C.accent3 });
  });
}

// ═════════ 28. SELF SCORE
sec("평가");
{
  const s = L("LIGHT", "평가");
  head(s, "INT. 리뷰 — DAY", "v4 재설계안 88점 — 남은 12점의 정체", "남은 12점", "추론");
  const g = [
    ["−5", "트랙션 0", "해외 수요·지불의향이 모두 가설이다", ["90일 검증: 일본 LOI 3건을 최우선으로", "WEUM 기존 광고주 대상 스폰서 의향 확보", "PASSPORT 예약금 받는 사전예약"]],
    ["−4", "글로벌 실무 인력 부재", "일본어·현지 네트워크를 가진 사람이 아직 없다", ["일본 네트워크 보유 글로벌 디렉터 1순위 채용", "일본 협력 에이전트와 수익 공유 계약", "콘진원 해외 비즈니스센터 활용"]],
    ["−3", "핵심 자산 미확정", "나문희 배우 참여와 플필 등 제휴가 미정", ["심영은 배우 EP.000으로 먼저 증명 후 제안", "제휴는 A안(공급) 우선, 불발 시 C안", "레전드 불참 시나리오(BEAR)도 손실 통제"]],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  g.forEach((x, i) => {
    const X = M + i * (cw + 0.3);
    card(s, X, 1.75, cw, 4.9);
    T(s, x[0], { x: X + 0.35, y: 1.95, w: 1.2, h: 0.6, fontSize: 30, fontFace: "Pretendard Black", color: C.accent2 });
    T(s, x[1], { x: X + 0.35, y: 2.65, w: cw - 0.7, h: 0.4, fontSize: 16, bold: true });
    T(s, "이유: " + x[2], { x: X + 0.35, y: 3.08, w: cw - 0.7, h: 0.6, fontSize: 12, color: C.accent3 });
    x[3].forEach((t, k) => {
      const y = 3.85 + k * 0.9;
      num(s, k + 1, X + 0.35, y + 0.02, 0.34, k === 0 ? C.accent1 : C.text1, k === 0 ? C.text1 : C.background1);
      T(s, t, { x: X + 0.85, y, w: cw - 1.15, h: 0.8, fontSize: 12.5, bold: k === 0, lineSpacingMultiple: 1.2 });
    });
  });
}

// ═════════ 29. SOURCES
{
  const s = L("LIGHT", "평가");
  head(s, "INT. 자료실 — DAY", "‘사실’ 표기의 근거 자료", null);
  const src = [
    "한국경제(2025.12) · 스마트투데이(2024) — 플필 투자: 프라이머(2024.1), 땡스벤처스·킹슬리벤처스 / 신인 회원 약 2만·기업 2천여, 에이전트 플필(2024~)",
    "THE VC — 플필 2022.4 설립, 대표 류민국, 임직원 8명(2026.6)  ·  plfil.com/plfilplus — 플필 플러스 월 29,800원",
    "스타데일리뉴스 — 스튜디오 플필 숏폼 조회수 1,000만  ·  비즈니스코리아 — 컴투스 ‘제우스: 오만의 신’ 캐스팅 중개",
    "KOCCA Welcon 디렉터리 · 잡코리아 — CASTLINK(루미크래프트) 프로필 3.4만+, 2025년 매출 약 16.5~16.7억, LA 법인, 수수료 모델",
    "디지털투데이 — 캐스틱(베피플) 글로벌 진출 계획  ·  plfil.notion.site — 해외 인재 가입 개방",
    "AP(Korea Times·Washington Times, 2025.12) — 한국 배우 할리우드 진출 증가, 헤드샷 문화 차이, 고정관념",
    "스포츠경향 — 채종협 TBS 황금시간대 첫 한국인 주연  ·  한국경제(2026.10) — ‘메리 베리 러브’ 72개국 톱10",
    "더팩트 — 한·일·미·태 공동제작 확대  ·  Outlook — 콘진원 방콕 12개 제작사 피칭(2026.7)",
    "USCIS · 이민법률 자료 — O-1B 노조 자문·고용주별 청원  ·  Wikipedia(Microdrama) — 상위권 현지 제작 80%",
    "mergr · Casting Networks — Cast & Crew의 Backstage 인수(2022), Talent Systems의 Spotlight 인수",
    "다음(파이낸셜뉴스) — 2026 콘텐츠산업포럼 ‘콘텐츠 투자금 급감’  ·  뉴시스 — 벤처투자 AI 쏠림(2026 상반기)",
  ];
  T(s, bullets(src), { x: M, y: 1.75, w: CW, h: 4.6, fontSize: 12, paraSpaceAfter: 5, lineSpacingMultiple: 1.1 });
  T(s, "※ 플필 내부 매출·기업 과금 구조는 비공개로 확인 불가. 모든 재무·점수는 검증용 가설(추론)입니다.", { x: M, y: 6.35, w: CW, h: 0.32, fontSize: 11.5, italic: true, color: C.accent3 });
}

// ═════════ 30. CLOSING
sec("마무리");
{
  const s = L("COVER", "마무리");
  s.addImage({ path: A + "mark_dark.png", x: W / 2 - 0.65, y: 0.6, w: 1.3, h: 1.3 });
  T(s, "FADE OUT.", { x: M, y: 2.05, w: CW, h: 0.35, fontSize: 14, fontFace: MONO, bold: true, color: C.accent1, align: "center" });
  T(s, "한국 배우를, 세계의 장면으로.", { x: M, y: 2.5, w: CW, h: 0.9, fontSize: 44, fontFace: "Pretendard Black", color: C.background1, align: "center", valign: "middle" });
  T(s, "ACT MATCH GLOBAL  by  WEUM CORPORATION", { x: M, y: 3.45, w: CW, h: 0.35, fontSize: 13, fontFace: MONO, bold: true, color: C.accent4, align: "center" });
  card(s, 1.4, 4.25, W - 2.8, 2.45, C.accent5);
  T(s, "권리 고지 · NOTICE OF RIGHTS", { x: 1.8, y: 4.45, w: W - 3.6, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "본 기획안에 포함된 사업 아이디어, 사업 재설계 구조, 브랜드명(ACT MATCH · ACT MATCH GLOBAL), 로고, 카피, 콘텐츠 포맷(LEGEND SCRIPT · ACT PASSPORT · GLOBAL DESK · SAFE CAST 등), 경쟁 분석, 수익모델, 재무 가설 및 모든 구성에 관한 저작권과 일체의 권리는 WEUM CORPORATION에 있습니다. 사전 서면 동의 없는 복제·배포·전송·변형·2차적 저작물 작성 및 영업상 이용을 금하며, 본 자료는 「저작권법」 및 「부정경쟁방지 및 영업비밀보호에 관한 법률」에 따라 보호되는 영업비밀입니다.",
    { x: 1.8, y: 4.85, w: W - 3.6, h: 1.3, fontSize: 11, color: C.accent4, lineSpacingMultiple: 1.3 });
  T(s, "© 2026 WEUM CORPORATION. All Rights Reserved.", { x: 1.8, y: 6.2, w: W - 3.6, h: 0.3, fontSize: 12, bold: true, color: C.background1 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written");
})();
