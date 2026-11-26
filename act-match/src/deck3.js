const pptxgen = require("pptxgenjs");
const { execFileSync } = require("child_process");
const { applyTheme } = require("/root/.claude/skills/synced/06c47b33-d640-428c-a3a5-a42ae98094c8_2ae10ec8-84d2-43dd-8c58-6c8a9c68ff68/pptx/scripts/apply_theme.js");
const A = "/home/user/worklog/act-match/assets/";
const OUT = "/home/user/worklog/act-match/ACT_MATCH_기획안_v3_WEUM.pptx";

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
pres.title = "ACT MATCH 사업 기획안 v3.0";
pres.author = "WEUM CORPORATION";
pres.company = "WEUM CORPORATION";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6, CW = W - 2 * M;
const MONO = "Courier New";
const FOOT = "ACT MATCH 사업 기획안 v3.0  ·  내부 검토용  ·  CONFIDENTIAL  ·  © 2026 WEUM CORPORATION";
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

// ═════════ 1. COVER
pres.addSection({ title: "표지" });
{
  const s = L("COVER", "표지");
  s.addImage({ path: A + "lockup_dark.png", x: M, y: 0.55, w: 3.75, h: 1.0 });
  T(s, "S#01.  EXT. 오디션장 밖 — NIGHT", { x: M, y: 2.2, w: 9, h: 0.35, fontSize: 15, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "오디션장 밖에서,", { x: M, y: 2.7, w: 11, h: 1.0, fontSize: 60, fontFace: "Pretendard Black", color: C.background1 });
  const w2 = measure("캐스팅된다.", 60, "bk");
  s.addShape(pres.ShapeType.rect, { x: M - 0.08, y: 3.78, w: w2 + 0.2, h: 0.98, fill: { color: C.accent1 }, line: { type: "none" } });
  T(s, "캐스팅된다.", { x: M, y: 3.72, w: 11, h: 1.0, fontSize: 60, fontFace: "Pretendard Black", color: C.text1 });
  T(s, "장면으로 발견하고, 서로 골라야 열리는 캐스팅 플랫폼 — 사업 기획안 v3.0", { x: M, y: 5.15, w: 11, h: 0.4, fontSize: 18, color: C.accent4 });
  T(s, "2026. 11  ·  WEUM CORPORATION  ·  내부 검토용  ·  CONFIDENTIAL", { x: M, y: 6.8, w: 9, h: 0.3, fontSize: 11, fontFace: MONO, color: C.accent3 });
}

// ═════════ 2. DECISIONS
pres.addSection({ title: "의사결정 요약" });
{
  const s = L("LIGHT", "의사결정 요약");
  head(s, "INT. 대표실 — DAY", "이 기획안이 대표님께 요청하는 결정은 3개입니다", "결정은 3개", "추론");
  card(s, M, 1.75, 4.3, 4.95, C.text1);
  T(s, "ONE LINE", { x: M + 0.4, y: 2.05, w: 3.5, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "같은 대본으로 비교하고,\n서로 골라야 열리는\n캐스팅.\n우리는 ‘신뢰’를 판다.", { x: M + 0.4, y: 2.5, w: 3.6, h: 2.6, fontSize: 24, fontFace: "Pretendard ExtraBold", color: C.background1, lineSpacingMultiple: 1.2 });
  const kv = [["MVP 기간", "6개월"], ["MVP 예산", "약 5억 원"], ["손익분기 목표", "3년차"]];
  kv.forEach((k, i) => {
    const y = 5.35 + i * 0.4;
    T(s, k[0], { x: M + 0.4, y, w: 1.8, h: 0.35, fontSize: 13, color: C.accent4 });
    T(s, k[1], { x: M + 2.2, y, w: 1.8, h: 0.35, fontSize: 13, bold: true, color: C.accent1 });
  });
  const D = [
    ["나문희 배우를 시즌 1 호스트로 제안할 것인가", "심영은 배우 EP.000 파일럿 후, 2027년 1분기 제안", "킬러 콘텐츠의 성패가 걸린 결정"],
    ["플랫폼을 별도 법인으로 세울 것인가", "별도 법인 설립. WEUM 매니지먼트와 방화벽을 법인 단위로 분리", "공정성 신뢰 + 외부 투자 유치에 유리"],
    ["MVP 예산 약 5억 원을 어떻게 집행할 것인가", "앱 개발은 외주, PM·운영·콘텐츠는 인하우스. 정부지원사업 병행", "속도와 노하우 내재화의 균형"],
  ];
  D.forEach((d, i) => {
    const y = 1.75 + i * 1.68, x = 5.2;
    card(s, x, y, CW - 4.6, 1.52, C.background2);
    num(s, i + 1, x + 0.3, y + 0.3, 0.5);
    T(s, d[0], { x: x + 1.05, y: y + 0.24, w: 6.6, h: 0.4, fontSize: 17, bold: true });
    s.addShape(pres.ShapeType.rect, { x: x + 1.05, y: y + 0.76, w: 0.62, h: 0.26, fill: { color: C.accent1 }, line: { type: "none" } });
    T(s, "권고", { x: x + 1.05, y: y + 0.76, w: 0.62, h: 0.26, fontSize: 11, bold: true, align: "center", valign: "middle" });
    T(s, d[1], { x: x + 1.8, y: y + 0.73, w: 5.85, h: 0.35, fontSize: 13.5, bold: true });
    T(s, d[2], { x: x + 1.8, y: y + 1.08, w: 5.85, h: 0.3, fontSize: 12, color: C.accent3 });
  });
}

// ═════════ 3. WHY NOW
pres.addSection({ title: "시장" });
{
  const s = L("LIGHT", "시장");
  head(s, "EXT. 한국 캐스팅 시장 — 2026", "배우는 넘치고, 일자리는 줄고, 신뢰는 무너졌다", "신뢰는 무너졌다", "사실");
  const st = [
    ["25.6 : 1", "연극·영화·방송 경쟁률", "2026학년도 수시 일반전형\n정원 2,033명에 52,010명 지원", "진학사 (대학 공시 기반)"],
    ["12,092명", "기획사 소속 예술인", "기획업 등록 4,471곳 전체의\n소속 인원 (2024년 기준)", "문체부·콘진원 실태조사 2025"],
    ["−130억", "실연자 사용료 감소", "2024년 약 470억 → 2025년 약 340억.\n제작 편수 감소가 원인", "파이낸셜뉴스 2026.11"],
    ["+75%", "기획사 피해구제 증가", "상반기 91건 (전년 동기 52건).\n피해자 절반이 아동·청소년", "한국소비자원"],
  ];
  const cw = (CW - 3 * 0.3) / 4;
  st.forEach((x, i) => {
    const X = M + i * (cw + 0.3);
    card(s, X, 1.75, cw, 3.6);
    const nw = measure(x[0], 36, "bk");
    if (i === 0) s.addShape(pres.ShapeType.rect, { x: X + 0.26, y: 2.3, w: nw + 0.08, h: 0.3, fill: { color: C.accent1 }, line: { type: "none" } });
    T(s, x[0], { x: X + 0.3, y: 1.98, w: cw - 0.5, h: 0.7, fontSize: 36, fontFace: "Pretendard Black", color: i === 3 ? C.accent2 : C.text1 });
    T(s, x[1], { x: X + 0.3, y: 2.88, w: cw - 0.6, h: 0.35, fontSize: 14.5, bold: true });
    T(s, x[2], { x: X + 0.3, y: 3.3, w: cw - 0.6, h: 1.0, fontSize: 12.5, color: C.accent3, lineSpacingMultiple: 1.25 });
    T(s, x[3], { x: X + 0.3, y: 4.9, w: cw - 0.6, h: 0.3, fontSize: 10, fontFace: MONO, color: C.accent3 });
  });
  card(s, M, 5.6, CW, 1.05, C.text1);
  tag(s, M + 0.35, 5.98, "추론");
  T(s, [
    { text: "소속사 문턱 밖에 있는 배우가 시장의 다수다. 그들에게 필요한 건 더 많은 공고가 아니라 " },
    { text: "‘검증된 발견’", options: { color: C.accent1, bold: true } },
    { text: "이다." },
  ], { x: M + 1.2, y: 5.6, w: CW - 1.5, h: 1.05, fontSize: 16.5, color: C.background1, valign: "middle" });
}

// ═════════ 4. PROBLEM
{
  const s = L("LIGHT", "시장");
  head(s, "INT. 오디션 대기실 — DAY", "지금의 캐스팅은 세 군데가 고장 나 있다", "세 군데", "추론");
  const P = [
    ["사진으로 먼저 판단한다", "프로필 사진·경력·소속이 먼저, 연기는 나중이다. 신인은 연기를 보여줄 기회조차 얻기 어렵다.", "장면으로 먼저 본다"],
    ["한쪽만 고른다", "배우는 감독의 현장 평판을 모른 채 지원한다. 정보는 늘 연출자 쪽에만 있다.", "서로 고른다"],
    ["믿을 수 없다", "허위 오디션, 트레이닝비 요구, 밀실 미팅. 피해구제 신청은 상반기에만 75% 늘었다.", "검증된 만남만 허용한다"],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  P.forEach((p, i) => {
    const X = M + i * (cw + 0.3);
    card(s, X, 1.75, cw, 4.9, C.background1, true);
    T(s, "0" + (i + 1), { x: X + 0.4, y: 2.0, w: 1, h: 0.6, fontSize: 34, fontFace: "Pretendard Black", color: C.accent3 });
    T(s, p[0], { x: X + 0.4, y: 2.75, w: cw - 0.8, h: 0.45, fontSize: 20, bold: true });
    T(s, p[1], { x: X + 0.4, y: 3.35, w: cw - 0.8, h: 1.6, fontSize: 14, color: C.accent3, lineSpacingMultiple: 1.35 });
    s.addShape(pres.ShapeType.line, { x: X + 0.4, y: 5.25, w: cw - 0.8, h: 0, line: { color: HEX.accent4, width: 0.75, dashType: "dash" } });
    T(s, "처방", { x: X + 0.4, y: 5.45, w: 1, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent2 });
    T(s, "→ " + p[2], { x: X + 0.4, y: 5.8, w: cw - 0.8, h: 0.45, fontSize: 17, bold: true, color: C.accent2 });
  });
}

// ═════════ 5. COMPETITION
{
  const s = L("LIGHT", "시장");
  head(s, "EXT. 경쟁 서비스 — DAY", "‘영상 프로필 + AI + 스와이프’는 이미 레드오션이다", "레드오션", "사실");
  const hdr = { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: 13, valign: "middle" };
  const c = (t, o) => ({ text: t, options: Object.assign({ fontSize: 13, valign: "middle", color: C.text1 }, o || {}) });
  const rows = [
    [{ text: "서비스", options: hdr }, { text: "확인된 사실", options: hdr }, { text: "빈틈 (추론)", options: Object.assign({}, hdr, { fill: { color: C.accent1 }, color: C.text1 }) }],
    [c("플필 PLFIL", { bold: true }), c("배우 약 1.67만 명을 캐스팅 공고와 연결, 유료 ‘플필 플러스’, 자체 숏폼 오디션 한 달 3,000명+ 지원 (2024.8)"), c("배우→공고 지원형. 연출자 정보는 비대칭")],
    [c("캐스틱 · 베피플", { bold: true }), c("캐스팅 검토 7일→1일 단축 / 숏폼 영상 프로필, 디렉터 48시간 승인제 (2026.5 보도)"), c("속도·검증은 있으나 ‘같은 대본 비교’ 없음")],
    [c("Backstage (미국)", { bold: true }), c("유료 구독자 25만+, StarNow·Mandy 인수로 글로벌 통합"), c("배우 구독 부담 모델, 한국 미진출")],
    [c("CASTON · SCENEPORT", { bold: true }), c("AI 영상 분석·추천 / 셀프테이프·지원 관리"), c("AI 점수화에 대한 예술계 반감, 데이터 동의")],
  ];
  s.addTable(rows, { x: M, y: 1.75, w: CW, colW: [2.35, 6.0, 3.783], rowH: [0.48, 0.8, 0.8, 0.7, 0.7], border: { type: "solid", pt: 0.75, color: HEX.accent4 }, margin: [0.06, 0.15, 0.06, 0.15] });
  card(s, M, 5.65, CW, 1.0, C.accent6);
  T(s, [
    { text: "아무도 갖지 못한 3가지   ", options: { bold: true, color: C.accent2 } },
    { text: "① 같은 대본 비교 데이터   ② 연출자 평판·포트폴리오   ③ 레전드 배우 콘텐츠", options: { bold: true } },
  ], { x: M + 0.35, y: 5.65, w: CW - 0.7, h: 1.0, fontSize: 16, valign: "middle" });
}

// ═════════ 6. RED PEN on GPT draft
pres.addSection({ title: "원안 수정" });
{
  const s = L("LIGHT", "원안 수정");
  head(s, "INT. 편집실 — NIGHT", "GPT 원안에 빨간 펜을 댔습니다", "빨간 펜", "추론");
  const sc = [["GPT 원안", "68"], ["v2.0", "92"], ["v3.0 (본안)", "95"]];
  sc.forEach((x, i) => {
    const X = 9.2 + i * 1.22;
    T(s, x[1], { x: X, y: 0.75, w: 1.1, h: 0.55, fontSize: 28, fontFace: "Pretendard Black", align: "center", color: i === 2 ? C.accent2 : C.accent3 });
    T(s, x[0], { x: X - 0.1, y: 1.28, w: 1.3, h: 0.25, fontSize: 10, align: "center", color: C.accent3 });
  });
  const R = [
    ["‘연기자용 틴더’ 포지셔닝", "장면 기반 케미스트리 플랫폼", "사기·성범죄 이슈 속 ‘데이팅’ 연상은 역효과"],
    ["8개 아카이브 동시 구축", "MVP 3레이어 + 연출자 아카이브", "범위 축소, 양방향은 진짜로"],
    ["Actor Pro 구독 수익", "배우 평생 무료 · 제작 측 과금", "배우 다수가 저소득, 구독은 가입 장벽"],
    ["익명 주간 대본 챌린지", "LEGEND SCRIPT (나문희·심영은)", "WEUM만 가능한 복제 불가 콘텐츠"],
    ["“운영 원칙을 분리해야 한다”", "방화벽 4규칙 + 별도 법인", "선언이 아닌 실행 규칙"],
    ["(없음)", "시장규모 · 3개년 손익 · MVP 예산", "의사결정 가능한 숫자"],
  ];
  T(s, "원안", { x: M, y: 1.75, w: 3, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3 });
  T(s, "수정안", { x: 4.65, y: 1.75, w: 3, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent3 });
  T(s, "빨간 펜 메모", { x: 8.9, y: 1.75, w: 3, h: 0.3, fontSize: 11, fontFace: MONO, bold: true, color: C.accent2 });
  R.forEach((r, i) => {
    const y = 2.12 + i * 0.75;
    T(s, r[0], { x: M, y, w: 3.85, h: 0.6, fontSize: 14, color: C.accent3, strike: "sngStrike", valign: "middle" });
    T(s, "→", { x: 4.25, y, w: 0.35, h: 0.6, fontSize: 16, bold: true, color: C.accent2, valign: "middle" });
    T(s, r[1], { x: 4.65, y, w: 4.1, h: 0.6, fontSize: 15, bold: true, valign: "middle" });
    T(s, r[2], { x: 8.9, y, w: 3.83, h: 0.6, fontSize: 13, italic: true, color: C.accent2, valign: "middle" });
    if (i < 5) s.addShape(pres.ShapeType.line, { x: M, y: y + 0.68, w: CW, h: 0, line: { color: HEX.accent4, width: 0.75 } });
  });
}

// ═════════ 7. CONCEPT
pres.addSection({ title: "컨셉" });
{
  const s = L("LIGHT", "컨셉");
  head(s, "INT. 콘셉트 — DAY", "SCENE-FIRST: 사람보다 장면을 먼저 본다", "장면을 먼저 본다", "추론");
  const P = [
    ["SCENE", "같은 장면", "모든 배우가 같은 대본·같은 조건으로 연기한다. 경력·외모·소속사가 아니라 ‘해석’이 비교된다."],
    ["CHEMISTRY", "서로 고른다", "연출자도 포트폴리오와 현장 평판을 공개한다. 둘 다 SPARK를 눌러야 대화가 열린다."],
    ["SAFE", "안심하고 만난다", "검증된 연출자만, 앱 안에서만, 공개된 장소에서만. 신뢰가 곧 우리의 상품이다."],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  P.forEach((p, i) => {
    const X = M + i * (cw + 0.3);
    card(s, X, 1.75, cw, 4.9, i === 1 ? C.text1 : C.background2);
    const dark = i === 1;
    T(s, "0" + (i + 1), { x: X + 0.4, y: 2.05, w: 1, h: 0.35, fontSize: 13, fontFace: MONO, bold: true, color: dark ? C.accent1 : C.accent3 });
    const ww = measure(p[0], 34, "bk");
    s.addShape(pres.ShapeType.rect, dark
      ? { x: X + 0.3, y: 2.5, w: ww + 0.2, h: 0.68, fill: { color: C.accent1 }, line: { type: "none" } }
      : { x: X + 0.36, y: 2.82, w: ww + 0.08, h: 0.3, fill: { color: C.accent1 }, line: { type: "none" } });
    T(s, p[0], { x: X + 0.4, y: 2.5, w: cw - 0.8, h: 0.7, fontSize: 34, fontFace: "Pretendard Black", color: C.text1 });
    T(s, p[1], { x: X + 0.4, y: 3.45, w: cw - 0.8, h: 0.45, fontSize: 19, bold: true, color: dark ? C.background1 : C.text1 });
    T(s, p[2], { x: X + 0.4, y: 4.1, w: cw - 0.8, h: 2.0, fontSize: 14, color: dark ? C.accent4 : C.accent3, lineSpacingMultiple: 1.4 });
  });
}

// ═════════ 8. FLOW
{
  const s = L("LIGHT", "컨셉");
  head(s, "INT. 앱 안 — CONTINUOUS", "대본 한 장에서 작품 한 편까지, 6단계", "6단계", "추론");
  const st = [
    ["SCENE DROP", "연출자·레전드가\n대본을 공개"], ["TAKE", "배우가 30~60초\n해석 연기 업로드"], ["KEEP", "연출자가 장면을\n후보 보드에 저장"],
    ["SPARK", "배우가 연출자 평판을\n보고 수락하면 매칭"], ["MEET", "앱 내 대화 →\n공개 장소 오디션"], ["LEDGER", "작품 참여가\n검증 이력으로 기록"],
  ];
  const cw = 1.85, gap = (CW - 6 * cw) / 5;
  st.forEach((x, i) => {
    const X = M + i * (cw + gap), hot = i === 3;
    s.addShape(pres.ShapeType.ellipse, { x: X + cw / 2 - 0.55, y: 1.95, w: 1.1, h: 1.1, fill: { color: hot ? C.accent1 : C.text1 }, line: hot ? { color: HEX.dk1, width: 1.5 } : { type: "none" } });
    T(s, String(i + 1), { x: X + cw / 2 - 0.55, y: 1.95, w: 1.1, h: 1.1, fontSize: 30, fontFace: "Pretendard Black", color: hot ? C.text1 : C.background1, align: "center", valign: "middle" });
    T(s, x[0], { x: X - 0.1, y: 3.25, w: cw + 0.2, h: 0.4, fontSize: 16, fontFace: "Pretendard Black", align: "center" });
    T(s, x[1], { x: X - 0.1, y: 3.7, w: cw + 0.2, h: 0.8, fontSize: 12.5, align: "center", color: C.accent3, lineSpacingMultiple: 1.25 });
    if (i < 5) s.addShape(pres.ShapeType.line, { x: X + cw / 2 + 0.65, y: 2.5, w: cw + gap - 1.3, h: 0, line: { color: HEX.accent4, width: 1.5, endArrowType: "triangle" } });
  });
  const hw = (CW - 0.3) / 2;
  [["배우가 얻는 것", "프로필 사진·경력 없이 ‘연기 한 장면’으로 검증된 연출자에게 직접 도달한다. 거절권은 배우에게도 있다."],
   ["연출자가 얻는 것", "같은 대본으로 수백 명의 해석을 나란히 비교한다. 작품이 없을 때도 ‘언젠가 함께할 배우’ 풀을 쌓는다."]].forEach((x, i) => {
    const X = M + i * (hw + 0.3);
    card(s, X, 4.85, hw, 1.8);
    T(s, x[0], { x: X + 0.35, y: 5.1, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent2 });
    T(s, x[1], { x: X + 0.35, y: 5.5, w: hw - 0.7, h: 1.0, fontSize: 14, lineSpacingMultiple: 1.3 });
  });
}

// ═════════ 9. UX MOCKS
{
  const s = L("LIGHT", "컨셉");
  head(s, "INT. 스마트폰 화면 — DAY", "스와이프는 가볍게, 결정은 무겁게", "결정은 무겁게", "추론");
  const pw = 2.6, ph = 4.55, gap = (CW - 3 * pw) / 3, y0 = 1.72;
  const xs = [0, 1, 2].map((i) => M + gap / 2 + i * (pw + gap));
  xs.forEach((x) => s.addShape(pres.ShapeType.roundRect, { x, y: y0, w: pw, h: ph, rectRadius: 0.28, fill: { color: C.text1 }, line: { type: "none" } }));
  const inner = (x, y, w, h, fill) => s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: { type: "none" } });
  // phone 1: scene feed
  let x = xs[0];
  T(s, "SCENE FEED", { x: x + 0.25, y: y0 + 0.25, w: 2, h: 0.25, fontSize: 10, fontFace: MONO, bold: true, color: C.accent1 });
  inner(x + 0.18, y0 + 0.6, pw - 0.36, 2.05, "3A3A3A");
  T(s, "▶  EP.001 · 30s", { x: x + 0.32, y: y0 + 0.72, w: 2, h: 0.25, fontSize: 10, bold: true, color: C.background1 });
  T(s, "“미안해. 사실 나\n그날 거기 있었어.”", { x: x + 0.3, y: y0 + 1.3, w: pw - 0.6, h: 0.8, fontSize: 13, bold: true, align: "center", color: C.background1 });
  T(s, "배우 A · 해석: 죄책감", { x: x + 0.25, y: y0 + 2.8, w: pw - 0.5, h: 0.3, fontSize: 12.5, bold: true, color: C.background1 });
  T(s, "#절제  #생활연기  #스릴러", { x: x + 0.25, y: y0 + 3.13, w: pw - 0.5, h: 0.25, fontSize: 10.5, color: C.accent4 });
  [["PASS", "5A5A5A", "FFFFFF"], ["KEEP", "FFFFFF", "141414"], ["SPARK", HEX.accent1, "141414"]].forEach((b, i) => {
    s.addText(b[0], { x: x + 0.2 + i * 0.75, y: y0 + 3.7, w: 0.68, h: 0.4, shape: pres.ShapeType.roundRect, rectRadius: 0.2, isTextBox: true, fill: { color: b[1] }, color: b[2], fontSize: 9.5, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  // phone 2: director profile
  x = xs[1];
  T(s, "DIRECTOR", { x: x + 0.25, y: y0 + 0.25, w: 2, h: 0.25, fontSize: 10, fontFace: MONO, bold: true, color: C.accent1 });
  s.addShape(pres.ShapeType.ellipse, { x: x + 0.25, y: y0 + 0.65, w: 0.7, h: 0.7, fill: { color: "5A5A5A" }, line: { type: "none" } });
  T(s, "감독 B", { x: x + 1.08, y: y0 + 0.68, w: 1.4, h: 0.35, fontSize: 15, bold: true, color: C.background1 });
  s.addText("✓ VERIFIED", { x: x + 1.08, y: y0 + 1.05, w: 1.05, h: 0.26, shape: pres.ShapeType.roundRect, rectRadius: 0.13, isTextBox: true, fill: { color: C.accent1 }, color: C.text1, fontSize: 9, bold: true, align: "center", valign: "middle", margin: 0 });
  T(s, "차기작 · 심리 드라마\n#인간의 내면  #작가주의", { x: x + 0.25, y: y0 + 1.55, w: pw - 0.5, h: 0.6, fontSize: 11, color: C.accent4, lineSpacingMultiple: 1.25 });
  inner(x + 0.18, y0 + 2.3, pw - 0.36, 1.55, "2C2C2C");
  T(s, "SET REVIEW  ★ 4.8 (12)", { x: x + 0.32, y: y0 + 2.42, w: pw - 0.6, h: 0.28, fontSize: 11, bold: true, color: C.accent1 });
  T(s, "“리허설 시간을 충분히 줬고,\n디렉션이 명확했다.”\n— 함께한 배우 (검증 매칭)", { x: x + 0.32, y: y0 + 2.78, w: pw - 0.6, h: 1.0, fontSize: 10.5, color: C.background1, lineSpacingMultiple: 1.25 });
  T(s, "SPARK 보내기", { x: x + 0.25, y: y0 + 3.98, w: pw - 0.5, h: 0.4, fontSize: 11, bold: true, color: C.text1, align: "center", valign: "middle", shape: pres.ShapeType.roundRect, rectRadius: 0.2, fill: { color: C.accent1 } });
  // phone 3: match + safe meet
  x = xs[2];
  T(s, "MATCH", { x: x + 0.25, y: y0 + 0.25, w: 2, h: 0.25, fontSize: 10, fontFace: MONO, bold: true, color: C.accent1 });
  inner(x + 0.18, y0 + 0.6, pw - 0.36, 1.15, HEX.accent1);
  T(s, "IT’S A SCENE!", { x: x + 0.18, y: y0 + 0.72, w: pw - 0.36, h: 0.5, fontSize: 20, fontFace: "Pretendard Black", align: "center", color: C.text1 });
  T(s, "배우 A × 감독 B", { x: x + 0.18, y: y0 + 1.25, w: pw - 0.36, h: 0.3, fontSize: 11, bold: true, align: "center", color: C.text1 });
  inner(x + 0.18, y0 + 1.95, pw - 0.36, 2.15, "2C2C2C");
  T(s, "SAFE MEET", { x: x + 0.32, y: y0 + 2.07, w: 2, h: 0.28, fontSize: 11, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, bullets(["장소: 등록된 사무실", "시간: 평일 14:00 (주간)", "대화 기록 보존 ON", "금전 요구 자동 차단"]), { x: x + 0.32, y: y0 + 2.45, w: pw - 0.6, h: 1.5, fontSize: 11, color: C.background1, paraSpaceAfter: 4 });
  ["① 장면으로 발견", "② 연출자 평판을 보고 수락", "③ 안전 규칙 안에서 만남"].forEach((t, i) => {
    T(s, t, { x: xs[i] - 0.2, y: 6.38, w: pw + 0.4, h: 0.32, fontSize: 13, bold: true, align: "center" });
  });
}

// ═════════ 10. SAFE CAST
{
  const s = L("LIGHT", "컨셉");
  head(s, "INT. 안전 관제실 — 24H", "안전은 기능이 아니라 상품이다", "상품", "추론");
  card(s, M, 1.75, 3.6, 4.9, C.text1);
  T(s, "WHY", { x: M + 0.35, y: 2.05, w: 2, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "366건", { x: M + 0.35, y: 2.45, w: 3, h: 0.8, fontSize: 44, fontFace: "Pretendard Black", color: C.background1 });
  T(s, "학원형 연예기획사 피해구제\n(2023 ~ 2026.6)\n\n환급 거부가 약 절반,\n‘당일 합격 통보’로 계약 유도", { x: M + 0.35, y: 3.35, w: 2.9, h: 2.2, fontSize: 13.5, color: C.accent4, lineSpacingMultiple: 1.3 });
  tag(s, M + 0.35, 6.1, "사실");
  T(s, "SAFE CAST 5원칙", { x: 4.6, y: 1.78, w: 6, h: 0.45, fontSize: 20, fontFace: "Pretendard Black" });
  const r = [
    ["연출자 3중 인증", "실명 + 사업자·소속 + 필모그래피 확인. 48시간 심사 후 VERIFIED 배지"],
    ["매칭 전 DM 불가", "양쪽이 SPARK한 뒤에만 대화. 대화는 앱 안에서만, 기록 보존"],
    ["공개 장소 · 주간 원칙", "오디션 장소·시간을 앱에 등록. ‘밀실 오디션’ 신고 시 즉시 계정 정지"],
    ["금전 요구 원천 차단", "‘트레이닝비·프로필비·등록비’ 키워드 감지 → 경고 후 차단"],
    ["미성년자 보호 모드", "만 19세 미만은 보호자 연동 계정, 모든 대화에 보호자 자동 참여"],
  ];
  r.forEach((x, i) => {
    const y = 2.4 + i * 0.86;
    num(s, i + 1, 4.6, y + 0.03, 0.44, C.accent1, C.text1);
    T(s, x[0], { x: 5.25, y, w: 7.4, h: 0.35, fontSize: 15.5, bold: true });
    T(s, x[1], { x: 5.25, y: y + 0.38, w: 7.4, h: 0.35, fontSize: 13, color: C.accent3 });
  });
}

// ═════════ 11. LEGEND SCRIPT (dark)
pres.addSection({ title: "킬러 콘텐츠" });
{
  const s = L("DARK", "킬러 콘텐츠");
  head(s, "INT. 촬영장 — NIGHT", "같은 대사, 천 개의 얼굴 — LEGEND SCRIPT", "천 개의 얼굴", "추론", true);
  card(s, M, 1.75, 5.5, 4.9, C.background1);
  T(s, "EP.001  ·  LEGEND SCRIPT", { x: M + 0.4, y: 2.05, w: 4.7, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent3 });
  T(s, "마지막 전화", { x: M + 0.4, y: 2.45, w: 4.7, h: 0.7, fontSize: 32, fontFace: "Pretendard Black" });
  const q = "“미안해. 사실 나 그날 거기 있었어.”";
  s.addShape(pres.ShapeType.rect, { x: M + 0.36, y: 3.42, w: measure(q, 17) + 0.08, h: 0.2, fill: { color: C.accent1 }, line: { type: "none" } });
  T(s, q, { x: M + 0.4, y: 3.25, w: 4.8, h: 0.45, fontSize: 17, bold: true });
  T(s, [
    { text: "시즌 1 호스트", options: { bold: true, breakLine: true } },
    { text: "나문희 배우 — 대본 선정 · 30초 시범 연기 (본인 동의 전제)", options: { color: C.accent3, breakLine: true } },
    { text: " ", options: { fontSize: 6, breakLine: true } },
    { text: "EP.000 파일럿 · 릴레이 호스트", options: { bold: true, breakLine: true } },
    { text: "심영은 배우 → 협업 감독·작가 순환", options: { color: C.accent3 } },
  ], { x: M + 0.4, y: 4.05, w: 4.7, h: 2.3, fontSize: 13.5, lineSpacingMultiple: 1.3 });
  const st = [
    ["WEEK 1", "레전드가 고른 3줄 대본 + 시범 연기 공개 → 숏폼 바이럴 원천"],
    ["WEEK 2", "신인 누구나 같은 대본으로 TAKE 업로드 · 비용 0원"],
    ["WEEK 3", "레전드가 TOP 10 장면에 손편지 코멘트 → 최고의 PR 자산"],
    ["WEEK 4", "TOP 3 → ACT ORIGINALS 단편 출연 확정 · 촬영기 공개"],
  ];
  st.forEach((x, i) => {
    const y = 1.8 + i * 0.95;
    T(s, x[0], { x: 6.5, y, w: 1.2, h: 0.35, fontSize: 13, fontFace: MONO, bold: true, color: C.accent1 });
    T(s, x[1], { x: 7.75, y, w: 4.98, h: 0.75, fontSize: 14.5, color: C.background1, lineSpacingMultiple: 1.25 });
  });
  card(s, 6.5, 5.75, 6.23, 0.9, C.accent1);
  T(s, "근거(사실): 플필 신인 오디션은 ‘레전드’ 없이도 한 달 만에 3,000명+ 지원 (2024.8)", { x: 6.75, y: 5.75, w: 5.8, h: 0.9, fontSize: 13, bold: true, valign: "middle" });
}

// ═════════ 12. ARCHIVE
{
  const s = L("LIGHT", "킬러 콘텐츠");
  head(s, "INT. 아카이브 — DAY", "배우만 기록되는 플랫폼은 반쪽이다", "반쪽", "추론");
  const hw = (CW - 0.3) / 2;
  const sides = [
    ["ACTOR ARCHIVE", "MVP 3레이어", [["SCENE REEL", "챌린지 대본 연기 — 비교 가능한 표준 데이터"], ["FREE REEL", "독백·자유연기·발성 — 개성의 기록"], ["CAREER LEDGER", "매칭·오디션·출연 이력 — 플랫폼 검증 스탬프"]], "Phase 2: Voice · Growth(연도별 비교) 추가"],
    ["DIRECTOR ARCHIVE", "원안에 없던 것", [["DIRECTOR REEL", "연출작·차기작 톤·현장 철학 30초 영상"], ["SET REVIEW", "함께 작업한 배우의 익명 현장 후기 (검증 매칭만)"], ["CAST HISTORY", "이 감독이 발굴한 신인들 — 연출자의 ‘안목’이 브랜드"]], "배우가 연출자를 고를 근거가 생겨야 ‘양방향’이 진짜가 된다"],
  ];
  sides.forEach((sd, i) => {
    const X = M + i * (hw + 0.3), dark = i === 1;
    card(s, X, 1.75, hw, 4.9, dark ? C.text1 : C.background2);
    T(s, sd[0], { x: X + 0.4, y: 2.0, w: 3.5, h: 0.4, fontSize: 18, fontFace: "Pretendard Black", color: dark ? C.background1 : C.text1 });
    s.addText(sd[1], { x: X + hw - 2.1, y: 2.03, w: 1.7, h: 0.32, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: C.accent1 }, color: C.text1, fontSize: 11, bold: true, align: "center", valign: "middle", margin: 0 });
    sd[2].forEach((r, k) => {
      const y = 2.65 + k * 1.0;
      T(s, r[0], { x: X + 0.4, y, w: hw - 0.8, h: 0.38, fontSize: 17, bold: true, color: dark ? C.background1 : C.text1 });
      T(s, r[1], { x: X + 0.4, y: y + 0.42, w: hw - 0.8, h: 0.38, fontSize: 13.5, color: dark ? C.accent4 : C.accent3 });
    });
    T(s, sd[3], { x: X + 0.4, y: 5.85, w: hw - 0.8, h: 0.5, fontSize: 12.5, bold: dark, italic: !dark, color: dark ? C.accent1 : C.accent3 });
  });
}

// ═════════ 13. AI CHARTER
{
  const s = L("LIGHT", "킬러 콘텐츠");
  head(s, "INT. 서버실 — NIGHT", "AI는 찾아주고, 판단은 사람이 한다", "판단은 사람이", "추론");
  const cols = [
    ["AI가 하는 일", C.text1, ["장면 태그 자동 생성 (감정·톤·장르·발성)", "자연어 검색: “눈빛으로 불안을 절제하는 40대”", "비슷한 해석 묶어 보기 · 연출자 취향 학습"]],
    ["AI가 하지 않는 일", C.accent2, ["연기력 점수화 · 순위 매기기", "외모 기반 추천 · 랭킹", "동의 없는 얼굴·목소리 학습, 디지털 복제"]],
    ["배우 데이터 헌장", C.accent1, ["영상 소유권은 100% 배우에게", "AI 학습은 별도 opt-in, 작품별 개별 동의", "탈퇴 시 30일 내 영상·파생데이터 삭제"]],
  ];
  const cw = (CW - 2 * 0.3) / 3;
  cols.forEach((c, i) => {
    const X = M + i * (cw + 0.3);
    card(s, X, 1.75, cw, 3.8);
    s.addShape(pres.ShapeType.ellipse, { x: X + 0.35, y: 2.05, w: 0.46, h: 0.46, fill: { color: c[1] }, line: { color: HEX.dk1, width: 0.75 } });
    T(s, c[0], { x: X + 0.95, y: 2.08, w: cw - 1.2, h: 0.4, fontSize: 18, fontFace: "Pretendard Black" });
    T(s, bullets(c[2]), { x: X + 0.35, y: 2.85, w: cw - 0.7, h: 2.5, fontSize: 14, paraSpaceAfter: 10, lineSpacingMultiple: 1.15 });
  });
  card(s, M, 5.8, CW, 0.85, C.accent6);
  tag(s, M + 0.35, 6.08, "사실");
  T(s, "SAG-AFTRA(미국 배우조합) 협약: 디지털 복제는 사전 동의·용도 고지·보상 필수, 포괄적 동의 불가 → 이 기준을 플랫폼 약관에 그대로 넣는다", { x: M + 1.15, y: 5.8, w: CW - 1.45, h: 0.85, fontSize: 13.5, valign: "middle" });
}

// ═════════ 14. MARKET SIZE
pres.addSection({ title: "사업성" });
{
  const s = L("LIGHT", "사업성");
  head(s, "INT. 재무 회의실 — DAY", "시장 규모: 작게 시작해서 확실히 이긴다", "확실히 이긴다", "추론");
  // concentric circles
  const cx = M + 2.55, cy = 4.2;
  [[2.45, C.background2], [1.65, C.accent4], [0.9, C.accent1]].forEach((c) => s.addShape(pres.ShapeType.ellipse, { x: cx - c[0], y: cy - c[0], w: c[0] * 2, h: c[0] * 2, fill: { color: c[1] }, line: { type: "none" } }));
  T(s, "TAM  235억", { x: cx - 1.5, y: cy - 2.2, w: 3, h: 0.35, fontSize: 14, bold: true, align: "center" });
  T(s, "SAM  70억", { x: cx - 1.5, y: cy - 1.45, w: 3, h: 0.35, fontSize: 14, bold: true, align: "center" });
  T(s, "SOM\n18억", { x: cx - 0.9, y: cy - 0.45, w: 1.8, h: 0.9, fontSize: 16, fontFace: "Pretendard Black", align: "center", valign: "middle" });
  const rows = [
    ["TAM", "사실 + 추론", "국내 대중문화예술 기획업 4,471곳 + 제작업 2,289곳 = 6,760곳 (2024년 기준) × Studio Pro 연 348만 원", "≈ 235억/년"],
    ["SAM", "추론", "그중 상시 캐스팅 수요가 있는 30% ≈ 2,000곳 × 연 348만 원", "≈ 70억/년"],
    ["SOM", "추론", "3년차 350곳(전체의 5.2%) 유료 12.2억 + 브랜드 챌린지 12건 3.6억 + Plus·Originals 2.2억", "≈ 18억/년"],
  ];
  rows.forEach((r, i) => {
    const y = 1.85 + i * 1.42, X = 5.9;
    card(s, X, y, CW - 5.3, 1.25, i === 2 ? C.accent6 : C.background2);
    T(s, r[0], { x: X + 0.3, y: y + 0.18, w: 1.0, h: 0.4, fontSize: 18, fontFace: "Pretendard Black" });
    T(s, r[1], { x: X + 0.3, y: y + 0.62, w: 1.2, h: 0.3, fontSize: 10.5, fontFace: MONO, color: C.accent3 });
    T(s, r[2], { x: X + 1.55, y: y + 0.1, w: 3.85, h: 1.05, fontSize: 12.5, valign: "middle", lineSpacingMultiple: 1.2 });
    T(s, r[3], { x: X + 5.45, y: y + 0.1, w: 1.35, h: 1.05, fontSize: 15, bold: true, align: "right", valign: "middle", color: i === 2 ? C.accent2 : C.text1 });
  });
  T(s, "확장: 글로벌 Backstage 유료 구독자 25만+ (사실) → 해외 캐스팅은 3단계", { x: 5.9, y: 6.2, w: CW - 5.3, h: 0.4, fontSize: 12, italic: true, color: C.accent3 });
}

// ═════════ 15. REVENUE MODEL
{
  const s = L("LIGHT", "사업성");
  head(s, "INT. 가격표 — DAY", "배우는 평생 무료, 시간을 아끼는 쪽이 지불한다", "평생 무료", "추론");
  const rows = [
    ["Actor", "무료", "업로드·챌린지·매칭 전부 무료  (Plus 월 9,900원: 조회 통계·추가 저장)"],
    ["Director", "무료", "검증 연출자 월 5회 SPARK, 후보 보드 1개"],
    ["Studio Pro", "월 29만 원~", "팀 후보 보드 · NDA 대본 배포 · 셀프테이프 수집 · 무제한 SPARK"],
    ["Brand Challenge", "회당 3,000만 원~", "브랜드 대본 챌린지 → 광고 모델 발굴 (WEUM 광고 기획 역량)"],
    ["ACT ORIGINALS", "IP 수익", "챌린지 TOP 배우 출연 단편·숏드라마 유통·판권"],
  ];
  rows.forEach((r, i) => {
    const y = 1.75 + i * 0.9, hot = i === 2 || i === 3;
    card(s, M, y, CW, 0.78, hot ? C.accent6 : C.background2);
    T(s, r[0], { x: M + 0.35, y, w: 2.6, h: 0.78, fontSize: 16, bold: true, valign: "middle" });
    T(s, r[1], { x: M + 3.0, y, w: 2.4, h: 0.78, fontSize: 15, bold: true, color: hot ? C.accent2 : C.text1, valign: "middle" });
    T(s, r[2], { x: M + 5.4, y, w: CW - 5.7, h: 0.78, fontSize: 13.5, color: C.accent3, valign: "middle" });
  });
  T(s, "원칙: 초기 캐스팅 성사 수수료 0원 — 직업소개 규제 리스크를 피하고 배우 신뢰를 지킨다", { x: M, y: 6.35, w: CW, h: 0.32, fontSize: 13, bold: true, color: C.accent2 });
}

// ═════════ 16. 3-YEAR P&L
{
  const s = L("LIGHT", "사업성");
  head(s, "INT. 재무 회의실 — NIGHT", "3년차 손익분기 — 숫자로 본 가설", "3년차 손익분기", "추론");
  s.addChart(pres.ChartType.bar, [
    { name: "매출", labels: ["1년차", "2년차", "3년차"], values: [1.0, 6.0, 18.0] },
    { name: "비용", labels: ["1년차", "2년차", "3년차"], values: [6.5, 9.0, 14.0] },
  ], {
    x: M, y: 1.7, w: 6.6, h: 4.95, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.accent1, HEX.dk1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0.0",
    showTitle: true, title: "매출 vs 비용 (억 원)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    catAxisLabelColor: HEX.accent3, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.accent3,
  });
  const X = 7.6, w = CW - 7.0;
  const tbl = [
    ["", "1년차", "2년차", "3년차"],
    ["Studio Pro 유료사", "20곳", "120곳", "350곳"],
    ["브랜드 챌린지", "2건", "4건", "12건"],
    ["매출", "1.0억", "6.0억", "18.0억"],
    ["영업손익", "−5.5억", "−3.0억", "+4.0억"],
  ];
  s.addTable(tbl.map((r, i) => r.map((c, j) => ({ text: c, options: { fontSize: 12.5, bold: i === 0 || j === 0 || i === 4, align: j ? "center" : "left", valign: "middle", color: i === 4 && j === 3 ? C.accent2 : C.text1, fill: { color: i === 0 ? HEX.accent1 : (i % 2 ? "FFFFFF" : HEX.lt2) } } }))),
    { x: X, y: 1.8, w, colW: [1.95, 0.95, 0.95, w - 3.85], rowH: 0.45, border: { type: "solid", pt: 0.75, color: HEX.accent4 }, margin: [0.04, 0.1, 0.04, 0.1] });
  T(s, "핵심 가정", { x: X, y: 4.2, w, h: 0.3, fontSize: 13, bold: true });
  T(s, bullets(["Studio Pro 월 29만 원, 해지율 월 3% 이하", "비용: 1년차 개발·콘텐츠 집중, 이후 인력·마케팅 중심", "브랜드 챌린지 회당 3,000만 원 (WEUM 광고 네트워크 활용)"]), { x: X, y: 4.55, w, h: 1.7, fontSize: 12.5, paraSpaceAfter: 6, color: C.accent3 });
}

// ═════════ 17. MVP BUDGET
{
  const s = L("LIGHT", "사업성");
  head(s, "INT. 예산 회의 — DAY", "MVP 6개월, 약 5억 원이면 증명할 수 있다", "약 5억 원", "추론");
  s.addChart(pres.ChartType.bar, [{ name: "예산(억 원)", labels: ["앱·웹 개발 (외주)", "운영 인력 3명 × 6개월", "콘텐츠 (LEGEND·단편)", "런칭 마케팅", "법률·상표·보안"], values: [2.2, 1.2, 0.8, 0.5, 0.3] }], {
    x: M, y: 1.7, w: 7.3, h: 4.95, barDir: "bar", chartColors: [HEX.dk1], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "0.0",
    showTitle: true, title: "MVP 예산 배분 (억 원, 합계 5.0)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  });
  const X = 8.3, w = CW - 7.7;
  card(s, X, 1.8, w, 2.4, C.text1);
  T(s, "재원 조달 옵션", { x: X + 0.35, y: 2.0, w: w - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  T(s, bullets(["WEUM 자체 투자", "콘진원·창업 정부지원사업 (매칭 펀드)", "엔젤·시드 (M6 성과 리포트 후)"]), { x: X + 0.35, y: 2.45, w: w - 0.7, h: 1.6, fontSize: 13.5, color: C.background1, paraSpaceAfter: 8 });
  card(s, X, 4.4, w, 2.25);
  T(s, "비용을 줄이는 3가지 장치", { x: X + 0.35, y: 4.6, w: w - 0.7, h: 0.35, fontSize: 15, bold: true });
  T(s, bullets(["AI 고도화는 데이터 축적 후 (MVP 제외)", "영상 저장은 60초 제한 + 압축 스트리밍", "단편 제작은 WEUM 제작 인력 공유"]), { x: X + 0.35, y: 5.05, w: w - 0.7, h: 1.5, fontSize: 13, color: C.accent3, paraSpaceAfter: 6 });
}

// ═════════ 18. GTM 90 DAYS
pres.addSection({ title: "실행" });
{
  const s = L("LIGHT", "실행");
  head(s, "EXT. 숏폼 피드 — 런칭 90일", "런칭 90일, 화제는 우연이 아니라 설계다", "설계다", "추론");
  const beats = [
    ["D-30", "티저", "“지금도 누군가, 당신의 장면을\n다시 보고 있다.”", "대본 넘기는 손 · 실루엣만 공개"],
    ["D-DAY", "EP.001 공개", "레전드의 30초 시범 연기", "같은 대본 챌린지 오픈"],
    ["D+7", "천 개의 얼굴", "같은 대사 100명 몽타주", "쇼츠·릴스·틱톡 동시"],
    ["D+21", "손편지 공개", "레전드가 TOP 10에 쓴 코멘트", "언론 보도자료 배포"],
    ["D+60", "촬영기", "TOP 3의 단편 촬영 다큐", "‘발견→작품’ 증명"],
    ["D+90", "리포트", "매칭·오디션 성과 공개", "B2B 세일즈 자료로 전환"],
  ];
  const cw = (CW - 5 * 0.2) / 6;
  s.addShape(pres.ShapeType.line, { x: M, y: 2.12, w: CW, h: 0, line: { color: HEX.dk1, width: 1.5 } });
  beats.forEach((b, i) => {
    const X = M + i * (cw + 0.2), hot = i === 1 || i === 3;
    s.addShape(pres.ShapeType.ellipse, { x: X, y: 1.98, w: 0.28, h: 0.28, fill: { color: hot ? C.accent1 : C.text1 }, line: { color: HEX.dk1, width: 1 } });
    T(s, b[0], { x: X, y: 2.42, w: cw, h: 0.35, fontSize: 15, fontFace: MONO, bold: true, color: hot ? C.accent2 : C.text1 });
    card(s, X, 2.85, cw, 2.6, hot ? C.accent6 : C.background2);
    T(s, b[1], { x: X + 0.18, y: 3.0, w: cw - 0.36, h: 0.35, fontSize: 15, bold: true });
    T(s, b[2], { x: X + 0.18, y: 3.42, w: cw - 0.36, h: 1.1, fontSize: 12, lineSpacingMultiple: 1.25 });
    T(s, b[3], { x: X + 0.18, y: 4.6, w: cw - 0.36, h: 0.75, fontSize: 11, color: C.accent3, lineSpacingMultiple: 1.2 });
  });
  card(s, M, 5.7, CW, 0.95, C.text1);
  T(s, [
    { text: "채널  ", options: { bold: true, color: C.accent1 } },
    { text: "인스타·유튜브 쇼츠·틱톡  ·  연극영화과 커뮤니티 앰배서더  ·  독립영화제 현장 부스  ·  WEUM 소속 배우 SNS", options: { color: C.background1 } },
  ], { x: M + 0.35, y: 5.7, w: CW - 0.7, h: 0.95, fontSize: 14, valign: "middle" });
}

// ═════════ 19. ROADMAP + GATE
{
  const s = L("LIGHT", "실행");
  head(s, "INT. 작전실 — DAY", "6개월 뒤, 숫자로 Go / No-Go를 결정한다", "Go / No-Go", "추론");
  const ms = [["M1", "SEED", "독립영화 감독 30명 사전 합류"], ["M2", "BETA", "배우 300명 초대 · EP.000 파일럿"], ["M3", "OPEN", "주간 챌린지 · SAFE CAST 가동"], ["M4", "FILM", "TOP 3 배우 단편 1편 촬영"], ["M5", "B2B", "Studio Pro 유료 파일럿 3곳"], ["M6", "GATE", "성과 리포트 · 투자 판단"]];
  const cw = (CW - 5 * 0.15) / 6;
  ms.forEach((m, i) => {
    const X = M + i * (cw + 0.15), last = i === 5;
    card(s, X, 1.75, cw, 1.85, last ? C.accent1 : C.background2, last);
    T(s, m[0], { x: X + 0.2, y: 1.9, w: 1, h: 0.45, fontSize: 22, fontFace: "Pretendard Black", color: last ? C.text1 : C.accent3 });
    T(s, m[1], { x: X + 0.2, y: 2.38, w: cw - 0.4, h: 0.3, fontSize: 12, fontFace: MONO, bold: true });
    T(s, m[2], { x: X + 0.2, y: 2.72, w: cw - 0.4, h: 0.8, fontSize: 12, color: last ? C.text1 : C.accent3, lineSpacingMultiple: 1.2 });
  });
  const hw = (CW - 0.3) / 2;
  card(s, M, 3.85, hw, 2.8, C.text1);
  T(s, "GO — 4개 중 3개 이상 충족 시 시드 투자 유치", { x: M + 0.35, y: 4.05, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  const g = [["매칭→오디션 전환", "20건"], ["연출자 주간 활성률", "40%"], ["배우 재업로드율", "30%"], ["유료 B2B", "3곳"]];
  g.forEach((x, i) => {
    const y = 4.55 + i * 0.5;
    T(s, x[0], { x: M + 0.35, y, w: 3.8, h: 0.4, fontSize: 14, color: C.background1, valign: "middle" });
    T(s, x[1], { x: M + hw - 1.7, y, w: 1.35, h: 0.4, fontSize: 17, fontFace: "Pretendard Black", align: "right", color: C.accent1, valign: "middle" });
  });
  const X2 = M + hw + 0.3;
  card(s, X2, 3.85, hw, 2.8, C.background1, true);
  T(s, "NO-GO — 손실을 5억 안에서 멈춘다", { x: X2 + 0.35, y: 4.05, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent2 });
  T(s, bullets(["플랫폼을 WEUM 내부 캐스팅 툴로 축소 전환", "LEGEND SCRIPT는 WEUM 유튜브 콘텐츠로 유지", "축적된 배우 풀은 ACT LAB 워크숍 모객에 활용"]), { x: X2 + 0.35, y: 4.55, w: hw - 0.7, h: 1.9, fontSize: 14, paraSpaceAfter: 8 });
}

// ═════════ 20. MOAT
{
  const s = L("LIGHT", "실행");
  head(s, "INT. 경쟁사 회의실 — DAY", "플필이 내일 그대로 따라 한다면?", "그대로 따라 한다면?", "추론");
  const m = [
    ["레전드 배우 콘텐츠", "나문희·심영은 배우는 WEUM에 있다. 기능은 복제돼도 사람은 복제되지 않는다.", "복제 불가"],
    ["연출자 평판 데이터", "Set Review는 실제 매칭·작업 후에만 쌓인다. 시간이 곧 해자다.", "시간 해자"],
    ["같은 대본 비교 데이터", "시즌이 쌓일수록 ‘한 대본, 천 개의 해석’ 데이터베이스는 커진다.", "데이터 해자"],
    ["제작 파이프라인", "매칭이 ACT ORIGINALS 작품으로 끝난다. 플랫폼 단독 사업자는 못 한다.", "수직 통합"],
  ];
  const hw = (CW - 0.3) / 2;
  m.forEach((x, i) => {
    const X = M + (i % 2) * (hw + 0.3), y = 1.75 + Math.floor(i / 2) * 2.0;
    card(s, X, y, hw, 1.8);
    num(s, i + 1, X + 0.35, y + 0.3, 0.5);
    T(s, x[0], { x: X + 1.05, y: y + 0.3, w: hw - 3.1, h: 0.45, fontSize: 18, bold: true });
    s.addText(x[2], { x: X + hw - 1.75, y: y + 0.33, w: 1.4, h: 0.34, shape: pres.ShapeType.rect, isTextBox: true, fill: { color: C.accent1 }, color: C.text1, fontSize: 11.5, bold: true, align: "center", valign: "middle", margin: 0 });
    T(s, x[1], { x: X + 1.05, y: y + 0.85, w: hw - 1.4, h: 0.8, fontSize: 13.5, color: C.accent3, lineSpacingMultiple: 1.3 });
  });
  T(s, "결론: 기능 경쟁이 아니라 ‘사람·시간·작품’ 경쟁으로 판을 옮긴다", { x: M, y: 5.85, w: CW, h: 0.5, fontSize: 17, bold: true, color: C.accent2 });
}

// ═════════ 21. WEUM ECOSYSTEM + FIREWALL
{
  const s = L("LIGHT", "실행");
  head(s, "INT. WEUM 본사 — DAY", "플랫폼은 모두에게 열고, 매니지먼트는 방화벽 뒤에 둔다", "방화벽 뒤에", "추론");
  const eco = [["ACT MATCH", "유입·아카이브"], ["ACT LAB", "워크숍·인큐베이팅"], ["WEUM MGMT", "선별 매니지먼트"], ["ACT ORIGINALS", "단편·숏드라마"], ["CONTENT IP", "유통·판권·광고"]];
  const cw = 2.2, gap = (CW - 5 * cw) / 4;
  eco.forEach((e, i) => {
    const X = M + i * (cw + gap), first = i === 0;
    card(s, X, 1.75, cw, 1.35, first ? C.accent1 : C.text1, first);
    T(s, e[0], { x: X + 0.1, y: 1.93, w: cw - 0.2, h: 0.42, fontSize: 15.5, fontFace: "Pretendard Black", align: "center", color: first ? C.text1 : C.background1 });
    T(s, e[1], { x: X + 0.1, y: 2.45, w: cw - 0.2, h: 0.35, fontSize: 12.5, align: "center", color: first ? C.text1 : C.accent4 });
    if (i < 4) T(s, "›", { x: X + cw, y: 2.1, w: gap, h: 0.6, fontSize: 24, bold: true, color: C.accent3, align: "center", valign: "middle" });
  });
  T(s, "FIREWALL 4규칙 — ‘분리해야 한다’를 ‘이렇게 분리한다’로", { x: M, y: 3.4, w: CW, h: 0.4, fontSize: 17, fontFace: "Pretendard Black" });
  const fw = [["노출 가중치 0", "추천·검색 알고리즘에 소속사 변수 미사용, 코드 수준에서 명시"], ["동일 조건 스카우트", "WEUM 스카우트도 일반 연출자와 같은 인증·같은 SPARK 한도"], ["7일 숙려 고지", "전속 제안 시 ‘거절해도 플랫폼 불이익 없음’ 고지 + 7일 숙려"], ["연 1회 외부 점검", "추천 공정성 리포트를 외부 검토 후 공개"]];
  const hw = (CW - 0.3) / 2;
  fw.forEach((f, i) => {
    const X = M + (i % 2) * (hw + 0.3), y = 3.95 + Math.floor(i / 2) * 1.37;
    card(s, X, y, hw, 1.2);
    T(s, f[0], { x: X + 0.35, y: y + 0.2, w: hw - 0.7, h: 0.38, fontSize: 15.5, bold: true });
    T(s, f[1], { x: X + 0.35, y: y + 0.63, w: hw - 0.7, h: 0.4, fontSize: 13, color: C.accent3 });
  });
}

// ═════════ 22. LEGAL
{
  const s = L("LIGHT", "실행");
  head(s, "INT. 법무 검토 — DAY", "출시 전에 반드시 막아야 할 4개의 리스크", "4개의 리스크");
  const r = [
    ["대중문화예술산업발전법", "사실", "기획업 미등록 시 2년 이하 징역 또는 2천만 원 이하 벌금 (보도 기준). 문체부, 일제 등록 계도 후 엄정 조치 방침", "WEUM 기획업 등록 상태 점검 · 플랫폼 법인은 ‘정보 매칭’으로 한정"],
    ["캐스팅 알선 수수료", "추론", "출연 성사 수수료 도입 시 유료직업소개 등 별도 규제 해당 가능성", "초기 수수료 0원 · 구독/SaaS만 과금 · 전문 변호사 사전 자문"],
    ["영상·음성 개인정보", "추론", "얼굴·목소리 영상은 개인정보. AI 처리 시 민감도 상승", "처리방침 · AI 학습 별도 동의 · 공개 범위 3단계"],
    ["상표 · 도메인", "추론", "ACT MATCH 유사 상표·도메인 선점 가능성", "KIPRIS 검색 → 9·35·41·42류 출원, 도메인 선확보"],
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

// ═════════ 23. BRAND
pres.addSection({ title: "브랜딩" });
{
  const s = L("LIGHT", "브랜딩");
  head(s, "INT. 디자인 스튜디오 — DAY", "두 장면이 겹치는 곳에 형광펜을 긋는다", "형광펜을 긋는다", "추론");
  card(s, M, 1.75, 7.1, 2.65);
  s.addImage({ path: A + "lockup_light.png", x: M + 0.5, y: 2.3, w: 6.1, h: 1.63 });
  card(s, M + 7.4, 1.75, 2.4, 2.65, C.text1);
  s.addImage({ path: A + "lockup_dark.png", x: M + 7.55, y: 2.5, w: 2.1, h: 0.56 });
  T(s, "다크 버전", { x: M + 7.4, y: 3.85, w: 2.4, h: 0.3, fontSize: 11, color: C.accent4, align: "center" });
  s.addImage({ path: A + "app_icon.png", x: M + 10.25, y: 1.85, w: 1.85, h: 1.85 });
  T(s, "앱 아이콘", { x: M + 10.1, y: 3.85, w: 2.1, h: 0.3, fontSize: 11, color: C.accent3, align: "center" });
  const mean = [["두 개의 프레임", "배우의 장면과 연출자의 장면. 어느 한쪽이 아니라 둘이 겹쳐야 한다"], ["노란 교집합", "형광펜 = 대본에서 ‘내 대사’를 표시하는 배우의 습관. 서로의 장면이 겹치는 순간이 MATCH"], ["대본 포맷", "S#(씬 넘버)·타자기 서체로 ‘현장의 언어’를 브랜드 언어로"]];
  mean.forEach((x, i) => {
    const y = 4.65 + i * 0.68;
    T(s, x[0], { x: M, y, w: 2.2, h: 0.5, fontSize: 15, bold: true });
    T(s, x[1], { x: M + 2.2, y, w: 5.0, h: 0.6, fontSize: 13, color: C.accent3, lineSpacingMultiple: 1.2 });
  });
  const sw = [["INK", "141414", C.text1], ["HIGHLIGHTER", "FFE14D", C.accent1], ["RED PEN", "D92B2B", C.accent2], ["PAPER", "FFFFFF", C.background1]];
  sw.forEach((c, i) => {
    const X = 8.0 + (i % 2) * 2.4, y = 4.65 + Math.floor(i / 2) * 0.85;
    s.addShape(pres.ShapeType.rect, { x: X, y, w: 0.6, h: 0.6, fill: { color: c[2] }, line: { color: HEX.accent4, width: 0.75 } });
    T(s, c[0], { x: X + 0.75, y: y + 0.03, w: 1.6, h: 0.28, fontSize: 12, bold: true });
    T(s, "#" + c[1], { x: X + 0.75, y: y + 0.32, w: 1.6, h: 0.25, fontSize: 11, fontFace: MONO, color: C.accent3 });
  });
  T(s, "서체: Pretendard (본문·제목) + Courier New (씬 넘버)", { x: 8.0, y: 6.38, w: 4.8, h: 0.3, fontSize: 11.5, color: C.accent3 });
}

// ═════════ 24. COPY
{
  const s = L("DARK", "브랜딩");
  head(s, "INT. 카피룸 — NIGHT", "시선을 잡는 3개의 문장, 우선순위대로", "3개의 문장", "추론", true);
  const cp = [
    ["1순위 · 런칭 메인", "오디션장 밖에서,\n캐스팅된다.", "캐스팅이 일어나는 장소를 ‘오디션장’에서 ‘일상’으로 옮긴다. 신인에게는 기회의 확장, 연출자에게는 상시 발견. 짧고 단정해 옥외·표지·앱스토어 어디든 맞는다."],
    ["2순위 · 챌린지 캠페인", "같은 대사,\n천 개의 얼굴.", "LEGEND SCRIPT의 본질(비교·해석)을 그대로 담았다. 숏폼 몽타주 제목으로 쓰기 좋고, 보는 사람도 참여하고 싶게 만든다."],
    ["3순위 · 리텐션·리타겟팅", "지금도 누군가,\n당신의 장면을\n다시 보고 있다.", "FOMO와 자부심을 동시에 건드린다. 푸시 알림·티저·재방문 광고용. 업로드를 멈추지 않게 하는 문장."],
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

// ═════════ 25. SELF SCORE
pres.addSection({ title: "평가" });
{
  const s = L("LIGHT", "평가");
  head(s, "INT. 리뷰 — DAY", "95점 — 남은 5점을 채우는 방법", "남은 5점", "추론");
  const g = [
    ["−2", "실측 수요 데이터 없음", "지망생 지불의향·연출자 사용의향을 직접 검증한 데이터가 없다", ["랜딩페이지·SNS 광고 50만 원으로 사전등록률 측정", "연극영화과·연기학원 300명 설문", "독립영화 감독 20명 인터뷰로 Studio Pro 지불의향 검증"]],
    ["−2", "나문희 배우 참여 미확정", "킬러 콘텐츠가 본인 동의를 전제로 한다", ["심영은 배우 EP.000 파일럿 → 반응 데이터로 제안", "참여 최소화: 대본 선정 + 손편지, 촬영 1회", "수익 일부를 신인 장학금으로 — 참여 명분 강화"]],
    ["−1", "법률 구조 미확정", "알선 수수료·직업소개 해당 여부 해석이 필요하다", ["출시 전 엔터 전문 변호사 사전 자문", "초기 수수료 0원 구조로 리스크 회피", "콘진원·규제샌드박스 상담 채널 활용"]],
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

// ═════════ 26. SOURCES
{
  const s = L("LIGHT", "평가");
  head(s, "INT. 자료실 — DAY", "‘사실’ 표기의 근거 자료", null);
  const src = [
    "문체부·한국콘텐츠진흥원 ‘2025 대중문화예술산업 실태조사’ (한국경제·뉴스1 보도) — 기획업 4,471곳, 제작업 2,289곳, 소속 예술인 12,092명 (2024년 기준)",
    "진학사 — 2026학년도 연극·영화·방송 계열 수시 일반전형 정원 2,033명, 지원 52,010명, 경쟁률 25.58:1",
    "파이낸셜뉴스(2026.11) — 실연자 사용료 징수액 2024 약 470억 → 2025 약 340억 원",
    "한국소비자원 / 코리안센터 — 학원형 연예기획사 피해구제 366건, 상반기 91건(+75%)",
    "세계일보(2025.9) — 문체부 기획업 일제 등록 계도기간 및 미등록 엄정 조치 방침",
    "혁신의숲 · 스포츠경향(2024.8) — 플필 배우 약 1.67만 명 연결, 신인 오디션 3,000명+ 지원",
    "머니투데이 유니콘팩토리(2026.5) — 베피플·캐스틱 캐스팅 사업 사례",
    "TA Associates 보도자료 — Backstage 유료 구독자 25만+, StarNow·Mandy 인수",
    "TheWrap · Authors Guild — SAG-AFTRA AI 디지털 복제 동의·보상 조항",
  ];
  T(s, bullets(src), { x: M, y: 1.8, w: CW, h: 4.4, fontSize: 13, paraSpaceAfter: 7, lineSpacingMultiple: 1.1 });
  T(s, "※ 수치는 보도 시점 기준. 외부 제출 전 원문 재확인 권장. 시장규모·손익·예산은 모두 검증용 가설(추론)입니다.", { x: M, y: 6.3, w: CW, h: 0.35, fontSize: 11.5, italic: true, color: C.accent3 });
}

// ═════════ 27. CLOSING + RIGHTS
pres.addSection({ title: "마무리" });
{
  const s = L("COVER", "마무리");
  s.addImage({ path: A + "mark_dark.png", x: W / 2 - 0.65, y: 0.6, w: 1.3, h: 1.3 });
  T(s, "FADE OUT.", { x: M, y: 2.05, w: CW, h: 0.35, fontSize: 14, fontFace: MONO, bold: true, color: C.accent1, align: "center" });
  const t = "오디션장 밖에서, 캐스팅된다.";
  T(s, t, { x: M, y: 2.5, w: CW, h: 0.9, fontSize: 44, fontFace: "Pretendard Black", color: C.background1, align: "center", valign: "middle" });
  T(s, "ACT MATCH  by  WEUM CORPORATION", { x: M, y: 3.45, w: CW, h: 0.35, fontSize: 13, fontFace: MONO, bold: true, color: C.accent4, align: "center" });
  card(s, 1.4, 4.25, W - 2.8, 2.45, C.accent5);
  T(s, "권리 고지 · NOTICE OF RIGHTS", { x: 1.8, y: 4.45, w: W - 3.6, h: 0.3, fontSize: 12, fontFace: MONO, bold: true, color: C.accent1 });
  T(s, "본 기획안에 포함된 사업 아이디어, 서비스 구조, 브랜드명(ACT MATCH), 로고, 카피, 콘텐츠 포맷(LEGEND SCRIPT · SAFE CAST · SET REVIEW 등), 수익모델, 재무 가설 및 모든 구성에 관한 저작권과 일체의 권리는 WEUM CORPORATION에 있습니다. 사전 서면 동의 없는 복제·배포·전송·변형·2차적 저작물 작성 및 영업상 이용을 금하며, 본 자료는 「저작권법」 및 「부정경쟁방지 및 영업비밀보호에 관한 법률」에 따라 보호되는 영업비밀입니다.",
    { x: 1.8, y: 4.85, w: W - 3.6, h: 1.3, fontSize: 11.5, color: C.accent4, lineSpacingMultiple: 1.3 });
  T(s, "© 2026 WEUM CORPORATION. All Rights Reserved.", { x: 1.8, y: 6.2, w: W - 3.6, h: 0.3, fontSize: 12, bold: true, color: C.background1 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written");
})();
