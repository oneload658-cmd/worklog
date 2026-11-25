const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/06c47b33-d640-428c-a3a5-a42ae98094c8_2ae10ec8-84d2-43dd-8c58-6c8a9c68ff68/pptx/scripts/apply_theme.js");
const A = "/home/user/worklog/act-match/assets/";
const OUT = "/home/user/worklog/act-match/ACT_MATCH_기획안_WEUM.pptx";

const THEME = {
  name: "ACT MATCH",
  headFontFace: "Pretendard ExtraBold",
  bodyFontFace: "Pretendard",
  colors: {
    dk1: "111111", lt1: "FFFFFF", dk2: "2A2724", lt2: "F3F1EE",
    accent1: "E8432E", accent2: "F2A93B", accent3: "6B6660", accent4: "DDD8D2",
    accent5: "1E1C1A", accent6: "FCE9E5", hlink: "E8432E", folHlink: "6B6660",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "ACT MATCH 사업 기획안";
pres.author = "WEUM CORPORATION";
pres.company = "WEUM CORPORATION";
const C = pres.SchemeColor;
const W = 13.333, M = 0.6;
const FOOT = "ACT·MATCH 사업 기획안 v2.0  |  CONFIDENTIAL  |  © 2026 WEUM CORPORATION";

function frame(name, dark) {
  pres.defineSlideMaster({
    title: name,
    background: { color: dark ? C.text1 : C.background1 },
    objects: [
      { image: { x: M, y: 6.93, w: 0.26, h: 0.26, path: A + (dark ? "mark_dark_t.png" : "mark_light_t.png") } },
      { text: { text: FOOT, options: { x: M + 0.36, y: 6.93, w: 8, h: 0.26, fontSize: 9, color: C.accent3, margin: 0, valign: "middle", charSpacing: 1 } } },
      { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.42, w: 9, h: 0.32, fontSize: 12, bold: true, color: C.accent1, margin: 0, charSpacing: 3 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.78, w: W - 2 * M, h: 0.85, fontSize: 30, color: dark ? C.background1 : C.text1, margin: 0, valign: "top", align: "left", fontFace: "Pretendard ExtraBold" }, text: "" } },
    ],
    slideNumber: { x: W - M - 0.6, y: 6.93, w: 0.6, h: 0.26, fontSize: 9, color: C.accent3, align: "right" },
  });
}
frame("LIGHT", false);
frame("DARK", true);
pres.defineSlideMaster({ title: "COVER", background: { color: C.text1 }, objects: [] });

function head(s, kicker, title) {
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
}
// 사실 / 추론 badge
function tag(s, x, y, kind) {
  const fact = kind === "사실";
  s.addText(kind, {
    x, y, w: 0.62, h: 0.28, shape: pres.ShapeType.roundRect, rectRadius: 0.14, isTextBox: true,
    fill: { color: fact ? C.text1 : C.accent1 }, color: C.background1,
    fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0,
  });
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { type: "none" } });
}
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text1, fontSize: 14 }, o)); }

let n = 0;
const sec = (t) => pres.addSection({ title: t });

// ───────── 1. COVER
sec("표지");
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
  s.addImage({ path: A + "lockup_dark.png", x: M, y: 0.6, w: 3.5, h: 1.0 });
  T(s, "BUSINESS PROPOSAL  v2.0", { x: M, y: 2.35, w: 8, h: 0.35, fontSize: 13, bold: true, color: C.accent1, charSpacing: 4 });
  T(s, [
    { text: "프로필은 꾸밀 수 있다.", options: { breakLine: true } },
    { text: "연기는 꾸밀 수 없다.", options: { color: C.accent1 } },
  ], { x: M, y: 2.8, w: 11, h: 1.9, fontSize: 50, fontFace: "Pretendard Black", color: C.background1, lineSpacingMultiple: 1.05 });
  T(s, "연기 아카이브 기반 양방향 캐스팅 플랫폼 — 사업 기획안", { x: M, y: 4.95, w: 10, h: 0.4, fontSize: 18, color: C.accent4 });
  T(s, "2026. 11  ·  WEUM CORPORATION  ·  CONFIDENTIAL", { x: M, y: 6.75, w: 8, h: 0.3, fontSize: 11, color: C.accent3, charSpacing: 2 });
  s.addImage({ path: A + "mark_dark_t.png", x: 9.3, y: 2.2, w: 3.6, h: 3.6, transparency: 82 });
}

// ───────── 2. EXEC SUMMARY
sec("요약");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "요약" });
  head(s, "EXECUTIVE SUMMARY", "스와이프 앱이 아니라, ‘연기를 증명하는 인프라’");
  card(s, M, 1.85, 4.6, 4.75, C.text1);
  T(s, "ONE LINE", { x: M + 0.4, y: 2.2, w: 3.8, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 3 });
  T(s, "같은 대본으로\n배우를 비교하고,\n서로 고르는 순간\n불이 붙는다.", { x: M + 0.4, y: 2.65, w: 3.9, h: 2.6, fontSize: 26, fontFace: "Pretendard ExtraBold", color: C.background1, lineSpacingMultiple: 1.15 });
  T(s, "원안(GPT) 대비 5가지를 바꿨습니다 →", { x: M + 0.4, y: 5.75, w: 3.9, h: 0.4, fontSize: 14, color: C.accent4 });
  tag(s, M + 0.4, 6.15, "추론");
  const rows = [
    ["01", "포지셔닝", "‘연기자용 틴더’ → 케미스트리 엔진", "오디션 사기·성범죄 이슈 속 ‘데이팅’ 연상은 역효과"],
    ["02", "매칭 단위", "사람(프로필) → 장면(SCENE)", "같은 대본·같은 조건 비교가 핵심 데이터"],
    ["03", "킬러 콘텐츠", "익명 대본 → LEGEND SCRIPT", "나문희·심영은 배우가 건네는 대본 (WEUM 독점 자산)"],
    ["04", "수익 구조", "배우 구독 → 제작 측 과금", "배우는 평생 무료, 시간을 아끼는 쪽이 지불"],
    ["05", "신뢰 설계", "‘필요하다’ 선언 → 실행 규칙", "SAFE CAST 5원칙 · 배우 데이터 헌장 · WEUM 방화벽"],
  ];
  rows.forEach((r, i) => {
    const y = 1.85 + i * 0.97;
    T(s, r[0], { x: 5.65, y, w: 0.6, h: 0.5, fontSize: 22, fontFace: "Pretendard Black", color: C.accent1 });
    T(s, r[1], { x: 6.3, y: y + 0.04, w: 1.6, h: 0.4, fontSize: 14, bold: true, color: C.accent3 });
    T(s, r[2], { x: 7.9, y: y + 0.01, w: 4.85, h: 0.4, fontSize: 17, bold: true });
    T(s, r[3], { x: 7.9, y: y + 0.42, w: 4.85, h: 0.35, fontSize: 13, color: C.accent3 });
    if (i < 4) s.addShape(pres.ShapeType.line, { x: 5.65, y: y + 0.88, w: 7.08, h: 0, line: { color: C.accent4, width: 0.75 } });
  });
}

// ───────── 3. MARKET
sec("시장과 경쟁");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "시장과 경쟁" });
  head(s, "MARKET REALITY", "일은 줄고, 지망생은 넘치고, 사기는 늘었다");
  tag(s, 12.1, 0.43, "사실");
  const stats = [
    ["-130억", "실연자 사용료 징수액", "2024년 약 470억 → 2025년 약 340억 원.\n업계는 드라마 제작 편수 감소를 원인으로 지목", "파이낸셜뉴스, 2026.11"],
    ["80%", "출연료 1천만 원 미만", "방송배우 응답자 10명 중 8명.\n서면 계약은 절반에 그침", "서울시·한국방송실연자협회 조사(2020)"],
    ["+75%", "기획사 피해구제 급증", "2023~2026.6 피해구제 366건.\n상반기 91건(전년 동기 52건), 절반이 아동·청소년", "한국소비자원 / 코리안센터"],
    ["27%", "20대 숏드라마 이용률", "글로벌 숏폼드라마 시장 2025년 약 72.3억 달러,\n연평균 10.6% 성장 전망", "Research and Markets / 업계 보도"],
  ];
  const cw = (W - 2 * M - 3 * 0.3) / 4;
  stats.forEach((st, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 1.85, cw, 3.55);
    T(s, st[0], { x: x + 0.3, y: 2.1, w: cw - 0.6, h: 0.9, fontSize: 40, fontFace: "Pretendard Black", color: i === 2 ? C.accent1 : C.text1 });
    T(s, st[1], { x: x + 0.3, y: 3.05, w: cw - 0.6, h: 0.4, fontSize: 15, bold: true });
    T(s, st[2], { x: x + 0.3, y: 3.5, w: cw - 0.6, h: 1.2, fontSize: 12.5, color: C.accent3, lineSpacingMultiple: 1.2 });
    T(s, st[3], { x: x + 0.3, y: 4.95, w: cw - 0.6, h: 0.3, fontSize: 10, color: C.accent3, italic: true });
  });
  card(s, M, 5.65, W - 2 * M, 1.0, C.text1);
  tag(s, M + 0.35, 5.99, "추론");
  T(s, [
    { text: "시장이 원하는 건 ‘더 많은 오디션’이 아니라 ", options: {} },
    { text: "‘믿을 수 있는 발견’", options: { color: C.accent2, bold: true } },
    { text: "이다. 그래서 우리는 신뢰를 기능이 아닌 상품으로 판다.", options: {} },
  ], { x: M + 1.15, y: 5.65, w: 10.8, h: 1.0, fontSize: 17, color: C.background1, valign: "middle" });
}

// ───────── 4. COMPETITION
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "시장과 경쟁" });
  head(s, "COMPETITION", "‘영상 프로필 + AI + 스와이프’는 이미 레드오션이다");
  tag(s, 12.1, 0.43, "사실");
  const hdr = { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: 13, valign: "middle" };
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 13, valign: "middle", color: C.text1 }, o || {}) });
  const rows = [
    [{ text: "서비스", options: hdr }, { text: "확인된 사실", options: hdr }, { text: "빈틈 (추론)", options: Object.assign({}, hdr, { fill: { color: C.accent1 } }) }],
    [cell("플필 PLFIL", { bold: true }), cell("등록 배우 약 1.67만 명 연결, 유료 ‘플필 플러스’, 자체 숏폼드라마 오디션 한 달 3,000명+ 지원(2024.8)"), cell("배우→공고 지원형. 연출자 쪽 정보는 비대칭")],
    [cell("캐스틱 / 베피플", { bold: true }), cell("캐스팅 검토 7일→1일 단축 / 숏폼 영상 프로필, 디렉터 48시간 승인제(2026.5 보도)"), cell("속도·검증은 하지만 ‘같은 대본 비교’ 없음")],
    [cell("Backstage (美)", { bold: true }), cell("유료 구독자 25만+ 보유, StarNow·Mandy 인수로 글로벌 통합"), cell("구독 모델이 배우 부담. 한국 시장 미진출")],
    [cell("CASTON · SCENEPORT", { bold: true }), cell("AI 영상 분석·추천 / 셀프테이프·지원 관리"), cell("AI 점수화는 예술계 반감, 데이터 동의 이슈")],
  ];
  s.addTable(rows, { x: M, y: 1.85, w: W - 2 * M, colW: [2.4, 5.9, 3.833], rowH: [0.5, 0.78, 0.78, 0.78, 0.78], border: { type: "solid", pt: 0.75, color: HEX.accent4 }, margin: [0.06, 0.15, 0.06, 0.15] });
  card(s, M, 5.75, W - 2 * M, 0.9, C.accent6);
  T(s, [
    { text: "아무도 갖지 못한 3가지  ", options: { bold: true, color: C.accent1 } },
    { text: "① 같은 대본 비교 데이터  ② 연출자 평판·포트폴리오  ③ 레전드 배우 콘텐츠", options: { bold: true } },
  ], { x: M + 0.35, y: 5.75, w: W - 2 * M - 0.7, h: 0.9, fontSize: 16, valign: "middle" });
}

// ───────── 5. DIAGNOSIS
sec("원안 진단");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "원안 진단" });
  head(s, "DIAGNOSIS · 원안(GPT) 평가", "좋은 비전, 그러나 ‘실행되지 않는’ 5개의 구멍");
  tag(s, 12.1, 0.43, "추론");
  card(s, M, 1.85, 3.1, 4.8, C.text1);
  T(s, "원안 점수", { x: M + 0.35, y: 2.15, w: 2.4, h: 0.3, fontSize: 13, bold: true, color: C.accent2, charSpacing: 2 });
  T(s, [{ text: "68", options: { fontSize: 88, fontFace: "Pretendard Black" } }, { text: " /100", options: { fontSize: 20, color: C.accent3 } }], { x: M + 0.35, y: 2.5, w: 2.6, h: 1.4, color: C.background1, valign: "middle" });
  T(s, "비전·구조 정리는 우수.\n차별화·수익·안전·공급 확보의 ‘방법’이 비어 있음.", { x: M + 0.35, y: 4.2, w: 2.45, h: 1.6, fontSize: 13, color: C.accent4, lineSpacingMultiple: 1.25 });
  const items = [
    ["‘연기자용 틴더’ 프레이밍", "학원형 기획사 피해가 75% 늘어난 시장에서 ‘데이팅’ 연상은 부모·여성 배우의 불안을 키운다", "케미스트리 언어 + SAFE CAST"],
    ["최대 경쟁자 플필 누락", "이미 영상 프로필·유료 멤버십·자체 숏폼 제작까지 하는 1위 사업자를 빠뜨림", "‘장면 비교’로 차별 축 재정의"],
    ["연출자 공급 해법 없음", "양면 플랫폼은 공급(연출자)이 없으면 배우가 이탈. 획득 전략 부재", "Director Seed 30"],
    ["배우 구독 의존", "배우 다수가 저소득. 구독료는 가입 장벽이 된다", "배우 평생 무료, B2B 과금"],
    ["8개 아카이브 · 방화벽 선언", "MVP에 과한 범위, WEUM 이해상충은 ‘분리해야 한다’로만 끝남", "3레이어 축소 + 방화벽 4규칙"],
  ];
  items.forEach((it, i) => {
    const y = 1.85 + i * 0.97, x = 4.1;
    T(s, String(i + 1).padStart(2, "0"), { x, y, w: 0.55, h: 0.45, fontSize: 20, fontFace: "Pretendard Black", color: C.accent1 });
    T(s, it[0], { x: x + 0.6, y: y + 0.02, w: 5.3, h: 0.38, fontSize: 16, bold: true });
    T(s, it[1], { x: x + 0.6, y: y + 0.41, w: 5.3, h: 0.5, fontSize: 12, color: C.accent3 });
    T(s, "→ " + it[2], { x: 10.05, y: y + 0.1, w: 2.7, h: 0.5, fontSize: 12, bold: true, color: C.accent1 });
    if (i < 4) s.addShape(pres.ShapeType.line, { x, y: y + 0.9, w: 8.63, h: 0, line: { color: C.accent4, width: 0.75 } });
  });
}

// ───────── 6. POSITIONING (dark)
sec("컨셉");
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "컨셉" });
  head(s, "NEW POSITIONING", "PROOF OF ACTING — 우리는 배우를 소개하지 않는다. 연기를 증명한다.");
  tag(s, 12.1, 0.43, "추론");
  const P = [
    ["SCENE", "같은 장면", "모든 배우가 같은 대본·같은 조건으로 연기한다. 경력·외모·소속사가 아니라 ‘해석’이 비교된다."],
    ["CHEMISTRY", "서로 고른다", "연출자도 포트폴리오와 현장 평판을 공개한다. 둘 다 ‘SPARK’를 눌러야 대화가 열린다."],
    ["SAFE", "안심하고 만난다", "검증된 연출자만, 앱 안에서만, 공개된 장소에서만. 신뢰가 우리의 상품이다."],
  ];
  const cw = (W - 2 * M - 2 * 0.3) / 3;
  P.forEach((p, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 2.0, cw, 4.55, C.accent5);
    T(s, "0" + (i + 1), { x: x + 0.4, y: 2.35, w: 1, h: 0.4, fontSize: 14, bold: true, color: C.accent1 });
    T(s, p[0], { x: x + 0.4, y: 2.8, w: cw - 0.8, h: 0.8, fontSize: 34, fontFace: "Pretendard Black", color: C.background1 });
    T(s, p[1], { x: x + 0.4, y: 3.65, w: cw - 0.8, h: 0.45, fontSize: 18, bold: true, color: C.accent2 });
    T(s, p[2], { x: x + 0.4, y: 4.3, w: cw - 0.8, h: 1.9, fontSize: 14, color: C.accent4, lineSpacingMultiple: 1.3 });
  });
}

// ───────── 7. FLOW
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "컨셉" });
  head(s, "SERVICE FLOW", "매칭의 단위는 ‘사람’이 아니라 ‘장면’이다");
  tag(s, 12.1, 0.43, "추론");
  const steps = [
    ["SCENE DROP", "연출자·레전드가\n대본 공개"],
    ["TAKE", "배우가 30~60초\n해석 연기 업로드"],
    ["KEEP", "연출자가 장면을\n후보 보드에 저장"],
    ["SPARK", "배우가 연출자\n포트폴리오 확인 후 수락"],
    ["MEET", "앱 내 대화 →\n공개 장소 오디션"],
    ["LEDGER", "작품 참여가\n검증 이력으로 기록"],
  ];
  const cw = 1.82, gap = (W - 2 * M - 6 * cw) / 5;
  steps.forEach((st, i) => {
    const x = M + i * (cw + gap);
    const hot = i === 3;
    s.addShape(pres.ShapeType.ellipse, { x: x + cw / 2 - 0.55, y: 2.0, w: 1.1, h: 1.1, fill: { color: hot ? C.accent1 : C.text1 }, line: { type: "none" } });
    T(s, String(i + 1), { x: x + cw / 2 - 0.55, y: 2.0, w: 1.1, h: 1.1, fontSize: 30, fontFace: "Pretendard Black", color: C.background1, align: "center", valign: "middle" });
    T(s, st[0], { x, y: 3.3, w: cw, h: 0.4, fontSize: 16, fontFace: "Pretendard Black", align: "center", color: hot ? C.accent1 : C.text1 });
    T(s, st[1], { x, y: 3.75, w: cw, h: 0.8, fontSize: 13, align: "center", color: C.accent3, lineSpacingMultiple: 1.2 });
    if (i < 5) s.addShape(pres.ShapeType.line, { x: x + cw / 2 + 0.65, y: 2.55, w: cw + gap - 1.3, h: 0, line: { color: C.accent4, width: 1.5, endArrowType: "triangle" } });
  });
  card(s, M, 4.95, (W - 2 * M - 0.3) / 2, 1.7);
  card(s, M + (W - 2 * M + 0.3) / 2, 4.95, (W - 2 * M - 0.3) / 2, 1.7);
  const hw = (W - 2 * M - 0.3) / 2;
  T(s, "배우가 얻는 것", { x: M + 0.35, y: 5.2, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  T(s, "프로필 사진·경력 없이 ‘연기 한 장면’으로 검증된 연출자에게 직접 도달. 거절권은 배우에게도 있다.", { x: M + 0.35, y: 5.6, w: hw - 0.7, h: 0.9, fontSize: 14, lineSpacingMultiple: 1.25 });
  const x2 = M + hw + 0.3;
  T(s, "연출자가 얻는 것", { x: x2 + 0.35, y: 5.2, w: hw - 0.7, h: 0.35, fontSize: 15, bold: true, color: C.accent1 });
  T(s, "같은 대본으로 수백 명의 해석을 나란히 비교. 작품이 없을 때도 ‘언젠가 함께할 배우’ 풀을 쌓는다.", { x: x2 + 0.35, y: 5.6, w: hw - 0.7, h: 0.9, fontSize: 14, lineSpacingMultiple: 1.25 });
}

// ───────── 8. UX + SAFE CAST
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "컨셉" });
  head(s, "UX · SAFE CAST PROTOCOL", "스와이프는 가볍게, 안전은 무겁게");
  tag(s, 12.1, 0.43, "추론");
  // phone mock
  const px = 1.0, py = 1.8, pw = 2.75, ph = 4.95;
  s.addShape(pres.ShapeType.roundRect, { x: px, y: py, w: pw, h: ph, rectRadius: 0.3, fill: { color: C.text1 }, line: { type: "none" } });
  s.addShape(pres.ShapeType.roundRect, { x: px + 0.15, y: py + 0.5, w: pw - 0.3, h: 2.6, rectRadius: 0.12, fill: { color: C.accent3 }, line: { type: "none" } });
  T(s, "▶  SCENE #001 · 30s", { x: px + 0.3, y: py + 0.65, w: pw - 0.6, h: 0.3, fontSize: 11, bold: true, color: C.background1 });
  T(s, "“미안해. 사실 나\n그날 거기 있었어.”", { x: px + 0.3, y: py + 1.4, w: pw - 0.6, h: 0.9, fontSize: 14, bold: true, color: C.background1, align: "center" });
  T(s, "배우 A · 해석: 죄책감", { x: px + 0.3, y: py + 3.25, w: pw - 0.6, h: 0.3, fontSize: 13, bold: true, color: C.background1 });
  T(s, "#절제 #생활연기 #스릴러", { x: px + 0.3, y: py + 3.6, w: pw - 0.6, h: 0.3, fontSize: 11, color: C.accent4 });
  const btn = [["PASS", C.accent3], ["KEEP", C.accent2], ["SPARK", C.accent1]];
  btn.forEach((b, i) => {
    s.addText(b[0], { x: px + 0.2 + i * 0.8, y: py + 4.2, w: 0.72, h: 0.42, shape: pres.ShapeType.roundRect, rectRadius: 0.21, isTextBox: true, fill: { color: b[1] }, color: C.background1, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  const rules = [
    ["연출자 3중 인증", "실명 + 사업자/소속 + 필모그래피 확인, 48시간 심사 후 ‘VERIFIED’ 배지"],
    ["매칭 전 DM 불가", "양쪽 SPARK 이후에만 대화. 대화는 앱 안에서만, 기록 보존"],
    ["공개 장소·주간 원칙", "오디션 장소·시간을 앱에 등록. ‘밀실 오디션’ 신고 시 즉시 정지"],
    ["금전 요구 원천 차단", "‘트레이닝비·프로필비·등록비’ 키워드 감지 → 경고·차단"],
    ["미성년자 보호", "만 19세 미만은 보호자 연동 계정, 대화에 보호자 자동 참여"],
  ];
  T(s, "SAFE CAST 5원칙", { x: 4.4, y: 1.85, w: 6, h: 0.4, fontSize: 18, fontFace: "Pretendard Black", color: C.accent1 });
  rules.forEach((r, i) => {
    const y = 2.4 + i * 0.86;
    s.addShape(pres.ShapeType.ellipse, { x: 4.4, y: y + 0.03, w: 0.42, h: 0.42, fill: { color: C.text1 }, line: { type: "none" } });
    T(s, String(i + 1), { x: 4.4, y: y + 0.03, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, r[0], { x: 5.0, y, w: 7.7, h: 0.35, fontSize: 15, bold: true });
    T(s, r[1], { x: 5.0, y: y + 0.36, w: 7.7, h: 0.4, fontSize: 13, color: C.accent3 });
  });
}

// ───────── 9. LEGEND SCRIPT (dark)
sec("킬러 콘텐츠");
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "킬러 콘텐츠" });
  head(s, "KILLER CONTENT · LEGEND SCRIPT", "“나문희가 당신의 연기를 봅니다.”");
  tag(s, 12.1, 0.43, "추론");
  card(s, M, 1.9, 5.6, 4.75, C.accent5);
  T(s, "EP.001  ·  LEGEND SCRIPT", { x: M + 0.4, y: 2.2, w: 4.8, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 3 });
  T(s, "마지막 전화", { x: M + 0.4, y: 2.6, w: 4.8, h: 0.7, fontSize: 32, fontFace: "Pretendard Black", color: C.background1 });
  T(s, "“미안해. 사실 나 그날 거기 있었어.”", { x: M + 0.4, y: 3.4, w: 4.8, h: 0.5, fontSize: 17, italic: true, color: C.accent4 });
  T(s, [
    { text: "대본 선정 · 30초 시범 연기", options: { bold: true, color: C.background1, breakLine: true } },
    { text: "나문희 배우 (시즌 1 호스트, 본인 동의 전제)", options: { color: C.accent4, breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "릴레이 호스트", options: { bold: true, color: C.background1, breakLine: true } },
    { text: "심영은 배우 → 협업 감독·작가 순환", options: { color: C.accent4 } },
  ], { x: M + 0.4, y: 4.2, w: 4.8, h: 2.2, fontSize: 14, lineSpacingMultiple: 1.25 });
  const steps = [
    ["WEEK 1", "레전드가 고른 3줄 대본 + 시범 연기 공개 (숏폼 바이럴 소재)"],
    ["WEEK 2", "신인 누구나 같은 대본으로 TAKE 업로드 · 비용 0원"],
    ["WEEK 3", "레전드가 TOP 10 장면에 ‘손편지 코멘트’ — 최고의 PR 자산"],
    ["WEEK 4", "TOP 3 → ACT ORIGINALS 단편 출연 확정 · 촬영기 공개"],
  ];
  steps.forEach((st, i) => {
    const y = 1.9 + i * 0.98;
    T(s, st[0], { x: 6.8, y, w: 1.2, h: 0.4, fontSize: 13, bold: true, color: C.accent1 });
    T(s, st[1], { x: 8.0, y, w: 4.75, h: 0.8, fontSize: 15, color: C.background1, lineSpacingMultiple: 1.2 });
  });
  card(s, 6.8, 5.85, 5.93, 0.8, C.accent1);
  T(s, "근거: 플필 신인 오디션은 ‘레전드’ 없이도 한 달 3,000명+ 지원 (사실, 2024.8)", { x: 7.05, y: 5.85, w: 5.5, h: 0.8, fontSize: 13, bold: true, color: C.background1, valign: "middle" });
}

// ───────── 10. ARCHIVE
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "킬러 콘텐츠" });
  head(s, "ARCHIVE · 양방향 기록", "8개 아카이브를 3개로 줄이고, 연출자 쪽 기록을 새로 만든다");
  tag(s, 12.1, 0.43, "추론");
  const hw = (W - 2 * M - 0.3) / 2;
  card(s, M, 1.85, hw, 4.8);
  card(s, M + hw + 0.3, 1.85, hw, 4.8, C.text1);
  T(s, "ACTOR ARCHIVE — MVP 3레이어", { x: M + 0.4, y: 2.1, w: hw - 0.8, h: 0.4, fontSize: 16, fontFace: "Pretendard Black", color: C.accent1 });
  const a = [
    ["SCENE REEL", "챌린지 대본 연기 — 비교 가능한 표준 데이터"],
    ["FREE REEL", "독백·자유연기·발성 — 개성의 기록"],
    ["CAREER LEDGER", "매칭·오디션·출연 이력 — 플랫폼 검증 스탬프"],
  ];
  a.forEach((r, i) => {
    const y = 2.7 + i * 1.05;
    T(s, r[0], { x: M + 0.4, y, w: hw - 0.8, h: 0.4, fontSize: 18, bold: true });
    T(s, r[1], { x: M + 0.4, y: y + 0.42, w: hw - 0.8, h: 0.4, fontSize: 14, color: C.accent3 });
  });
  T(s, "Phase 2: Voice · Growth(연도별 비교) 아카이브 추가", { x: M + 0.4, y: 5.95, w: hw - 0.8, h: 0.4, fontSize: 12, italic: true, color: C.accent3 });
  const x2 = M + hw + 0.3;
  T(s, "DIRECTOR ARCHIVE — 원안에 없던 것", { x: x2 + 0.4, y: 2.1, w: hw - 0.8, h: 0.4, fontSize: 16, fontFace: "Pretendard Black", color: C.accent2 });
  const d = [
    ["DIRECTOR REEL", "연출작·차기작 톤·현장 철학 30초 영상"],
    ["SET REVIEW", "함께 작업한 배우의 익명 현장 후기 (검증된 매칭만)"],
    ["CAST HISTORY", "이 감독이 발굴한 신인들 — 연출자의 ‘안목’이 브랜드"],
  ];
  d.forEach((r, i) => {
    const y = 2.7 + i * 1.05;
    T(s, r[0], { x: x2 + 0.4, y, w: hw - 0.8, h: 0.4, fontSize: 18, bold: true, color: C.background1 });
    T(s, r[1], { x: x2 + 0.4, y: y + 0.42, w: hw - 0.8, h: 0.4, fontSize: 14, color: C.accent4 });
  });
  T(s, "→ 배우가 연출자를 고를 근거가 생겨야 ‘양방향’이 진짜가 된다", { x: x2 + 0.4, y: 5.95, w: hw - 0.8, h: 0.4, fontSize: 13, bold: true, color: C.accent1 });
}

// ───────── 11. AI & DATA CHARTER
sec("AI와 신뢰");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "AI와 신뢰" });
  head(s, "AI · ACTOR DATA CHARTER", "AI는 ‘찾아주는’ 일만 한다. 판단과 소유는 사람에게");
  const cw = (W - 2 * M - 2 * 0.3) / 3;
  const cols = [
    ["AI가 하는 일", C.text1, ["장면 태그 자동 생성 (감정·톤·장르·발성)", "자연어 검색: “눈빛으로 불안을 절제하는 40대”", "비슷한 해석 묶어 보기 · 연출자 취향 학습"], "추론"],
    ["AI가 하지 않는 일", C.accent1, ["연기력 점수화 · 순위 매기기", "외모 기반 추천·랭킹", "동의 없는 얼굴·목소리 학습, 디지털 복제"], "추론"],
    ["배우 데이터 헌장", C.accent2, ["영상 소유권은 100% 배우에게", "AI 학습은 별도 opt-in, 작품별 개별 동의", "탈퇴 시 30일 내 영상·파생데이터 삭제"], "추론"],
  ];
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 1.85, cw, 3.75);
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.35, y: 2.15, w: 0.5, h: 0.5, fill: { color: c[1] }, line: { type: "none" } });
    T(s, c[0], { x: x + 1.0, y: 2.2, w: cw - 1.3, h: 0.4, fontSize: 18, fontFace: "Pretendard Black" });
    T(s, c[2].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < c[2].length - 1 } })), { x: x + 0.35, y: 2.95, w: cw - 0.7, h: 2.4, fontSize: 14, paraSpaceAfter: 10, lineSpacingMultiple: 1.15 });
  });
  card(s, M, 5.85, W - 2 * M, 0.8, C.accent6);
  tag(s, M + 0.35, 6.11, "사실");
  T(s, "할리우드 SAG-AFTRA 협약: 디지털 복제는 사전 동의·용도 고지·보상 필수, 포괄적(blanket) 동의 불가 — 이 기준을 한국 최초로 플랫폼 약관에 넣는다", { x: M + 1.15, y: 5.85, w: W - 2 * M - 1.5, h: 0.8, fontSize: 13.5, valign: "middle" });
}

// ───────── 12. REVENUE
sec("수익과 생태계");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "수익과 생태계" });
  head(s, "REVENUE MODEL", "배우는 평생 무료. 시간을 아끼는 쪽이 지불한다");
  tag(s, 12.1, 0.43, "추론");
  const rows = [
    ["Actor", "무료", "업로드·챌린지·매칭 전부 무료 (Plus 월 9,900원: 조회 통계·추가 저장)"],
    ["Director", "무료", "검증 연출자 월 5회 SPARK, 후보 보드 1개"],
    ["Studio Pro", "월 29만 원~", "팀 후보 보드·NDA 대본 배포·셀프테이프 수집·무제한 SPARK"],
    ["Brand Challenge", "회당 3,000만 원~", "브랜드 대본 챌린지 → 광고 모델 발굴 (WEUM 광고 기획 역량)"],
    ["ACT ORIGINALS", "IP 수익", "챌린지 TOP 배우 출연 단편·숏드라마 유통·판권"],
  ];
  rows.forEach((r, i) => {
    const y = 1.85 + i * 0.95;
    card(s, M, y, 7.4, 0.8, i === 2 || i === 3 ? C.accent6 : C.background2);
    T(s, r[0], { x: M + 0.3, y, w: 2.0, h: 0.8, fontSize: 15, bold: true, valign: "middle" });
    T(s, r[1], { x: M + 2.3, y, w: 1.75, h: 0.8, fontSize: 14, bold: true, color: C.accent1, valign: "middle" });
    T(s, r[2], { x: M + 4.05, y, w: 3.2, h: 0.8, fontSize: 12, color: C.accent3, valign: "middle" });
  });
  s.addChart(pres.ChartType.doughnut, [{ name: "2년차 매출 구성 목표", labels: ["Studio Pro", "Brand Challenge", "ACT ORIGINALS", "Actor Plus"], values: [45, 30, 15, 10] }], {
    x: 8.4, y: 1.75, w: 4.35, h: 4.3, holeSize: 58, showTitle: true, title: "2년차 매출 구성 목표 (%)", titleFontSize: 13, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    chartColors: [HEX.dk1, HEX.accent1, HEX.accent2, HEX.accent4], showValue: true, showPercent: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt",
    showLegend: true, legendPos: "b", legendFontSize: 11, legendFontFace: "+mn-lt", legendColor: HEX.accent3,
  });
  T(s, "가격·비중은 검증용 가설치이며 MVP 파일럿 후 확정", { x: 8.4, y: 6.15, w: 4.35, h: 0.35, fontSize: 11, italic: true, color: C.accent3, align: "center" });
}

// ───────── 13. WEUM ECOSYSTEM + FIREWALL
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "수익과 생태계" });
  head(s, "WEUM ECOSYSTEM", "플랫폼은 모두에게 열고, 매니지먼트는 방화벽 뒤에 둔다");
  tag(s, 12.1, 0.43, "추론");
  const eco = [["ACT MATCH", "유입·아카이브"], ["ACT LAB", "워크숍·인큐베이팅"], ["WEUM MGMT", "선별 매니지먼트"], ["ACT ORIGINALS", "단편·숏드라마"], ["CONTENT IP", "유통·판권·광고"]];
  const cw = 2.25, gap = (W - 2 * M - 5 * cw) / 4;
  eco.forEach((e, i) => {
    const x = M + i * (cw + gap);
    card(s, x, 1.85, cw, 1.45, i === 0 ? C.accent1 : C.text1);
    T(s, e[0], { x: x + 0.2, y: 2.05, w: cw - 0.4, h: 0.45, fontSize: 16, fontFace: "Pretendard Black", color: C.background1, align: "center" });
    T(s, e[1], { x: x + 0.2, y: 2.6, w: cw - 0.4, h: 0.4, fontSize: 13, color: C.accent4, align: "center" });
    if (i < 4) T(s, "›", { x: x + cw, y: 2.25, w: gap, h: 0.6, fontSize: 26, bold: true, color: C.accent3, align: "center", valign: "middle" });
  });
  T(s, "FIREWALL 4규칙 — 원안의 ‘분리해야 한다’를 ‘이렇게 분리한다’로", { x: M, y: 3.65, w: 11, h: 0.4, fontSize: 17, fontFace: "Pretendard Black", color: C.accent1 });
  const fw = [
    ["노출 가중치 0", "추천·검색 알고리즘에 소속사 변수 미사용, 코드 수준 명시"],
    ["동일 조건 스카우트", "WEUM 스카우트도 일반 연출자와 같은 인증·같은 SPARK 한도"],
    ["7일 숙려 고지", "전속 제안 시 ‘거절해도 플랫폼 이용 불이익 없음’ 고지 + 7일 숙려"],
    ["연 1회 외부 점검", "추천 공정성 리포트를 외부 검토 후 공개"],
  ];
  const fwW = (W - 2 * M - 0.3) / 2;
  fw.forEach((f, i) => {
    const x = M + (i % 2) * (fwW + 0.3), y = 4.2 + Math.floor(i / 2) * 1.25;
    card(s, x, y, fwW, 1.1);
    T(s, f[0], { x: x + 0.35, y: y + 0.18, w: fwW - 0.7, h: 0.35, fontSize: 15, bold: true });
    T(s, f[1], { x: x + 0.35, y: y + 0.58, w: fwW - 0.7, h: 0.4, fontSize: 13, color: C.accent3 });
  });
}

// ───────── 14. ROADMAP
sec("실행 계획");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "실행 계획" });
  head(s, "6-MONTH MVP ROADMAP", "연출자를 먼저 모으고, 레전드로 불을 붙인다");
  tag(s, 12.1, 0.43, "추론");
  const ms = [
    ["M1", "SEED", "WEUM 네트워크 독립영화 감독 30명 사전 합류 · 배우 웨이팅리스트"],
    ["M2", "CLOSED BETA", "배우 300명 초대 · EP.001 LEGEND SCRIPT 공개"],
    ["M3", "OPEN", "주간 챌린지 정례화 · SAFE CAST 운영 시작"],
    ["M4", "FIRST FILM", "TOP 3 배우 ACT ORIGINALS 단편 1편 촬영"],
    ["M5", "B2B PILOT", "Studio Pro 유료 파일럿 3개사 · 브랜드 챌린지 1건"],
    ["M6", "PROOF", "성과 리포트 공개 · 시드 투자 / 정부지원사업 연계"],
  ];
  const cw = (W - 2 * M - 5 * 0.2) / 6;
  ms.forEach((m, i) => {
    const x = M + i * (cw + 0.2);
    card(s, x, 1.85, cw, 2.75, i === 1 ? C.text1 : C.background2);
    const dark = i === 1;
    T(s, m[0], { x: x + 0.25, y: 2.05, w: cw - 0.5, h: 0.5, fontSize: 24, fontFace: "Pretendard Black", color: C.accent1 });
    T(s, m[1], { x: x + 0.25, y: 2.6, w: cw - 0.5, h: 0.35, fontSize: 13, bold: true, color: dark ? C.background1 : C.text1 });
    T(s, m[2], { x: x + 0.25, y: 3.0, w: cw - 0.5, h: 1.5, fontSize: 12, color: dark ? C.accent4 : C.accent3, lineSpacingMultiple: 1.2 });
  });
  T(s, "6개월 검증 KPI (목표치 · 실적 아님)", { x: M, y: 4.85, w: 8, h: 0.35, fontSize: 15, bold: true });
  const k = [["1,000명", "가입 배우"], ["50명", "검증 연출자"], ["5,000개", "장면 영상"], ["20건", "매칭→오디션"], ["3곳", "유료 B2B"]];
  const kw = (W - 2 * M - 4 * 0.2) / 5;
  k.forEach((kk, i) => {
    const x = M + i * (kw + 0.2);
    T(s, kk[0], { x, y: 5.3, w: kw, h: 0.65, fontSize: 30, fontFace: "Pretendard Black", color: i === 3 ? C.accent1 : C.text1 });
    T(s, kk[1], { x, y: 5.95, w: kw, h: 0.35, fontSize: 13, color: C.accent3 });
  });
  T(s, "북극성 지표: ‘매칭→오디션 전환 수’ — 가입자 수가 아니라 실제 만남이 성과다", { x: M, y: 6.4, w: 12, h: 0.32, fontSize: 12, bold: true, color: C.accent1 });
}

// ───────── 15. LEGAL & RISK
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "실행 계획" });
  head(s, "LEGAL & RISK", "출시 전에 막아야 할 4개의 리스크");
  const r = [
    ["대중문화예술산업발전법", "사실", "기획업 미등록 시 2년 이하 징역 또는 2천만 원 이하 벌금(보도 기준). 문체부, 2025년 말까지 일제 등록 계도 후 엄정 조치 방침", "WEUM 기획업 등록 상태 점검 · 플랫폼 법인은 ‘정보 매칭’으로 한정"],
    ["캐스팅 알선 수수료", "추론", "출연 성사 수수료 도입 시 유료직업소개 등 별도 규제 해당 가능성", "초기 수수료 0원 · 구독/SaaS만 과금 · 전문 변호사 사전 자문"],
    ["영상·음성 개인정보", "추론", "얼굴·목소리 영상은 개인정보, AI 처리 시 민감도 상승", "개인정보 처리방침 · AI 학습 별도 동의 · 공개 범위 3단계"],
    ["상표·도메인", "추론", "ACT MATCH 유사 상표·도메인 선점 가능성", "KIPRIS 검색 → 9·35·41·42류 출원, 도메인 선확보"],
  ];
  T(s, "리스크", { x: M, y: 1.85, w: 3, h: 0.3, fontSize: 12, bold: true, color: C.accent3 });
  T(s, "근거", { x: 4.1, y: 1.85, w: 3, h: 0.3, fontSize: 12, bold: true, color: C.accent3 });
  T(s, "대응", { x: 8.9, y: 1.85, w: 3, h: 0.3, fontSize: 12, bold: true, color: C.accent3 });
  r.forEach((x, i) => {
    const y = 2.25 + i * 1.1;
    card(s, M, y, W - 2 * M, 0.98, i % 2 ? C.background1 : C.background2);
    T(s, x[0], { x: M + 0.3, y, w: 2.6, h: 0.98, fontSize: 15, bold: true, valign: "middle" });
    tag(s, 3.35, y + 0.35, x[1]);
    T(s, x[2], { x: 4.1, y, w: 4.55, h: 0.98, fontSize: 12.5, color: C.accent3, valign: "middle" });
    T(s, x[3], { x: 8.9, y, w: 3.75, h: 0.98, fontSize: 12.5, bold: true, valign: "middle" });
  });
}

// ───────── 16. BRANDING
sec("브랜딩");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "브랜딩" });
  head(s, "BRAND IDENTITY", "두 개의 성냥이 기대어 ‘A’가 되고, 하나의 불꽃이 핀다");
  tag(s, 12.1, 0.43, "추론");
  card(s, M, 1.85, 6.3, 2.6);
  s.addImage({ path: A + "lockup_light.png", x: M + 0.5, y: 2.45, w: 5.25, h: 1.5 });
  card(s, M + 6.6, 1.85, 2.3, 2.6, C.text1);
  s.addImage({ path: A + "mark_dark_t.png", x: M + 6.85, y: 2.05, w: 1.8, h: 1.8 });
  T(s, "앱 아이콘", { x: M + 6.6, y: 3.95, w: 2.3, h: 0.3, fontSize: 11, color: C.accent4, align: "center" });
  const sw = [["INK", "111111", C.text1], ["FLAME", "E8432E", C.accent1], ["AMBER", "F2A93B", C.accent2], ["ASH", "6B6660", C.accent3]];
  sw.forEach((c, i) => {
    const y = 1.85 + i * 0.66;
    s.addShape(pres.ShapeType.roundRect, { x: 9.8, y, w: 0.55, h: 0.55, rectRadius: 0.08, fill: { color: c[2] }, line: { color: HEX.accent4, width: 0.5 } });
    T(s, c[0], { x: 10.5, y: y + 0.02, w: 2.2, h: 0.28, fontSize: 12, bold: true });
    T(s, "#" + c[1], { x: 10.5, y: y + 0.29, w: 2.2, h: 0.25, fontSize: 11, color: C.accent3 });
  });
  const m = [
    ["두 성냥", "배우와 연출자. 어느 한쪽이 아니라 서로 기대야 선다"],
    ["A 형태", "ACT의 A. 가로대는 클래퍼보드(슬레이트)"],
    ["하나의 불꽃", "MATCH = 성냥 + 매칭. 맞닿는 순간 재능에 불이 붙는다"],
    ["서체", "Pretendard Black/Regular — 한글 가독성·자간 안정성 우선"],
  ];
  m.forEach((x, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const xx = M + col * 6.15, y = 4.75 + row * 0.95;
    T(s, x[0], { x: xx, y, w: 1.7, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
    T(s, x[1], { x: xx + 1.7, y, w: 4.3, h: 0.8, fontSize: 13.5, lineSpacingMultiple: 1.2 });
  });
}

// ───────── 17. COPY
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "브랜딩" });
  head(s, "HOOK COPY · 우선순위", "시선을 잡는 3개의 문장");
  tag(s, 12.1, 0.43, "추론");
  const cp = [
    ["1순위 · 런칭 메인", "프로필은 꾸밀 수 있다.\n연기는 꾸밀 수 없다.", "기존 시장의 불신(사진·스펙·인맥)을 정면으로 찌르는 대비형. 배우·연출자 모두 공감 → 브랜드 철학 그 자체"],
    ["2순위 · 가입 유도", "한 장면이면\n충분하다.", "경력·프로필 촬영비 없이 시작할 수 있다는 진입장벽 제거. 앱스토어·퍼포먼스 광고용"],
    ["3순위 · 리텐션", "오디션 없는 날에도,\n당신의 연기는\n일하고 있다.", "아카이브가 24시간 연출자에게 노출된다는 가치. 재방문·업로드 지속 동기"],
  ];
  const cw = (W - 2 * M - 2 * 0.3) / 3;
  cp.forEach((c, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 1.85, cw, 4.8, i === 0 ? C.accent1 : C.accent5);
    T(s, c[0], { x: x + 0.4, y: 2.15, w: cw - 0.8, h: 0.35, fontSize: 13, bold: true, color: i === 0 ? C.background1 : C.accent2 });
    T(s, c[1], { x: x + 0.4, y: 2.65, w: cw - 0.8, h: 1.7, fontSize: 25, fontFace: "Pretendard Black", color: C.background1, lineSpacingMultiple: 1.15 });
    T(s, c[2], { x: x + 0.4, y: 4.55, w: cw - 0.8, h: 1.9, fontSize: 13.5, color: i === 0 ? C.background1 : C.accent4, lineSpacingMultiple: 1.3 });
  });
}

// ───────── 18. SELF SCORE
sec("평가");
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "평가" });
  head(s, "SELF-SCORE · 본 기획안", "92점 — 남은 8점을 채우는 방법");
  tag(s, 12.1, 0.43, "추론");
  const g = [
    ["−3", "나문희 배우 참여 미확정", "킬러 콘텐츠가 본인 동의를 전제로 함", ["심영은 배우 EP.000 파일럿 → 반응 데이터로 제안", "참여 부담 최소화: 대본 선정 + 손편지 코멘트, 촬영 1회", "수익 일부를 신인 장학금으로 — 참여 명분 강화"]],
    ["−3", "국내 1차 시장 데이터 부족", "지망생 수·연기학원 시장 공식 통계 미확인", ["연극영화과·연기학원 300명 설문 (2주)", "독립영화 감독 20명 FGI로 지불의향 검증", "예술인복지재단 예술활동증명 통계 정보공개청구"]],
    ["−2", "법률 구조 미확정", "알선 수수료·직업소개 해당 여부 해석 필요", ["엔터 전문 변호사 사전 자문 (출시 전 필수)", "초기 수수료 0원 구조로 리스크 회피", "콘진원·규제샌드박스 상담 채널 활용"]],
  ];
  const cw = (W - 2 * M - 2 * 0.3) / 3;
  g.forEach((x, i) => {
    const X = M + i * (cw + 0.3);
    card(s, X, 1.85, cw, 4.8);
    T(s, x[0], { x: X + 0.35, y: 2.05, w: 1.2, h: 0.6, fontSize: 30, fontFace: "Pretendard Black", color: C.accent1 });
    T(s, x[1], { x: X + 0.35, y: 2.75, w: cw - 0.7, h: 0.4, fontSize: 16, bold: true });
    T(s, "이유: " + x[2], { x: X + 0.35, y: 3.2, w: cw - 0.7, h: 0.6, fontSize: 12.5, color: C.accent3 });
    x[3].forEach((t, k) => {
      const y = 3.95 + k * 0.88;
      s.addShape(pres.ShapeType.ellipse, { x: X + 0.35, y: y + 0.02, w: 0.34, h: 0.34, fill: { color: k === 0 ? C.accent1 : C.text1 }, line: { type: "none" } });
      T(s, String(k + 1), { x: X + 0.35, y: y + 0.02, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
      T(s, t, { x: X + 0.85, y, w: cw - 1.2, h: 0.75, fontSize: 12.5, bold: k === 0, lineSpacingMultiple: 1.15 });
    });
  });
}

// ───────── 19. SOURCES
{
  const s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "평가" });
  head(s, "SOURCES", "‘사실’ 표기 근거 자료");
  const src = [
    "파이낸셜뉴스(2026.11) — 실연자 사용료 징수액 2024 약 470억 → 2025 약 340억 원",
    "더팩트(2026) — 드라마 업계 12개 단체, 제작 편수 회복 정부 지원 촉구",
    "아시아경제(2020.12) — 서울시·한국방송실연자협회 방송배우 실태조사 (연 1,000만 원 미만 80%)",
    "한국소비자원 / 코리안센터 — 학원형 연예기획사 피해구제 366건, 상반기 +75%",
    "세계일보(2025.9) — 문체부 기획업 일제 등록 계도기간 및 미등록 엄정 조치 방침",
    "ET뉴스 · 데일리안(2026) — 글로벌 숏폼드라마 시장, 20대 숏드라마 이용률 27.0%",
    "혁신의숲 · 스포츠경향(2024.8) — 플필 등록 배우 약 1.67만 명, 신인 오디션 3,000명+ 지원",
    "머니투데이 유니콘팩토리(2026.5) — 베피플·캐스틱 캐스팅 사업 사례",
    "TA Associates 보도자료 — Backstage 유료 구독자 25만+, StarNow·Mandy 인수",
    "TheWrap · Authors Guild — SAG-AFTRA AI 디지털 복제 동의·보상 조항",
  ];
  T(s, src.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < src.length - 1 } })), { x: M, y: 1.85, w: W - 2 * M, h: 4.4, fontSize: 13.5, paraSpaceAfter: 6, color: C.text1 });
  T(s, "※ 수치는 보도 시점 기준이며, 투자·대외 제출 전 원문 재확인을 권장합니다.", { x: M, y: 6.3, w: 12, h: 0.35, fontSize: 11.5, italic: true, color: C.accent3 });
}

// ───────── 20. CLOSING + RIGHTS
sec("마무리");
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "마무리" });
  s.addImage({ path: A + "mark_dark_t.png", x: W / 2 - 0.75, y: 0.7, w: 1.5, h: 1.5 });
  T(s, "한 장면이면 충분하다.", { x: M, y: 2.35, w: W - 2 * M, h: 1.0, fontSize: 44, fontFace: "Pretendard Black", color: C.background1, align: "center", valign: "middle" });
  T(s, "ACT·MATCH  by  WEUM CORPORATION", { x: M, y: 3.35, w: W - 2 * M, h: 0.4, fontSize: 14, bold: true, color: C.accent1, align: "center", charSpacing: 3 });
  card(s, 1.4, 4.35, W - 2.8, 2.35, C.accent5);
  T(s, "권리 고지 · NOTICE OF RIGHTS", { x: 1.8, y: 4.55, w: W - 3.6, h: 0.3, fontSize: 12, bold: true, color: C.accent2, charSpacing: 2 });
  T(s, "본 기획안에 포함된 사업 아이디어, 서비스 구조, 브랜드명(ACT MATCH), 로고, 카피, 콘텐츠 포맷(LEGEND SCRIPT · SAFE CAST 등), 수익모델 및 모든 구성에 관한 저작권과 일체의 권리는 WEUM CORPORATION에 있습니다. 사전 서면 동의 없는 복제·배포·전송·변형·2차적 저작물 작성 및 영업상 이용을 금하며, 본 자료는 「저작권법」 및 「부정경쟁방지 및 영업비밀보호에 관한 법률」에 따라 보호되는 영업비밀입니다.",
    { x: 1.8, y: 4.95, w: W - 3.6, h: 1.2, fontSize: 11.5, color: C.accent4, lineSpacingMultiple: 1.3 });
  T(s, "© 2026 WEUM CORPORATION. All Rights Reserved.", { x: 1.8, y: 6.2, w: W - 3.6, h: 0.3, fontSize: 12, bold: true, color: C.background1 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written");
})();
