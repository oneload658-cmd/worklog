const sharp = require('sharp');
const OUT = '/home/user/worklog/act-match/assets/';
const Y = '#FFE14D';
function mark(ink, ox = 0, oy = 0, s = 1) {
  // two overlapping scene frames; overlap = highlighted "match"
  return `<g transform="translate(${ox},${oy}) scale(${s})">
    <rect x="140" y="160" width="120" height="80" fill="${Y}"/>
    <rect x="50" y="70" width="210" height="170" rx="26" fill="none" stroke="${ink}" stroke-width="26"/>
    <rect x="140" y="160" width="210" height="170" rx="26" fill="none" stroke="${ink}" stroke-width="26"/>
  </g>`;
}
async function render(svg, file, w) { await sharp(Buffer.from(svg)).resize({ width: w }).png().toFile(OUT + file); }
(async () => {
  await render(`<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400" viewBox="0 0 400 400">${mark('#141414')}</svg>`, 'mark_light.png', 800);
  await render(`<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400" viewBox="0 0 400 400">${mark('#FFFFFF')}</svg>`, 'mark_dark.png', 800);
  // app icon
  await render(`<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400" viewBox="0 0 400 400"><rect width="400" height="400" rx="88" fill="#141414"/>${mark('#FFFFFF', 60, 60, 0.7)}</svg>`, 'app_icon.png', 800);
  for (const [name, ink, sub] of [['lockup_light', '#141414', '#7A7A7A'], ['lockup_dark', '#FFFFFF', '#A8A8A8']]) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1500" height="400" viewBox="0 0 1500 400">
      ${mark(ink)}
      <rect x="742" y="182" width="${ink === '#FFFFFF' ? 0 : 0}" height="0"/>
      ${ink === "#FFFFFF" ? `<rect x="757" y="112" width="588" height="150" fill="${Y}"/>` : `<rect x="762" y="178" width="578" height="70" fill="${Y}"/>`}
      <text x="400" y="240" font-family="Pretendard Black" font-weight="900" font-size="150" letter-spacing="2" fill="${ink}">ACT</text>
      <text x="772" y="240" font-family="Pretendard Black" font-weight="900" font-size="150" letter-spacing="2" fill="${ink === "#FFFFFF" ? "#141414" : ink}">MATCH</text>
      <text x="404" y="322" font-family="Liberation Mono" font-size="34" letter-spacing="4" fill="${sub}">SCENE-FIRST CASTING · WEUM</text></svg>`;
    await render(svg, name + '.png', 2250);
  }
})();
