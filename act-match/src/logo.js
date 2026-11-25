const sharp = require('sharp');
const OUT = '/home/user/worklog/act-match/assets/';
const flameDefs = `
<defs>
  <linearGradient id="fl" x1="0" y1="1" x2="0" y2="0">
    <stop offset="0" stop-color="#E8432E"/><stop offset="0.55" stop-color="#F26B2A"/><stop offset="1" stop-color="#F7B23B"/>
  </linearGradient>
  <linearGradient id="core" x1="0" y1="1" x2="0" y2="0">
    <stop offset="0" stop-color="#FFE3A3"/><stop offset="1" stop-color="#FFF6E0"/>
  </linearGradient>
</defs>`;
function mark(ink){ // 400x400 box
  return `
  <g stroke-linecap="round">
    <line x1="112" y1="372" x2="193" y2="158" stroke="${ink}" stroke-width="24"/>
    <line x1="288" y1="372" x2="207" y2="158" stroke="${ink}" stroke-width="24"/>
    <line x1="152" y1="282" x2="248" y2="282" stroke="${ink}" stroke-width="16"/>
  </g>
  <path d="M200,12 C214,52 250,82 250,124 C250,154 227,176 200,176 C173,176 150,154 150,124 C150,96 172,80 180,56 C187,78 195,88 206,92 C204,66 195,40 200,12 Z" fill="url(#fl)"/>
  <path d="M200,98 C208,116 226,128 226,146 C226,160 214,170 200,170 C186,170 174,160 174,146 C174,130 190,120 200,98 Z" fill="url(#core)"/>`;
}
async function render(svg, file, w){ await sharp(Buffer.from(svg)).resize({width:w}).png().toFile(OUT+file); }
(async()=>{
  for (const [name, ink, bg] of [['mark_dark','#FFFFFF','#111111'],['mark_light','#111111',null]]) {
    await render(`<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400" viewBox="0 0 400 400">${flameDefs}${bg?`<rect width="400" height="400" fill="${bg}"/>`:''}${mark(ink)}</svg>`, name+'.png', 800);
    await render(`<svg xmlns="http://www.w3.org/2000/svg" width="400" height="400" viewBox="0 0 400 400">${flameDefs}${mark(ink)}</svg>`, name+'_t.png', 800);
  }
  // horizontal lockup, transparent
  for (const [name, ink, sub] of [['lockup_dark','#FFFFFF','#B9B4AE'],['lockup_light','#111111','#6B6660']]) {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1400" height="400" viewBox="0 0 1400 400">${flameDefs}
    <g transform="translate(0,0)">${mark(ink)}</g>
    <text x="400" y="250" font-family="Pretendard Black" font-weight="900" font-size="150" letter-spacing="6" fill="${ink}">ACT<tspan fill="#E8432E">·</tspan>MATCH</text>
    <text x="406" y="330" font-family="Pretendard Medium" font-weight="500" font-size="38" letter-spacing="15" fill="${sub}">BY WEUM CORPORATION</text></svg>`;
    await render(svg, name+'.png', 2100);
  }
})();
