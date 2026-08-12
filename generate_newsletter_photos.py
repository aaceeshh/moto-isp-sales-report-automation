#!/usr/bin/env python3
"""
generate_newsletter_photos.py
-----------------------------
Renders the 5 leaderboard sections of the East Reliance Sales Newsletter as
five redesigned PNG cards and posts each as a Telegram photo (sendPhoto).

Runs daily at 9:00 AM IST via GitHub Actions.

Design goal: a fresh, phone-friendly leaderboard card per category — NOT the
website's look. To guarantee the numbers always match the live newsletter, the
data is scraped straight from the rendered worker page (the worker owns all the
ranking math), then re-skinned here.

Flow:
  1. headless Chromium loads NEWSLETTER_URL, waits for the leaderboard to render
  2. a small JS extractor returns the 5 sections as structured JSON
  3. each section is rendered to a standalone HTML card and screenshotted (2x)
  4. each PNG is sent to Telegram as a photo, in order 1..5

Env vars:
  NEWSLETTER_URL        default: the East Reliance newsletter worker
  TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
Optional:
  PREVIEW_JSON          path to a saved sections JSON -> render locally, no
                        browser fetch (used for offline design preview)
  OUT_DIR               where PNGs/HTML are written (default: ./_nl_photos)
  DRY_RUN               "1" -> render PNGs but do not send to Telegram
"""

import os
import re
import json
import html
import hashlib
from datetime import datetime, timezone, timedelta

IST = timezone(timedelta(hours=5, minutes=30))

NEWSLETTER_URL = os.environ.get("NEWSLETTER_URL") or \
    "https://east-reliance-newsletter.ashish-imnu.workers.dev/"
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID   = os.environ.get("TELEGRAM_CHAT_ID", "")
OUT_DIR = os.environ.get("OUT_DIR") or "_nl_photos"

CARD_WIDTH = 1240   # px (Telegram displays photos up to ~1280 wide)
SCALE      = 2      # device scale factor -> crisp text

# Browser-side extractor: reads the worker's rendered DOM into clean JSON.
EXTRACT_JS = r"""
() => {
  const clean = s => (s||"").replace(/\s+/g,' ').trim();
  const strip = (s,suf) => (suf && s.endsWith(suf)) ? s.slice(0,s.length-suf.length).trim() : s;
  return [...document.querySelectorAll('.cat-section')].map(sec => {
    const num = clean(sec.querySelector('.num')?.textContent);
    let title = clean(sec.querySelector('.section-title')?.textContent);
    const range = clean(sec.querySelector('.range')?.textContent);
    if (num && title.startsWith(num)) title = title.slice(num.length).trim();
    if (range && title.endsWith(range)) title = title.slice(0,title.length-range.length).trim();
    const blocks = [...sec.querySelectorAll('.block')].map(bl => {
      let tag = clean(bl.querySelector('.tag, .block-head')?.textContent);
      const m = tag.match(/(Top\s*\d+)\s*$/i);
      let topN=''; if(m){ topN=m[1]; tag=tag.slice(0,tag.length-m[0].length).trim(); }
      const rows = [...bl.querySelectorAll('.isd')].map(r => {
        const ap = clean(r.querySelector('.apbadge')?.textContent);
        let name = strip(clean(r.querySelector('.nm-text, .nm')?.textContent), ap);
        let metric = clean(r.querySelector('.metric')?.textContent), value=metric, units='';
        const mm = metric.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
        if(mm){ value=mm[1].trim(); units=mm[2].trim(); }
        return { rank: clean(r.querySelector('.rank')?.textContent), name,
                 store: clean(r.querySelector('.store')?.textContent),
                 loc: clean(r.querySelector('.loc')?.textContent), value, units, ap };
      });
      return { tag, topN, rows };
    });
    return { num, title, range, blocks };
  });
}
"""

# ── Rendering helpers ────────────────────────────────────────────────────────

AV_PALETTE = ["#3b82f6", "#8b5cf6", "#ec4899", "#f59e0b", "#10b981",
              "#06b6d4", "#ef4444", "#6366f1", "#14b8a6", "#f97316"]

def esc(s):
    return html.escape(str(s or ""), quote=True)

def initials(name):
    parts = [p for p in re.split(r"\s+", (name or "").strip()) if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()

def av_color(name):
    h = int(hashlib.md5((name or "").encode("utf-8")).hexdigest(), 16)
    return AV_PALETTE[h % len(AV_PALETTE)]

def medal_class(rank):
    return {"1": "g1", "2": "g2", "3": "g3"}.get(str(rank).strip(), "gn")

def render_row(r):
    ap = ('<span class="ap">%s</span>' % esc(r.get("ap"))) if r.get("ap") else ""
    units = ('<div class="u">%s</div>' % esc(r.get("units"))) if r.get("units") else ""
    return (
        '<div class="row">'
        '<div class="medal %s">%s</div>'
        '<div class="av" style="background:%s">%s</div>'
        '<div class="info"><div class="nm">%s%s</div>'
        '<div class="sub">%s%s%s</div></div>'
        '<div class="val"><div class="v">%s</div>%s</div>'
        '</div>'
    ) % (
        medal_class(r.get("rank")), esc(r.get("rank")),
        av_color(r.get("name")), esc(initials(r.get("name"))),
        esc(r.get("name")), ap,
        esc(r.get("store")), " · " if r.get("store") and r.get("loc") else "",
        esc(r.get("loc")),
        esc(r.get("value")), units,
    )

def render_block(b):
    rows = "".join(render_row(r) for r in b.get("rows", []))
    pill = ('<span class="ppill">%s</span>' % esc(b.get("topN"))) if b.get("topN") else ""
    return (
        '<div class="panel"><div class="phead"><div class="ptag">%s</div>%s</div>'
        '<div class="rows">%s</div></div>'
    ) % (esc(b.get("tag")), pill, rows)

def render_card(section, gen_str):
    from string import Template
    blocks = "".join(render_block(b) for b in section.get("blocks", []))
    return Template(CARD_TEMPLATE).substitute(
        w=CARD_WIDTH,
        num=esc(section.get("num")),
        title=esc(section.get("title")),
        rng=esc(section.get("range")),
        blocks=blocks,
        gen=esc(gen_str),
    )

CARD_TEMPLATE = r"""<!doctype html><html><head><meta charset="utf-8"><style>
* { margin:0; padding:0; box-sizing:border-box; }
html,body { background:#070b16; }
.card {
  width:${w}px; padding:44px 46px 34px;
  font-family:-apple-system,"Segoe UI",Roboto,"Noto Sans",Arial,sans-serif;
  color:#eef2fb;
  background:
    radial-gradient(1200px 500px at 88% -12%, rgba(59,130,246,.28), transparent 60%),
    radial-gradient(900px 500px at -8% 8%, rgba(139,92,246,.20), transparent 55%),
    linear-gradient(160deg,#0b1120 0%,#0d1526 55%,#0a0f1e 100%);
  background-color:#0b1120;
}
.mast { display:flex; align-items:center; justify-content:space-between;
  padding-bottom:22px; border-bottom:1px solid rgba(255,255,255,.09); margin-bottom:26px; }
.brand { display:flex; align-items:center; gap:12px; font-weight:800; letter-spacing:.5px;
  font-size:20px; color:#dbe3f4; }
.mlogo { width:38px; height:38px; border-radius:10px; display:grid; place-items:center;
  font-size:22px; font-weight:900; color:#fff;
  background:linear-gradient(135deg,#3b82f6,#8b5cf6); box-shadow:0 6px 18px rgba(59,130,246,.4); }
.brand .dim { color:#93a2c4; font-weight:600; }
.edition { font-size:14px; font-weight:700; letter-spacing:2px; text-transform:uppercase;
  color:#8ea3cf; background:rgba(255,255,255,.06); padding:8px 14px; border-radius:999px;
  border:1px solid rgba(255,255,255,.08); }
.hero { display:flex; align-items:center; gap:22px; margin-bottom:28px; }
.numbadge { width:86px; height:86px; border-radius:22px; flex:none; display:grid; place-items:center;
  font-size:46px; font-weight:900; color:#fff;
  background:linear-gradient(140deg,#2563eb,#7c3aed); box-shadow:0 12px 30px rgba(76,29,149,.5);
  border:1px solid rgba(255,255,255,.18); }
.htext { min-width:0; }
.title { font-size:42px; line-height:1.08; font-weight:900; letter-spacing:-.5px; color:#f7f9ff; }
.range { margin-top:8px; font-size:19px; font-weight:600; color:#9db0d8; }
.grid { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:20px; }
.panel { min-width:0; background:rgba(255,255,255,.045); border:1px solid rgba(255,255,255,.08);
  border-radius:20px; padding:18px 20px 8px; }
.phead { display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;
  padding-bottom:12px; border-bottom:1px solid rgba(255,255,255,.07); }
.ptag { font-size:21px; font-weight:800; color:#eaf1ff; }
.ppill { font-size:12px; font-weight:800; letter-spacing:1px; text-transform:uppercase;
  color:#0b1120; background:linear-gradient(135deg,#67e8f9,#3b82f6); padding:5px 11px; border-radius:999px; }
.rows { display:flex; flex-direction:column; }
.row { display:flex; align-items:center; gap:14px; padding:12px 2px;
  border-bottom:1px solid rgba(255,255,255,.055); }
.row:last-child { border-bottom:none; }
.medal { width:36px; height:36px; flex:none; border-radius:50%; display:grid; place-items:center;
  font-size:16px; font-weight:900; color:#0b1120; }
.medal.g1 { background:linear-gradient(140deg,#fde68a,#f59e0b); }
.medal.g2 { background:linear-gradient(140deg,#e5e7eb,#9aa7b5); }
.medal.g3 { background:linear-gradient(140deg,#f0b27a,#c77b3e); color:#fff; }
.medal.gn { background:rgba(255,255,255,.10); color:#aebbda; }
.av { width:46px; height:46px; flex:none; border-radius:50%; display:grid; place-items:center;
  font-size:17px; font-weight:800; color:#fff; box-shadow:inset 0 0 0 1px rgba(255,255,255,.15); }
.info { flex:1; min-width:0; }
.nm { font-size:21px; font-weight:800; color:#f4f7ff; white-space:nowrap; overflow:hidden;
  text-overflow:ellipsis; display:flex; align-items:center; gap:8px; }
.ap { font-size:11px; font-weight:800; color:#cfe0ff; background:rgba(99,102,241,.28);
  border:1px solid rgba(129,140,248,.4); padding:2px 7px; border-radius:999px; flex:none; }
.sub { margin-top:3px; font-size:14px; color:#8fa1c6; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.val { flex:none; text-align:right; padding-left:8px; }
.v { font-size:22px; font-weight:900; color:#7dd3fc; letter-spacing:-.3px; white-space:nowrap; }
.u { font-size:13px; font-weight:600; color:#8fa1c6; margin-top:2px; }
.foot { margin-top:26px; padding-top:16px; border-top:1px solid rgba(255,255,255,.08);
  font-size:14px; color:#7c8db3; display:flex; justify-content:space-between; }
.foot b { color:#9db0d8; font-weight:700; }
</style></head><body>
<div class="card">
  <div class="mast">
    <div class="brand"><span class="mlogo">M</span> MOTOROLA <span class="dim">· East Reliance</span></div>
    <div class="edition">Daily Leaderboard</div>
  </div>
  <div class="hero">
    <div class="numbadge">${num}</div>
    <div class="htext"><div class="title">${title}</div><div class="range">🗓 ${rng}</div></div>
  </div>
  <div class="grid">${blocks}</div>
  <div class="foot"><span>${gen}</span><span><b>Motorola East Reliance</b> · Sales Leaderboard</span></div>
</div>
</body></html>"""

# ── Telegram ─────────────────────────────────────────────────────────────────

def send_photo(png_path, caption):
    import urllib.request
    url = "https://api.telegram.org/bot%s/sendPhoto" % TELEGRAM_BOT_TOKEN
    boundary = "----motoNL" + hashlib.md5(png_path.encode()).hexdigest()[:12]
    with open(png_path, "rb") as f:
        img = f.read()
    parts = []
    def field(name, value):
        parts.append(("--%s\r\n" % boundary).encode())
        parts.append(('Content-Disposition: form-data; name="%s"\r\n\r\n' % name).encode())
        parts.append((value + "\r\n").encode())
    field("chat_id", TELEGRAM_CHAT_ID)
    field("caption", caption)
    field("parse_mode", "HTML")
    parts.append(("--%s\r\n" % boundary).encode())
    parts.append(('Content-Disposition: form-data; name="photo"; filename="card.png"\r\n'
                  'Content-Type: image/png\r\n\r\n').encode())
    parts.append(img)
    parts.append("\r\n".encode())
    parts.append(("--%s--\r\n" % boundary).encode())
    body = b"".join(parts)
    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "multipart/form-data; boundary=%s" % boundary})
    resp = urllib.request.urlopen(req, timeout=60)
    print("[NL] sent %s: HTTP %d" % (os.path.basename(png_path), resp.getcode()))

def caption_for(section, gen_str):
    icons = {"1": "\U0001F947", "2": "\U0001F4C8", "3": "\U0001F525",
             "4": "\U0001F4B0", "5": "\U0001F3A7"}
    ic = icons.get(str(section.get("num")), "\U0001F4CA")
    return "%s <b>%s · %s</b>\n%s\n<i>%s</i>" % (
        ic, esc(section.get("num")), esc(section.get("title")),
        esc(section.get("range")), esc(gen_str))

# ── Fetch (Playwright) + main ────────────────────────────────────────────────

def fetch_sections():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox"])
        page = browser.new_context(viewport={"width": 1400, "height": 1000},
                                   device_scale_factor=1).new_page()
        page.goto(NEWSLETTER_URL, wait_until="networkidle", timeout=90000)
        page.wait_for_selector(".cat-section .isd", timeout=60000)
        page.wait_for_timeout(1500)  # let all sections settle
        sections = page.evaluate(EXTRACT_JS)
        browser.close()
    sections = [s for s in sections if s.get("blocks")]
    if len(sections) < 5:
        print("[NL] WARNING: expected 5 sections, got %d" % len(sections))
    return sections

def shoot(sections, gen_str):
    """Render each section card to a PNG; return list of (section, png_path)."""
    from playwright.sync_api import sync_playwright
    os.makedirs(OUT_DIR, exist_ok=True)
    out = []
    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--no-sandbox"])
        ctx = browser.new_context(viewport={"width": CARD_WIDTH, "height": 800},
                                  device_scale_factor=SCALE)
        page = ctx.new_page()
        for s in sections:
            page.set_content(render_card(s, gen_str), wait_until="networkidle")
            el = page.query_selector(".card")
            path = os.path.join(OUT_DIR, "card_%s.png" % s.get("num"))
            el.screenshot(path=path)
            print("[NL] rendered", path)
            out.append((s, path))
        browser.close()
    return out

def main():
    gen_str = "Generated " + datetime.now(IST).strftime("%d %b %Y · %I:%M %p IST")

    preview = os.environ.get("PREVIEW_JSON")
    if preview:
        with open(preview, "r", encoding="utf-8") as f:
            sections = json.load(f)
        print("[NL] preview mode: %d sections from %s" % (len(sections), preview))
    else:
        sections = fetch_sections()
        print("[NL] fetched %d sections from newsletter" % len(sections))

    shots = shoot(sections, gen_str)

    if os.environ.get("DRY_RUN") == "1" or not (TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID):
        print("[NL] DRY_RUN / no token — %d PNGs in %s, not sending." % (len(shots), OUT_DIR))
        return
    for s, path in shots:
        send_photo(path, caption_for(s, gen_str))
    print("[NL] Done — sent %d photos." % len(shots))


if __name__ == "__main__":
    main()
