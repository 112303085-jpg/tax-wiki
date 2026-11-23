"""Generate a consistent SVG element kit for presentation slides."""
import math, os, sys

OUT = sys.argv[1]
P = dict(primary="#2D5BFF", violet="#7B61FF", orange="#FF8A3D", teal="#14B8A6",
         pink="#F43F5E", ink="#1E293B", muted="#64748B", light="#EEF2FF", white="#FFFFFF")
SERIES = [P["primary"], P["violet"], P["teal"], P["orange"], P["pink"]]
FONT = "font-family=\"'Noto Sans TC','Microsoft JhengHei','Segoe UI',sans-serif\""

def save(folder, name, w, h, body, vb=None):
    d = os.path.join(OUT, folder); os.makedirs(d, exist_ok=True)
    vb = vb or f"0 0 {w} {h}"
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="{vb}">\n'
           f'{body}\n</svg>\n')
    with open(os.path.join(d, name + ".svg"), "w", encoding="utf-8") as f:
        f.write(svg)

# ---------------- icons (24px grid, line style) ----------------
def icon(name, inner, color=P["primary"]):
    body = (f'<g fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" '
            f'stroke-linejoin="round">{inner}</g>')
    save("icons", name, 256, 256, body, vb="0 0 24 24")

def star_pts(cx, cy, ro, ri, n=5):
    pts = []
    for i in range(n * 2):
        r = ro if i % 2 == 0 else ri
        a = -math.pi / 2 + i * math.pi / n
        pts.append(f"{cx + r*math.cos(a):.2f},{cy + r*math.sin(a):.2f}")
    return " ".join(pts)

def gear():
    teeth = "".join(
        f'<line x1="{12+6.5*math.cos(a):.2f}" y1="{12+6.5*math.sin(a):.2f}" '
        f'x2="{12+9.5*math.cos(a):.2f}" y2="{12+9.5*math.sin(a):.2f}"/>'
        for a in [i*math.pi/4 for i in range(8)])
    return f'<circle cx="12" cy="12" r="6.5"/><circle cx="12" cy="12" r="2.5"/>{teeth}'

ICONS = {
 "01-lightbulb": '<path d="M9 18h6M10 21h4M12 3a6 6 0 0 0-4 10.5c.8.8 1 1.5 1 2.5h6c0-1 .2-1.7 1-2.5A6 6 0 0 0 12 3z"/>',
 "02-target": '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/>',
 "03-chart-growth": '<path d="M3 3v18h18"/><path d="M7 15l4-4 3 3 5-6"/><path d="M15 8h4v4"/>',
 "04-gear": gear(),
 "05-team": '<circle cx="9" cy="8" r="3.5"/><path d="M2 20c0-3.5 3-6 7-6s7 2.5 7 6"/><circle cx="17" cy="9" r="2.5"/><path d="M17 14c2.8 0 5 2 5 5"/>',
 "06-rocket": '<path d="M12 2c3 2 5 6 5 10l-2 4H9l-2-4c0-4 2-8 5-10z"/><circle cx="12" cy="10" r="2"/><path d="M7 12l-3 3v3h3M17 12l3 3v3h-3M10 19l2 3 2-3"/>',
 "07-check-circle": '<circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/>',
 "08-clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
 "09-globe": '<circle cx="12" cy="12" r="9"/><ellipse cx="12" cy="12" rx="4" ry="9"/><path d="M3 12h18"/>',
 "10-lock": '<rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/><circle cx="12" cy="16" r="1"/>',
 "11-book": '<path d="M2 5c3-1 7-1 10 1v14c-3-2-7-2-10-1zM22 5c-3-1-7-1-10 1v14c3-2 7-2 10-1z"/>',
 "12-chat": '<path d="M4 5h16v11H9l-5 4z"/><path d="M8 9h8M8 12h5"/>',
 "13-search": '<circle cx="11" cy="11" r="7"/><path d="M16 16l5 5"/>',
 "14-database": '<ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/><path d="M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3"/>',
 "15-cloud": '<path d="M7 18a5 5 0 0 1-.5-10A6 6 0 0 1 18 9a4.5 4.5 0 0 1-.5 9z"/>',
 "16-document": '<path d="M6 2h8l5 5v15H6z"/><path d="M14 2v5h5M9 13h7M9 17h7"/>',
 "17-calendar": '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
 "18-flag": '<path d="M5 21V3"/><path d="M5 4h12l-2 4 2 4H5"/>',
 "19-star": f'<polygon points="{star_pts(12, 12.5, 9.5, 4.2)}"/>',
 "20-trophy": '<path d="M8 3h8v6a4 4 0 0 1-8 0z"/><path d="M8 5H5a3 3 0 0 0 3 5M16 5h3a3 3 0 0 1-3 5M12 13v4M8 21h8M10 17h4v4"/>',
 "21-mail": '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/>',
 "22-warning": '<path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18v.01"/>',
 "23-money": '<circle cx="12" cy="12" r="9"/><path d="M15 9c-.5-1-1.6-1.5-3-1.5-1.7 0-3 .8-3 2s1.3 1.8 3 2.2 3 1 3 2.3-1.3 2-3 2c-1.4 0-2.6-.5-3-1.5M12 6v1.5M12 16.5V18"/>',
 "24-ai-chip": '<rect x="6" y="6" width="12" height="12" rx="2"/><rect x="9.5" y="9.5" width="5" height="5" rx="1"/><path d="M9 2v4M15 2v4M9 18v4M15 18v4M2 9h4M2 15h4M18 9h4M18 15h4"/>',
}
for n, inner in ICONS.items():
    icon(n, inner)

# ---------------- diagrams ----------------
def text(x, y, s, size=28, color=P["white"], weight=700, anchor="middle"):
    return (f'<text x="{x}" y="{y}" {FONT} font-size="{size}" font-weight="{weight}" '
            f'fill="{color}" text-anchor="{anchor}" dominant-baseline="middle">{s}</text>')

D = "diagrams"
# arrow right
save(D, "01-arrow-right", 600, 160,
     f'<path d="M10 55h430V15l150 65-150 65v-40H10z" fill="{P["primary"]}"/>')
# curved arrow
save(D, "02-arrow-curved", 500, 300,
     f'<path d="M40 260C60 110 200 40 380 70" fill="none" stroke="{P["orange"]}" stroke-width="18" stroke-linecap="round"/>'
     f'<path d="M360 20l90 55-80 60z" fill="{P["orange"]}"/>')
# chevron process 3
ch = ""
for i, c in enumerate(SERIES[:3]):
    x = 10 + i * 330
    notch = "" if i == 0 else f"L{x+50} 100"
    ch += (f'<path d="M{x} 10H{x+300}L{x+350} 100L{x+300} 190H{x}{notch}z" fill="{c}"/>'
           + text(x + 175 + (0 if i == 0 else 25) - 12, 100, f"STEP {i+1}", 34))
save(D, "03-chevron-process-3", 1020, 200, ch)
# timeline 5
tl = f'<line x1="60" y1="150" x2="1140" y2="150" stroke="{P["muted"]}" stroke-width="6" stroke-linecap="round"/>'
for i in range(5):
    x = 120 + i * 240; c = SERIES[i]
    tl += (f'<circle cx="{x}" cy="150" r="34" fill="{c}"/>' + text(x, 151, f"{i+1:02d}", 26)
           + f'<line x1="{x}" y1="{100 if i%2==0 else 200}" x2="{x}" y2="{60 if i%2==0 else 240}" stroke="{c}" stroke-width="4"/>'
           + text(x, 40 if i % 2 == 0 else 262, f"Phase {i+1}", 26, P["ink"]))
save(D, "04-timeline-5", 1200, 300, tl)
# cycle 4
cy4 = ""
cx, cyy, r = 300, 300, 200
for i in range(4):
    a0 = math.radians(-90 + i * 90 + 12); a1 = math.radians(-90 + (i + 1) * 90 - 12)
    x0, y0 = cx + r*math.cos(a0), cyy + r*math.sin(a0)
    x1, y1 = cx + r*math.cos(a1), cyy + r*math.sin(a1)
    cy4 += (f'<path d="M{x0:.1f} {y0:.1f}A{r} {r} 0 0 1 {x1:.1f} {y1:.1f}" fill="none" stroke="{SERIES[i]}" '
            f'stroke-width="26" stroke-linecap="round"/>')
    # arrowhead at end
    t = a1 + math.pi/2
    hx, hy = x1 + 22*math.cos(t), y1 + 22*math.sin(t)
    lx, ly = x1 + 26*math.cos(a1), y1 + 26*math.sin(a1)
    ix, iy = x1 - 26*math.cos(a1), y1 - 26*math.sin(a1)
    cy4 += f'<polygon points="{hx:.1f},{hy:.1f} {lx:.1f},{ly:.1f} {ix:.1f},{iy:.1f}" fill="{SERIES[i]}"/>'
    am = math.radians(-45 + i * 90)
    cy4 += (f'<circle cx="{cx + r*math.cos(am):.1f}" cy="{cyy + r*math.sin(am):.1f}" r="44" fill="{P["white"]}" '
            f'stroke="{SERIES[i]}" stroke-width="6"/>' + text(cx + r*math.cos(am), cyy + r*math.sin(am), str(i+1), 34, SERIES[i]))
cy4 += text(cx, cyy, "CYCLE", 40, P["ink"])
save(D, "05-cycle-4", 600, 600, cy4)
# venn 2 / 3
save(D, "06-venn-2", 700, 460,
     f'<circle cx="250" cy="230" r="200" fill="{P["primary"]}" fill-opacity=".55"/>'
     f'<circle cx="450" cy="230" r="200" fill="{P["orange"]}" fill-opacity=".55"/>'
     + text(170, 230, "A", 48) + text(530, 230, "B", 48) + text(350, 230, "A∩B", 32))
save(D, "07-venn-3", 640, 600,
     f'<circle cx="230" cy="220" r="180" fill="{P["primary"]}" fill-opacity=".5"/>'
     f'<circle cx="410" cy="220" r="180" fill="{P["violet"]}" fill-opacity=".5"/>'
     f'<circle cx="320" cy="380" r="180" fill="{P["teal"]}" fill-opacity=".5"/>'
     + text(170, 170, "A", 44) + text(470, 170, "B", 44) + text(320, 480, "C", 44))
# pyramid 4
py = ""
for i in range(4):
    yt, yb = 20 + i * 130, 20 + (i + 1) * 130 - 10
    hw_t, hw_b = 15 + i * 87.5 * 1.0, 15 + (i + 1) * 87.5 - 8
    hw_t = (yt - 20) / 520 * 330; hw_b = (yb - 20) / 520 * 330
    py += f'<path d="M{350-hw_t:.1f} {yt}H{350+hw_t:.1f}L{350+hw_b:.1f} {yb}H{350-hw_b:.1f}z" fill="{SERIES[i]}"/>'
    py += text(350, (yt + yb) / 2 + (15 if i == 0 else 0), f"L{4-i}", 30)
save(D, "08-pyramid-4", 700, 560, py)
# funnel 4
fn = ""
for i in range(4):
    yt, yb = 20 + i * 120, 20 + (i + 1) * 120 - 10
    wt, wb = 330 - i * 70, 330 - (i + 1) * 70 + 10
    fn += f'<path d="M{350-wt} {yt}H{350+wt}L{350+wb} {yb}H{350-wb}z" fill="{SERIES[i]}"/>' + text(350, (yt+yb)/2, f"Stage {i+1}", 28)
save(D, "09-funnel-4", 700, 500, fn)
# 2x2 matrix
mx = ""
labels = ["Quick Wins", "Major Projects", "Fill-ins", "Thankless"]
for i in range(4):
    x = 80 + (i % 2) * 300; y = 20 + (i // 2) * 250
    mx += f'<rect x="{x}" y="{y}" width="290" height="240" rx="18" fill="{SERIES[i]}" fill-opacity="{1 if i==0 else .85}"/>' + text(x+145, y+120, labels[i], 30)
mx += (f'<path d="M50 540V20M50 540H690" stroke="{P["ink"]}" stroke-width="4" fill="none"/>'
       f'<path d="M50 8l-10 20h20zM702 540l-20-10v20z" fill="{P["ink"]}"/>'
       + text(30, 280, "Impact", 24, P["ink"], 600).replace('<text', '<text transform="rotate(-90 30 280)"')
       + text(380, 572, "Effort", 24, P["ink"], 600))
save(D, "10-matrix-2x2", 720, 590, mx)
# hub & spoke
hs = ""
for i in range(6):
    a = math.radians(-90 + i * 60); x, y = 350 + 220*math.cos(a), 300 + 220*math.sin(a)
    hs += f'<line x1="350" y1="300" x2="{x:.1f}" y2="{y:.1f}" stroke="{P["muted"]}" stroke-width="4" stroke-dasharray="10 8"/>'
for i in range(6):
    a = math.radians(-90 + i * 60); x, y = 350 + 220*math.cos(a), 300 + 220*math.sin(a)
    c = SERIES[i % 5]
    hs += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="58" fill="{c}"/>' + text(x, y, str(i+1), 36)
hs += f'<circle cx="350" cy="300" r="95" fill="{P["ink"]}"/>' + text(350, 300, "CORE", 36)
save(D, "11-hub-spoke-6", 700, 600, hs)
# stairs
st = ""
for i in range(4):
    x = 20 + i * 200; h = 100 + i * 90
    st += f'<rect x="{x}" y="{440-h}" width="190" height="{h}" rx="10" fill="{SERIES[i]}"/>' + text(x+95, 440-h+45, f"0{i+1}", 40)
st += f'<path d="M60 70Q420 -10 780 30" fill="none" stroke="{P["orange"]}" stroke-width="6" stroke-dasharray="14 10"/>'
save(D, "12-growth-stairs", 820, 450, st)

# ---------------- decorative shapes ----------------
S = "shapes"
save(S, "01-blob-blue", 500, 500,
     f'<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="{P["primary"]}"/>'
     f'<stop offset="1" stop-color="{P["violet"]}"/></linearGradient></defs>'
     '<path d="M413 98c53 61 69 160 30 236s-133 129-221 120S53 380 35 293 52 117 125 67 360 37 413 98z" fill="url(#g)"/>')
save(S, "02-blob-orange", 500, 500,
     f'<defs><linearGradient id="g" x1="0" y1="1" x2="1" y2="0"><stop offset="0" stop-color="{P["orange"]}"/>'
     f'<stop offset="1" stop-color="{P["pink"]}"/></linearGradient></defs>'
     '<path d="M390 70c70 40 95 140 70 225s-110 170-205 165S70 385 45 300 40 105 120 60s200-30 270 10z" fill="url(#g)"/>')
save(S, "03-wave-divider", 1920, 200,
     f'<path d="M0 90C320 10 640 170 960 100s640-90 960 0V200H0z" fill="{P["primary"]}" fill-opacity=".25"/>'
     f'<path d="M0 130C320 60 640 200 960 140s640-80 960-10V200H0z" fill="{P["primary"]}"/>')
dots = "".join(f'<circle cx="{15+x*30}" cy="{15+y*30}" r="4"/>' for x in range(10) for y in range(10))
save(S, "04-dot-grid", 300, 300, f'<g fill="{P["primary"]}" fill-opacity=".45">{dots}</g>')
save(S, "05-corner-accent", 400, 400,
     f'<path d="M0 0H400A400 400 0 0 1 0 400z" fill="{P["primary"]}"/>'
     f'<path d="M0 0H260A260 260 0 0 1 0 260z" fill="{P["violet"]}"/>'
     f'<path d="M0 0H130A130 130 0 0 1 0 130z" fill="{P["orange"]}"/>')
save(S, "06-concentric-rings", 500, 500,
     "".join(f'<circle cx="250" cy="250" r="{r}" fill="none" stroke="{P["primary"]}" stroke-width="3" stroke-opacity="{o}"/>'
             for r, o in [(60, .9), (110, .7), (160, .5), (210, .3), (245, .15)]))
save(S, "07-gradient-orb", 500, 500,
     f'<defs><radialGradient id="g" cx=".35" cy=".35" r=".7"><stop offset="0" stop-color="#9DB4FF"/>'
     f'<stop offset=".55" stop-color="{P["primary"]}"/><stop offset="1" stop-color="{P["violet"]}"/></radialGradient></defs>'
     '<circle cx="250" cy="250" r="240" fill="url(#g)"/>')
save(S, "08-diagonal-stripes", 300, 300,
     f'<defs><clipPath id="c"><rect width="300" height="300" rx="24"/></clipPath></defs><g clip-path="url(#c)" stroke="{P["orange"]}" stroke-width="10">'
     + "".join(f'<line x1="{i}" y1="0" x2="{i-300}" y2="300"/>' for i in range(0, 620, 30)) + '</g>')
save(S, "09-brush-underline", 600, 80,
     f'<path d="M12 50C120 22 300 18 588 30c-120 4-330 14-560 40z" fill="{P["orange"]}"/>'
     f'<path d="M40 60C200 44 380 40 560 44" fill="none" stroke="{P["orange"]}" stroke-width="6" stroke-linecap="round" opacity=".6"/>')
hexes = ""
for row in range(5):
    for col in range(6):
        cxh = 40 + col * 69 + (34.5 if row % 2 else 0); cyh = 40 + row * 60
        pts = " ".join(f"{cxh+36*math.cos(math.radians(30+60*k)):.1f},{cyh+36*math.sin(math.radians(30+60*k)):.1f}" for k in range(6))
        hexes += f'<polygon points="{pts}"/>'
save(S, "10-hexagon-pattern", 460, 320, f'<g fill="none" stroke="{P["teal"]}" stroke-width="3" stroke-opacity=".6">{hexes}</g>')

# ---------------- callouts / labels ----------------
C = "callouts"
save(C, "01-speech-bubble", 500, 360,
     f'<path d="M40 20h420a30 30 0 0 1 30 30v190a30 30 0 0 1-30 30H180l-90 70 15-70H40a30 30 0 0 1-30-30V50a30 30 0 0 1 30-30z" fill="{P["primary"]}"/>'
     + text(250, 145, "Your message", 40))
save(C, "02-quote-mark", 300, 240,
     f'<g fill="{P["violet"]}"><path d="M20 230V140C20 70 50 25 120 10l14 30c-40 14-55 40-56 80h52v110z"/>'
     f'<path d="M160 230V140C160 70 190 25 260 10l14 30c-40 14-55 40-56 80h52v110z"/></g>')
save(C, "03-ribbon-label", 600, 140,
     f'<path d="M0 30h60v100H0l30-50z" fill="#1B3BB8"/><path d="M600 30h-60v100h60l-30-50z" fill="#1B3BB8"/>'
     f'<rect x="40" y="10" width="520" height="100" fill="{P["primary"]}"/>'
     f'<path d="M40 110l20 20v-20zM560 110l-20 20v-20z" fill="#0F2580"/>' + text(300, 62, "HIGHLIGHT", 44))
save(C, "04-number-badge", 200, 200,
     f'<circle cx="100" cy="100" r="92" fill="{P["orange"]}"/><circle cx="100" cy="100" r="76" fill="none" stroke="{P["white"]}" stroke-width="4" stroke-dasharray="6 8"/>'
     + text(100, 104, "01", 72))
save(C, "05-sticky-note", 400, 400,
     '<defs><filter id="s" x="-10%" y="-10%" width="130%" height="130%"><feDropShadow dx="4" dy="8" stdDeviation="8" flood-opacity=".18"/></filter></defs>'
     '<path d="M30 30h340v260l-80 80H30z" fill="#FFE37A" filter="url(#s)"/><path d="M370 290h-80v80z" fill="#E8C24A"/>'
     f'<rect x="150" y="14" width="100" height="32" rx="4" fill="{P["teal"]}" fill-opacity=".7"/>'
     + text(200, 190, "NOTE", 52, P["ink"]))

# ---------------- backgrounds (16:9) ----------------
B = "backgrounds"
W, H = 1920, 1080
save(B, "01-title-gradient", W, H,
     f'<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1B2B6B"/>'
     f'<stop offset=".6" stop-color="{P["primary"]}"/><stop offset="1" stop-color="{P["violet"]}"/></linearGradient></defs>'
     f'<rect width="{W}" height="{H}" fill="url(#g)"/>'
     f'<circle cx="1650" cy="180" r="380" fill="{P["white"]}" fill-opacity=".07"/>'
     f'<circle cx="1750" cy="900" r="260" fill="{P["orange"]}" fill-opacity=".5"/>'
     f'<circle cx="200" cy="1000" r="200" fill="{P["white"]}" fill-opacity=".05"/>')
save(B, "02-section-divider", W, H,
     f'<rect width="{W}" height="{H}" fill="{P["light"]}"/>'
     f'<path d="M0 0H800L560 1080H0z" fill="{P["primary"]}"/><path d="M800 0H880L640 1080H560z" fill="{P["orange"]}"/>')
save(B, "03-content-light", W, H,
     f'<rect width="{W}" height="{H}" fill="#FAFBFF"/>'
     f'<rect x="0" y="0" width="{W}" height="12" fill="{P["primary"]}"/>'
     f'<path d="M1920 1080H1500A420 420 0 0 1 1920 660z" fill="{P["primary"]}" fill-opacity=".08"/>'
     f'<path d="M1920 1080H1700A220 220 0 0 1 1920 860z" fill="{P["primary"]}" fill-opacity=".15"/>'
     f'<rect x="96" y="1020" width="160" height="6" rx="3" fill="{P["orange"]}"/>')
grid = "".join(f'<path d="M{x} 0V{H}"/>' for x in range(0, W + 1, 80)) + "".join(f'<path d="M0 {y}H{W}"/>' for y in range(0, H + 1, 80))
save(B, "04-dark-grid", W, H,
     f'<defs><radialGradient id="g" cx=".75" cy=".3" r=".8"><stop offset="0" stop-color="#1E2A5A"/><stop offset="1" stop-color="#0B1020"/></radialGradient></defs>'
     f'<rect width="{W}" height="{H}" fill="url(#g)"/><g stroke="#FFFFFF" stroke-opacity=".05" stroke-width="1">{grid}</g>'
     f'<circle cx="1450" cy="320" r="300" fill="{P["violet"]}" fill-opacity=".25"/>')
tri = ""
cols = [P["primary"], P["violet"], P["teal"], P["orange"]]
k = 0
for row in range(4):
    for col in range(6):
        x = 1200 + col * 120; y = row * 270
        if (row + col) % 3 == 0: continue
        tri += f'<path d="M{x} {y}h120v120z" fill="{cols[k%4]}" fill-opacity="{.15 + (k%3)*.2:.2f}"/>'; k += 1
save(B, "05-geometric-light", W, H, f'<rect width="{W}" height="{H}" fill="{P["white"]}"/>' + tri
     + f'<rect x="120" y="140" width="12" height="160" fill="{P["primary"]}"/>')

print("done")
