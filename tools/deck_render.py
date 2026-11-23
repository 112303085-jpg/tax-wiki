"""deck.py render 使用的排版：由 03_逐頁.md 的投影片資料產生 HTML 與 PowerPoint 繪圖指令。

HTML：白底黑字漫畫風（AGENTS.md 第 7.3 節），單一檔案，圖貼路徑 ../../assets/…，自製圖 images/…。
PowerPoint：輸出繪圖指令 JSON，交給 tools/deck_pptx.ps1 透過本機 PowerPoint 繪製。
"""
import html
import re

DISCLAIMER = "本內容為教學整理，資料基準日為{date}，僅供學習參考，不構成個案稅務或法律意見。法規與數額可能修正，實際適用請以財政部及主管機關最新公告為準，或洽詢專業人士。"
ACCENTS = ["purple", "orange", "teal", "blue"]
HEX = {"ink": "111111", "muted": "555555", "blue": "2D5BFF", "purple": "7B61FF", "teal": "14B8A6",
       "orange": "FF8A3D", "white": "FFFFFF", "tint": "F3F5FF", "tealtint": "E8FAF7", "orangetint": "FFF1E8"}


def esc(s):
    return html.escape(s or "", quote=True)


def split_item(s):
    head, _, tail = s.partition("｜")
    return head.strip(), tail.strip()


def asset(a):
    return "../../" + a


def is_diagram(a):
    return a.startswith("assets/diagrams/")


def icons(f):
    return [a for a in (f.get("圖貼") or []) if not is_diagram(a)]


def roc_long(roc):            # 115-11-24 → 115年11月24日
    y, m, d = roc.split("-")
    return f"{int(y)}年{int(m)}月{int(d)}日"


# ====================================================================== HTML

CSS = r"""
:root{--ink:#111;--paper:#fff;--muted:#555;--line:#111;--blue:#2D5BFF;--purple:#7B61FF;--teal:#14B8A6;--orange:#FF8A3D}
*{box-sizing:border-box}html,body{margin:0;height:100%;background:#2A2A2A}
body{font-family:'Noto Sans TC','Microsoft JhengHei','PingFang TC',sans-serif;color:var(--ink)}
.deck{position:fixed;inset:0;overflow:hidden}
.slide{position:absolute;left:50%;top:50%;width:1280px;height:720px;margin:-360px 0 0 -640px;background:var(--paper);overflow:hidden;padding:44px 72px 64px;display:none;transform-origin:center center}
.slide.active{display:block}
h1.t{position:relative;z-index:1;display:inline-block;margin:0 0 10px;font-size:50px;font-weight:900;line-height:1.15;transform:skewX(-6deg)}
h1.t::after{content:"";position:absolute;z-index:-1;left:-14px;right:-18px;bottom:2px;height:22px;background:var(--hl,var(--purple));transform:skewX(-18deg)}
.lead{font-size:24px;font-weight:700;color:var(--muted);margin:0 0 34px;line-height:1.45}
p,li{font-size:26px;line-height:1.55;margin:0}
.muted{color:var(--muted)}.small{font-size:17px;line-height:1.5;color:var(--muted)}
b,strong{font-weight:900}
.panel{position:relative;background:var(--paper);border:4px solid var(--line);box-shadow:8px 8px 0 var(--line);padding:24px 28px}
.clip{position:absolute;inset:0;overflow:hidden;pointer-events:none}
.panel.fill{background:var(--c);color:#fff}
.tag{position:absolute;top:-20px;left:18px;z-index:2;background:var(--c,var(--ink));color:#fff;font-weight:900;font-size:18px;padding:4px 14px;transform:skewX(-10deg);white-space:nowrap}
.bubble{position:relative;display:inline-block;background:var(--paper);border:4px solid var(--line);border-radius:30px;padding:12px 26px;font-size:23px;font-weight:700;line-height:1.45;margin-bottom:34px}
.bubble::before,.bubble::after{content:"";position:absolute;border-style:solid}
.bubble::before{left:54px;bottom:-30px;border-width:30px 20px 0 0;border-color:var(--line) transparent transparent transparent}
.bubble::after{left:59px;bottom:-20px;border-width:20px 13px 0 0;border-color:var(--paper) transparent transparent transparent}
.stat{font-weight:900;line-height:1;color:var(--c,var(--purple));-webkit-text-stroke:3px var(--ink);paint-order:stroke fill;text-shadow:6px 6px 0 var(--ink);white-space:nowrap}
.num{flex:none;width:54px;height:54px;border-radius:50%;border:4px solid var(--line);background:var(--c,var(--blue));color:#fff;font-weight:900;font-size:26px;display:grid;place-items:center;box-shadow:4px 4px 0 var(--line)}
.icon{flex:none;width:var(--s,110px);height:var(--s,110px);border-radius:50%;border:4px solid var(--line);background:var(--paper);box-shadow:6px 6px 0 var(--line);display:grid;place-items:center}
.icon img{width:62%;height:62%}
.speed{position:absolute;pointer-events:none;background:repeating-conic-gradient(from 0deg at 50% 50%,rgba(17,17,17,.85) 0deg .9deg,transparent .9deg 6deg);-webkit-mask:radial-gradient(circle at 50% 50%,transparent 0 32%,#000 62%);mask:radial-gradient(circle at 50% 50%,transparent 0 32%,#000 62%)}
.dots{position:absolute;pointer-events:none;background:url('../../assets/shapes/04-dot-grid.svg') repeat;background-size:150px;opacity:.55}
.stripes{position:absolute;pointer-events:none;background:url('../../assets/shapes/08-diagonal-stripes.svg') repeat;background-size:150px;opacity:.3}
.slab{position:absolute;background:var(--c)}.slab.k{background:var(--ink)}
table.tbl{border-collapse:collapse;border:4px solid var(--line);box-shadow:8px 8px 0 var(--line);background:var(--paper)}
.tbl th{background:var(--ink);color:#fff;font-weight:900;font-size:22px;text-align:left;padding:12px 20px}
.tbl td{border-top:2px solid var(--line);font-size:22px;padding:11px 20px}
.tbl td.r,.tbl th.r{text-align:right;font-variant-numeric:tabular-nums}
.tbl tr.em td{font-weight:900;color:var(--blue)}
.foot{position:absolute;left:72px;right:170px;bottom:22px;font-size:14px;color:var(--muted)}
.pno{position:absolute;right:44px;bottom:16px;border:3px solid var(--line);background:var(--paper);font-weight:900;font-size:16px;padding:1px 12px;transform:skewX(-10deg)}
.stamp{display:inline-block;border:5px solid var(--orange);color:var(--orange);font-weight:900;font-size:30px;letter-spacing:.3em;padding:4px 6px 4px 18px;transform:rotate(-8deg);background:var(--paper)}
.row{display:flex;gap:28px;align-items:center}.col{display:flex;flex-direction:column;gap:18px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:36px}
.abs{position:absolute}
.hint{position:fixed;left:50%;top:14px;transform:translateX(-50%);z-index:9;background:rgba(17,17,17,.85);color:#fff;font:14px 'Microsoft JhengHei',sans-serif;padding:6px 14px;border-radius:4px;transition:opacity .6s;pointer-events:none}
.hint.gone{opacity:0}
#deck-check{display:none}
@media print{@page{size:1280px 720px;margin:0}html,body{background:#fff;height:auto}.deck{position:static;overflow:visible}.slide{display:block!important;position:relative;left:0;top:0;margin:0;transform:none!important;break-after:page;page-break-after:always}.hint{display:none}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
"""

JS = r"""
(function(){
  var slides=[].slice.call(document.querySelectorAll('.slide')),total=slides.length,cur=0;
  var q=location.search;
  slides.forEach(function(s,i){if(i===0)return;var n=document.createElement('div');n.className='pno';n.textContent=(i+1)+' / '+total;s.appendChild(n);});
  function fit(){var k=Math.min(innerWidth/1280,innerHeight/720);slides.forEach(function(s){s.style.transform='scale('+k+')';});}
  function show(i){cur=Math.max(0,Math.min(total-1,i));slides.forEach(function(s,j){s.classList.toggle('active',j===cur);});if(location.hash!=='#'+(cur+1))history.replaceState(null,'','#'+(cur+1));}
  function fromHash(){var n=parseInt(location.hash.slice(1),10);return isNaN(n)?0:n-1;}
  document.addEventListener('keydown',function(e){var k=e.key;
    if(k==='ArrowRight'||k==='PageDown'||k===' '||k==='Enter'){e.preventDefault();show(cur+1);}
    else if(k==='ArrowLeft'||k==='PageUp'||k==='Backspace'){e.preventDefault();show(cur-1);}
    else if(k==='Home')show(0);else if(k==='End')show(total-1);
    else if(k==='f'||k==='F'){if(!document.fullscreenElement){document.documentElement.requestFullscreen&&document.documentElement.requestFullscreen();}else{document.exitFullscreen&&document.exitFullscreen();}}});
  document.getElementById('deck').addEventListener('click',function(e){show(e.clientX>innerWidth/2?cur+1:cur-1);});
  addEventListener('resize',fit);addEventListener('hashchange',function(){show(fromHash());});
  fit();show(fromHash());
  var h=document.querySelector('.hint');
  if(/check|shot/.test(q)){if(h)h.remove();}else{setTimeout(function(){if(h)h.classList.add('gone');},3500);}
  if(/check/.test(q)){
    function run(){
      var out={fontLoaded:document.fonts.check('900 20px "Noto Sans TC"'),slides:[]};
      slides.forEach(function(s,i){
        show(i);var sr=s.getBoundingClientRect(),k=sr.width/1280,issues=[];
        s.querySelectorAll('*').forEach(function(e){
          if(e.closest('.slab,.speed,.dots,.stripes'))return;
          var r=e.getBoundingClientRect();if(!r.width||!r.height)return;
          var L=(r.left-sr.left)/k,T=(r.top-sr.top)/k,R=(r.right-sr.left)/k,B=(r.bottom-sr.top)/k;
          var txt=(e.textContent||e.getAttribute('src')||'').trim().slice(0,16);
          if(R>1262||L<14||T<6||B>714)issues.push('超出投影片邊界：'+txt);
          else if(!e.closest('.foot,.pno')&&!s.classList.contains('nofoot')&&B>672)issues.push('壓到頁尾區：'+txt);
          if(e.classList.contains('stat')){var p=e.parentElement.closest('.panel')||s,pr=p.getBoundingClientRect();if(r.right>pr.right+2||r.left<pr.left-2)issues.push('大數字超出框：'+txt);}
          if(e.tagName==='IMG'&&(!e.complete||e.naturalWidth===0))issues.push('圖片無法載入：'+e.getAttribute('src'));
        });
        s.querySelectorAll('.panel,table.tbl').forEach(function(p){
          var pr=p.getBoundingClientRect();p.querySelectorAll('p,li,td,th,div,img').forEach(function(c){if(c.closest('.slab,.speed,.dots,.stripes,.tag'))return;var cr=c.getBoundingClientRect();if(cr.width&&(cr.right>pr.right+2||cr.bottom>pr.bottom+2))issues.push('內容超出框：'+(c.textContent||'').trim().slice(0,16));});
        });
        out.slides.push({no:i+1,issues:issues.filter(function(v,j,a){return a.indexOf(v)===j;})});
      });
      var pre=document.getElementById('deck-check');pre.textContent=JSON.stringify(out);document.title='DECK-CHECK-DONE';
    }
    var done=false;function go(){if(done)return;done=true;run();}
    (document.fonts&&document.fonts.ready?document.fonts.ready:Promise.resolve()).then(function(){setTimeout(go,300);});
    setTimeout(go,6000);
  }
})();
"""


def h_icon(a, size, extra=""):
    return f'<div class="icon" style="--s:{size}px;{extra}"><img src="{esc(asset(a))}" alt=""></div>'


def h_items(items, numbered=False, colors=None):
    out = []
    for i, it in enumerate(items):
        head, tail = split_item(it)
        badge = f'<div class="num" style="--c:var(--{(colors or ACCENTS)[i % 4]})">{i + 1}</div>' if numbered else ""
        body = f'<p style="font-weight:900;font-size:28px;">{esc(head)}</p>' + (f'<p class="muted">{esc(tail)}</p>' if tail else "")
        out.append(f'<div class="row" style="gap:20px;align-items:flex-start;">{badge}<div>{body}</div></div>')
    return "".join(out)


def h_table(t, rows_em=None, right_cols=()):
    rows = t["rows"]
    head = "".join(f'<th class="{"r" if j in right_cols else ""}">{esc(c)}</th>' for j, c in enumerate(rows[0]))
    body = ""
    for i, r in enumerate(rows[1:]):
        cls = ' class="em"' if rows_em and rows_em(r) else ""
        body += f"<tr{cls}>" + "".join(f'<td class="{"r" if j in right_cols else ""}">{esc(c)}</td>' for j, c in enumerate(r)) + "</tr>"
    return f'<table class="tbl"><tr>{head}</tr>{body}</table>'


def html_slide(s, ctx):
    f, lay, no = s["fields"], s["layout"], s["no"]
    accent = ACCENTS[no % 4]
    ic = icons(f)
    foot = ""
    if f.get("出處"):
        foot = f'<div class="foot">{"草稿｜" if ctx["draft"] else ""}出處：{esc(f["出處"])}</div>'
    elif lay not in ("封面", "分隔頁", "免責聲明"):
        foot = f'<div class="foot">{"草稿｜" if ctx["draft"] else ""}{esc(ctx["topic"])}</div>'
    head = f'<h1 class="t" style="--hl:var(--{accent})">{esc(s["title"])}</h1>' + (f'<p class="lead">{esc(f.get("主訊息"))}</p>' if f.get("主訊息") else "")

    if lay == "封面":
        stamp = '<div style="margin-top:34px;"><span class="stamp">草稿</span></div>' if ctx["draft"] else ""
        body = f"""<div class="slab k" style="right:-120px;top:-60px;width:640px;height:330px;transform:rotate(-14deg);"></div>
<div class="slab" style="--c:var(--purple);right:-110px;top:-80px;width:620px;height:310px;transform:rotate(-14deg);"></div>
<div class="slab k" style="left:-120px;bottom:-90px;width:560px;height:200px;transform:rotate(-14deg);"></div>
<div class="slab" style="--c:var(--orange);left:-130px;bottom:-104px;width:540px;height:190px;transform:rotate(-14deg);"></div>
<div class="speed" style="left:180px;top:-260px;width:1100px;height:1100px;opacity:.28;"></div>
<div class="dots" style="right:40px;bottom:60px;width:300px;height:210px;"></div>
<div class="abs" style="left:96px;top:190px;"><div class="tag" style="position:static;display:inline-block;font-size:22px;margin-bottom:18px;">{esc(f.get("副標"))}</div>
<div style="font-size:132px;font-weight:900;line-height:1;transform:skewX(-8deg);text-shadow:8px 8px 0 var(--purple);white-space:nowrap;">{esc(s["title"])}</div>{stamp}</div>
{h_icon(ic[0], 190, "position:absolute;right:150px;top:300px;") if ic else ""}
<div class="abs small" style="left:96px;top:530px;color:var(--ink);font-weight:700;">資料基準日：{esc(ctx["date_long"])}</div>"""
        return f'<section class="slide nofoot" aria-label="{esc(s["title"])}">{body}</section>'

    if lay == "分隔頁":
        small, _, big = s["title"].partition("：")
        big = big or s["title"]
        c = ["teal", "orange", "purple", "blue"][ctx["divider_count"] % 4]
        ctx["divider_count"] += 1
        body = f"""<div class="speed" style="left:-160px;top:-470px;width:1600px;height:1600px;opacity:.32;"></div>
<div class="slab k" style="left:-80px;top:300px;width:1500px;height:200px;transform:rotate(-8deg);"></div>
<div class="slab" style="--c:var(--{c});left:-80px;top:290px;width:1500px;height:190px;transform:rotate(-8deg);"></div>
<div class="abs" style="left:110px;top:250px;"><div class="tag" style="position:static;display:inline-block;font-size:24px;">{esc(f.get("小標") or small)}</div>
<div style="font-size:100px;font-weight:900;color:#fff;-webkit-text-stroke:4px var(--ink);paint-order:stroke fill;text-shadow:8px 8px 0 var(--ink);transform:skewX(-8deg);margin-top:14px;white-space:nowrap;">{esc(big)}</div></div>
{h_icon(ic[0], 170, "position:absolute;right:120px;top:100px;") if ic else ""}
<div class="dots" style="left:60px;bottom:40px;width:330px;height:150px;"></div>"""
        return f'<section class="slide nofoot" aria-label="{esc(s["title"])}">{body}</section>'

    if lay == "免責聲明":
        body = f"""<div class="slab k" style="right:-120px;top:-60px;width:560px;height:260px;transform:rotate(-14deg);"></div>
<div class="slab" style="--c:var(--purple);right:-110px;top:-80px;width:540px;height:240px;transform:rotate(-14deg);"></div>
<div class="dots" style="left:40px;bottom:40px;width:300px;height:160px;"></div>
<div class="abs" style="left:96px;top:160px;width:820px;"><div class="tag" style="position:static;display:inline-block;font-size:22px;">資料基準日：{esc(ctx["date_long"])}</div>
<div style="font-size:72px;font-weight:900;transform:skewX(-8deg);text-shadow:6px 6px 0 var(--orange);margin:18px 0 30px;">{esc(s["title"])}</div>
<div class="panel"><p>{esc(DISCLAIMER.format(date=ctx["date_long"]))}</p></div></div>
{h_icon(ic[0], 150, "position:absolute;right:140px;bottom:110px;") if ic else ""}"""
        return f'<section class="slide nofoot" aria-label="{esc(s["title"])}">{body}</section>'

    if lay == "目錄":
        cells = []
        for i, it in enumerate(f.get("項目") or []):
            t, d = split_item(it)
            icon = h_icon(ic[i], 88, "margin-left:auto;") if i < len(ic) else ""
            cells.append(f'<div class="panel row" style="gap:22px;min-height:170px;"><div class="num" style="--c:var(--{ACCENTS[i % 4]})">{i + 1}</div><div><p style="font-size:28px;font-weight:900;">{esc(t)}</p><p class="muted" style="font-size:22px;">{esc(d)}</p></div>{icon}</div>')
        body = f'<div class="grid2" style="row-gap:44px;margin-top:10px;">{"".join(cells)}</div>'

    elif lay == "兩欄比較":
        cols = []
        for i, side in enumerate(("左", "右")):
            icon = h_icon(ic[i], 80) if i < len(ic) else ""
            lis = "".join(f"<li>{esc(x)}</li>" for x in f.get(side + "項目") or [])
            c = ["blue", "orange"][i]
            cols.append(f'<div class="panel" style="padding-top:36px;min-height:220px;"><div class="tag" style="--c:var(--{c})">{esc(f.get(side + "標"))}</div><div class="row" style="gap:20px;align-items:flex-start;">{icon}<ul style="margin:0;padding-left:24px;display:flex;flex-direction:column;gap:8px;">{lis}</ul></div></div>')
        arrow = next((a for a in (f.get("圖貼") or []) if is_diagram(a)), None)
        mid = f'<img src="{esc(asset(arrow))}" alt="" style="width:64px;align-self:center;">' if arrow else ""
        grid = "1fr auto 1fr" if arrow else "1fr 1fr"
        extra = ""
        if f.get("補充"):
            extra = '<ul style="margin:40px 0 0;padding-left:28px;">' + "".join(f"<li><b>{esc(x)}</b></li>" for x in f["補充"]) + "</ul>"
        body = f'<div style="display:grid;grid-template-columns:{grid};gap:30px;margin-top:22px;">{cols[0]}{mid}{cols[1]}</div>{extra}'

    elif lay == "大數字":
        side = "".join(h_icon(a, 150 if len(ic) == 1 else 120) for a in ic)
        items = ""
        if f.get("項目"):
            n = len(f["項目"])
            cards = "".join(f'<div class="panel" style="padding:14px 18px;"><p style="font-weight:900;font-size:21px;">{esc(a)}</p><p class="muted" style="font-size:18px;">{esc(b)}</p></div>' for a, b in map(split_item, f["項目"]))
            items = f'<div style="display:grid;grid-template-columns:repeat({n},1fr);gap:26px;margin-top:34px;">{cards}</div>'
        body = f"""<div class="row" style="gap:48px;align-items:center;">
<div class="panel" style="flex:1;overflow:hidden;transform:rotate(-1deg);padding:22px 30px;"><div class="speed" style="left:-120px;top:-360px;width:900px;height:900px;opacity:.18;"></div>
<div style="position:relative;"><p class="small" style="color:var(--ink);font-weight:700;">{esc(f.get("前言"))}</p>
<div class="stat" style="--c:var(--{accent});font-size:64px;margin:12px 0 16px;">{esc(f.get("數字"))}</div>
<p style="font-size:25px;font-weight:900;">{esc(f.get("後語"))}</p></div></div>
<div class="row" style="gap:22px;">{side}</div></div>{items}"""

    elif lay == "表格":
        t = f["表格"]
        intro = f'<div class="bubble">{esc(f["說明"])}</div>' if f.get("說明") else ""
        note = f'<p class="small" style="margin-top:22px;">{esc(f["附註"])}</p>' if f.get("附註") else ""
        side = "".join(h_icon(a, 120) for a in ic)
        body = f'{intro}<div class="row" style="align-items:flex-start;gap:44px;"><div>{h_table(t)}{note}</div><div class="col" style="gap:30px;align-items:center;">{side}</div></div>'

    elif lay == "圖片":
        img = f'<div class="panel" style="padding:14px;width:{560 if f.get("項目") else 760}px;"><img src="{esc(f["圖"])}" alt="{esc(f.get("圖說"))}" style="width:100%;display:block;"></div>'
        if f.get("項目"):
            right = f'<div class="col" style="gap:20px;flex:1;">{h_items(f["項目"])}</div>'
            side = h_icon(ic[0], 84, "position:absolute;right:72px;top:40px;") if ic else ""
            body = f'<div class="row" style="align-items:flex-start;gap:40px;">{img}{right}</div>{side}'
        else:
            side = h_icon(ic[0], 84, "position:absolute;right:72px;top:40px;") if ic else ""
            body = f'<div class="row" style="justify-content:center;">{img}</div>{side}'

    elif lay == "例題":
        q = f'<p class="muted" style="margin-bottom:30px;">{esc(f.get("題目"))}</p>'
        ans = f"""<div class="panel" style="transform:rotate(1deg);text-align:center;padding:52px 26px 40px;">
<div class="clip"><div class="speed" style="left:-220px;top:-300px;width:900px;height:900px;opacity:.2;"></div></div><div class="tag">{esc(f.get("答案標籤") or "答案")}</div>
<div style="position:relative;"><div class="stat" style="--c:var(--{accent});font-size:84px;">{esc(f.get("答案"))}</div></div></div>"""
        expl = f'<p style="margin-top:28px;font-weight:700;">{esc(f["說明"])}</p>' if f.get("說明") else ""
        side = h_icon(ic[0], 84, "position:absolute;right:72px;top:40px;") if ic else ""
        if f.get("步驟"):
            t = h_table(f["步驟"], rows_em=lambda r: r[0].startswith("＝"), right_cols=(1,))
            body = f'{q}<div class="row" style="align-items:flex-start;gap:56px;"><div>{t}</div><div style="flex:1;">{ans}{expl}</div></div>{side}'
        else:
            formula = f'<div class="panel" style="overflow:hidden;padding:26px 34px;transform:rotate(-.5deg);"><div class="stripes" style="inset:0;"></div><p style="position:relative;font-size:42px;font-weight:700;white-space:nowrap;">{esc(f.get("算式"))}</p></div>' if f.get("算式") else ""
            body = f'{q}<div class="row" style="align-items:stretch;gap:44px;"><div style="flex:1.3;">{formula}{expl}</div><div style="flex:1;">{ans}</div></div>{side}'

    elif lay == "條列":
        side = h_icon(ic[0], 150) if ic else ""
        body = f'<div class="row" style="align-items:flex-start;gap:40px;"><div class="col" style="gap:28px;flex:1;">{h_items(f["項目"], numbered=True)}</div>{side}</div>'

    elif lay == "常見誤解":
        side = "".join(h_icon(a, 104) for a in ic)
        body = f'<div class="row" style="align-items:flex-start;gap:40px;">{h_table(f["表格"])}<div class="col" style="gap:34px;align-items:center;">{side}</div></div>'
    else:
        body = ""
    return f'<section class="slide" aria-label="{esc(s["title"])}">{head}{body}{foot}</section>'


def build_html(slides, ctx):
    ctx = dict(ctx, divider_count=0)
    secs = "\n".join(html_slide(s, ctx) for s in slides)
    return f"""<!doctype html>
<html lang="zh-Hant-TW">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(ctx["topic"])}</title>
<!--
  由 tools/deck.py render 從 03_逐頁.md 產生，請勿手改；要修改內容請改 03_逐頁.md 後重新 render。
  層級：{esc(ctx["level"])}｜狀態：{"草稿" if ctx["draft"] else "定稿"}｜資料基準日：{esc(ctx["date_long"])}
  引用頁：{esc("、".join(ctx["refs"]))}
  操作：→ / 空白鍵 / 點右半邊 下一頁；← / 點左半邊 上一頁；Home / End；F 全螢幕；列印時一頁一張。
-->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;700;900&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<main class="deck" id="deck">
{secs}
</main>
<div class="hint">← → 翻頁｜F 全螢幕</div>
<pre id="deck-check"></pre>
<script>{JS}</script>
</body>
</html>
"""


# ====================================================================== PowerPoint（繪圖指令）

def P(**kw):
    return kw


def p_panel(x, y, w, h, fill="white", shadow=True, line=True):
    return P(k="box", shape="rect", x=x, y=y, w=w, h=h, fill=HEX[fill], line=HEX["ink"] if line else None, lw=2.25, shadow=shadow)


def p_icon(a, x, y, d, root):
    return [P(k="box", shape="oval", x=x, y=y, w=d, h=d, fill=HEX["white"], line=HEX["ink"], lw=2.25, shadow=True),
            P(k="pic", path=str(root / a), x=x + d * .19, y=y + d * .19, w=d * .62, h=d * .62)]


def p_text(text, x, y, w, h, size, color="ink", bold=False, align="l", anchor="t", wrap=True, name=""):
    return P(k="text", text=text, x=x, y=y, w=w, h=h, size=size, color=HEX.get(color, color), bold=bold,
             align=align, anchor=anchor, wrap=wrap, name=name)


def p_badge(n, x, y, d, color):
    return [P(k="box", shape="oval", x=x, y=y, w=d, h=d, fill=HEX[color], line=HEX["ink"], lw=2.25, shadow=False,
              text=str(n), size=d * .42, color=HEX["white"], bold=True)]


def p_tag(text, x, y, color="ink", size=13):
    w = len(text) * size * 1.05 + 24
    return [P(k="box", shape="rect", x=x, y=y, w=w, h=size * 1.9, fill=HEX[color], line=None, shadow=False,
              text=text, size=size, color=HEX["white"], bold=True)]


def p_table(rows, x, y, colw, rowh=30, size=15, right_cols=(), em_row=None, maxh=None):
    return P(k="table", rows=rows, x=x, y=y, colw=colw, rowh=rowh, size=size, right_cols=list(right_cols),
             em_rows=[i for i, r in enumerate(rows) if em_row and i > 0 and em_row(r)], maxh=maxh)


def svg_ratio(path):
    try:
        m = re.search(r'viewBox="\s*[\d.]+\s+[\d.]+\s+([\d.]+)\s+([\d.]+)"', path.read_text(encoding="utf-8"))
        return float(m.group(2)) / float(m.group(1))
    except Exception:
        return 0.55


def pptx_slide(s, ctx):
    f, lay, no = s["fields"], s["layout"], s["no"]
    root, deck = ctx["root"], ctx["deck"]
    accent = ACCENTS[no % 4]
    ic = icons(f)
    sh = []
    content_top = 118

    def head():
        sh.append(P(k="box", shape="rect", x=40, y=58, w=min(len(s["title"]) * 30 * .98 + 28, 880), h=16, fill=HEX[accent], line=None, shadow=False))
        sh.append(p_text(s["title"], 48, 22, 864, 50, 30, bold=True, anchor="m", wrap=False, name="Title"))
        if f.get("主訊息"):
            sh.append(p_text(f["主訊息"], 48, 80, 864, 28, 16, "muted", bold=True, anchor="m", name="Lead"))

    def foot():
        src = f.get("出處")
        text = ("草稿｜" if ctx["draft"] else "") + (f"出處：{src}" if src else ctx["topic"])
        sh.append(p_text(text, 48, 510, 780, 16, 10, "muted", name="Source"))
        sh.append(p_text(f"{no} / {ctx['total']}", 852, 508, 64, 18, 11, bold=True, align="r", name="PageNo"))

    if lay == "封面":
        sh += [P(k="box", shape="rect", x=560, y=-60, w=520, h=250, fill=HEX["ink"], line=None, shadow=False, rot=-14),
               P(k="box", shape="rect", x=570, y=-75, w=505, h=235, fill=HEX["purple"], line=None, shadow=False, rot=-14),
               P(k="box", shape="rect", x=-100, y=478, w=430, h=150, fill=HEX["ink"], line=None, shadow=False, rot=-14),
               P(k="box", shape="rect", x=-110, y=490, w=420, h=140, fill=HEX["orange"], line=None, shadow=False, rot=-14)]
        sh += p_tag(f.get("副標", ""), 72, 150, size=15)
        sh.append(p_text(s["title"], 72, 196, 620, 110, 72, bold=True, anchor="m", wrap=False, name="Title"))
        if ctx["draft"]:
            sh.append(P(k="box", shape="rect", x=76, y=322, w=110, h=44, fill=HEX["white"], line=HEX["orange"], lw=3.5, shadow=False, rot=-8,
                        text="草稿", size=22, color=HEX["orange"], bold=True))
        if ic:
            sh += p_icon(ic[0], 690, 230, 150, root)
        sh.append(p_text(f"資料基準日：{ctx['date_long']}", 72, 392, 400, 20, 12, bold=True))
    elif lay == "分隔頁":
        small, _, big = s["title"].partition("：")
        c = ["teal", "orange", "purple", "blue"][ctx["divider_count"] % 4]
        ctx["divider_count"] += 1
        sh += [P(k="box", shape="rect", x=-60, y=232, w=1100, h=150, fill=HEX["ink"], line=None, shadow=False, rot=-8),
               P(k="box", shape="rect", x=-60, y=224, w=1100, h=142, fill=HEX[c], line=None, shadow=False, rot=-8)]
        sh += p_tag(f.get("小標") or small, 84, 192, size=16)
        sh.append(p_text(big or s["title"], 84, 262, 820, 84, 54, bold=True, anchor="m", wrap=False, name="Title"))
        if ic:
            sh += p_icon(ic[0], 760, 60, 130, root)
    elif lay == "免責聲明":
        sh += [P(k="box", shape="rect", x=600, y=-60, w=460, h=200, fill=HEX["ink"], line=None, shadow=False, rot=-14),
               P(k="box", shape="rect", x=610, y=-72, w=450, h=190, fill=HEX["purple"], line=None, shadow=False, rot=-14)]
        sh += p_tag(f"資料基準日：{ctx['date_long']}", 72, 120, size=15)
        sh.append(p_text(s["title"], 72, 160, 600, 60, 40, bold=True, anchor="m", name="Title"))
        sh.append(p_panel(72, 240, 620, 150))
        sh.append(p_text(DISCLAIMER.format(date=ctx["date_long"]), 92, 256, 580, 120, 16, anchor="m"))
        if ic:
            sh += p_icon(ic[0], 760, 290, 130, root)
    else:
        head()
        if lay == "目錄":
            for i, it in enumerate(f.get("項目") or []):
                t, d = split_item(it)
                x, y = 48 + (i % 2) * 444, content_top + 14 + (i // 2) * 180
                sh.append(p_panel(x, y, 420, 156))
                sh += p_badge(i + 1, x + 20, y + 22, 44, ACCENTS[i % 4])
                sh.append(p_text(t, x + 80, y + 22, 230, 32, 21, bold=True, anchor="m"))
                sh.append(p_text(d, x + 80, y + 60, 230, 70, 17, "muted"))
                if i < len(ic):
                    sh += p_icon(ic[i], x + 318, y + 40, 76, root)
        elif lay == "兩欄比較":
            arrow = next((a for a in (f.get("圖貼") or []) if is_diagram(a)), None)
            w = 396 if arrow else 420
            xs = (48, 516) if arrow else (48, 492)
            for i, side in enumerate(("左", "右")):
                x, y = xs[i], content_top + 26
                sh.append(p_panel(x, y, w, 120 + 34 * len(f.get(side + "項目") or [])))
                sh += p_tag(f.get(side + "標", ""), x + 14, y - 13, ["blue", "orange"][i], 13)
                off = 0
                if i < len(ic):
                    sh += p_icon(ic[i], x + 18, y + 32, 62, root)
                    off = 78
                sh.append(P(k="bullets", items=f.get(side + "項目") or [], x=x + 18 + off, y=y + 30, w=w - 36 - off, h=76 + 34 * len(f.get(side + "項目") or []), size=17))
            if arrow:
                sh.append(P(k="pic", path=str(root / arrow), x=452, y=content_top + 110, w=56, h=34))
            if f.get("補充"):
                sh.append(P(k="bullets", items=f["補充"], x=48, y=content_top + 60 + 120 + 34 * max(len(f.get("左項目") or []), len(f.get("右項目") or [])), w=864, h=60, size=18, bold=True))
        elif lay == "大數字":
            y = content_top + 10
            sh.append(p_panel(48, y, 560, 190))
            sh.append(p_text(f.get("前言", ""), 72, y + 16, 500, 22, 14, bold=True))
            sh.append(p_text(f.get("數字", ""), 72, y + 44, 520, 74, 46, accent, bold=True, anchor="m", wrap=False, name="Stat"))
            sh.append(p_text(f.get("後語", ""), 72, y + 124, 520, 54, 18, bold=True))
            for i, a in enumerate(ic):
                sh += p_icon(a, 660 + i * 130, y + 30, 116 if len(ic) > 1 else 140, root)
            items = f.get("項目") or []
            if items:
                n = len(items)
                cw = (864 - (n - 1) * 20) / n
                for i, it in enumerate(items):
                    a, b = split_item(it)
                    x = 48 + i * (cw + 20)
                    sh.append(p_panel(x, y + 222, cw, 112))
                    sh.append(p_text(a, x + 14, y + 234, cw - 28, 40, 14, bold=True))
                    sh.append(p_text(b, x + 14, y + 276, cw - 28, 52, 13, "muted"))
        elif lay == "表格":
            y = content_top + 4
            if f.get("說明"):
                sh.append(P(k="box", shape="round", x=48, y=y, w=min(len(f["說明"]) * 16 + 44, 864), h=38, fill=HEX["white"], line=HEX["ink"], lw=2.25, shadow=False,
                            text=f["說明"], size=15, color=HEX["ink"], bold=True))
                y += 56
            rows = f["表格"]["rows"]
            n = len(rows[0])
            colw = [480, 280] if n == 2 else [170, 300, 300] if n == 3 else [760 / n] * n
            sh.append(p_table(rows, 48, y, colw, rowh=36, size=16, maxh=490 - y))
            if f.get("附註"):
                sh.append(p_text(f["附註"], 48, y + 36 * len(rows) + 16, 780, 40, 12, "muted", name="Note"))
            for i, a in enumerate(ic):
                sh += p_icon(a, 836, y + 10 + i * 110, 84, root)
        elif lay == "圖片":
            path = deck / f["圖"]
            ratio = svg_ratio(path)
            if f.get("項目"):
                w = 430
                sh.append(p_panel(48, content_top + 6, w + 20, w * ratio + 20))
                sh.append(P(k="pic", path=str(path), x=58, y=content_top + 16, w=w, h=w * ratio))
                for i, it in enumerate(f["項目"]):
                    a, b = split_item(it)
                    yy = content_top + 10 + i * 84
                    sh.append(p_text(a, 530, yy, 382, 26, 17, bold=True))
                    sh.append(p_text(b, 530, yy + 28, 382, 46, 14, "muted"))
            else:
                w = 560
                sh.append(p_panel(200, content_top + 4, w + 20, w * ratio + 20))
                sh.append(P(k="pic", path=str(path), x=210, y=content_top + 14, w=w, h=w * ratio))
            if ic:
                sh += p_icon(ic[0], 852, 20, 60, root)
        elif lay == "例題":
            sh.append(p_text(f.get("題目", ""), 48, content_top, 790, 40, 16, "muted"))
            y = content_top + 52
            if f.get("步驟"):
                rows = f["步驟"]["rows"]
                sh.append(p_table(rows, 48, y, [260, 160], rowh=34, size=16, right_cols=(1,), em_row=lambda r: r[0].startswith("＝"), maxh=490 - y))
                ax, aw = 500, 412
            else:
                sh.append(p_panel(48, y, 864, 86))
                sh.append(p_text(f.get("算式", ""), 72, y + 8, 816, 70, 28, bold=True, anchor="m", wrap=False, name="Formula"))
                y += 112
                ax, aw = 48, 400
            sh.append(p_panel(ax, y, aw, 150))
            sh += p_tag(f.get("答案標籤") or "答案", ax + 14, y - 13, size=13)
            sh.append(p_text(f.get("答案", ""), ax + 16, y + 30, aw - 32, 90, 48, accent, bold=True, align="c", anchor="m", wrap=False, name="Stat"))
            if f.get("說明"):
                if f.get("步驟"):
                    sh.append(p_text(f["說明"], ax, y + 170, aw, 60, 15, bold=True))
                else:
                    sh.append(p_text(f["說明"], 480, y + 10, 432, 130, 17, bold=True, anchor="m"))
            if ic:
                sh += p_icon(ic[0], 852, 20, 60, root)
        elif lay == "條列":
            items = f.get("項目") or []
            step = min(84, 360 / max(len(items), 1))
            for i, it in enumerate(items):
                a, b = split_item(it)
                y = content_top + 6 + i * step
                sh += p_badge(i + 1, 48, y + 4, 40, ACCENTS[i % 4])
                sh.append(p_text(a, 104, y, 640, 30, 19, bold=True, anchor="m"))
                sh.append(p_text(b, 104, y + 32, 640, 28, 16, "muted"))
            if ic:
                sh += p_icon(ic[0], 780, content_top + 20, 120, root)
        elif lay == "常見誤解":
            rows = f["表格"]["rows"]
            sh.append(p_table(rows, 48, content_top + 6, [300, 460], rowh=44, size=17, maxh=480 - content_top))
            for i, a in enumerate(ic):
                sh += p_icon(a, 836, content_top + 20 + i * 130, 84, root)
        foot()
    notes = "；".join(x for x in [f.get("備註", ""), ("引用頁：" + f["引用頁"]) if f.get("引用頁") else ""] if x)
    return {"no": no, "shapes": sh, "notes": notes}


def build_pptx_spec(slides, ctx):
    ctx = dict(ctx, divider_count=0, total=len(slides))
    return {"font": "Microsoft JhengHei", "slides": [pptx_slide(s, ctx) for s in slides]}
