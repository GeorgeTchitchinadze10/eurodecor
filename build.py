# -*- coding: utf-8 -*-
"""
Eurodecor static-site generator.
Reads 'Eurodecor - Catalog.xlsx' + the category photo folders, optimizes images,
and writes a static site into ./docs (ready for GitHub Pages).

Run:  py build.py
"""
import os, re, shutil, html
import openpyxl
from PIL import Image, ImageOps

# ---------------------------------------------------------------- paths
ROOT     = r"D:\Eurodecor"
XLSX     = os.path.join(ROOT, "Eurodecor - Catalog.xlsx")
LOGO_SRC = os.path.join(ROOT, "Eurodecor logo.jpg")
HERE     = os.path.dirname(os.path.abspath(__file__))
OUT      = os.path.join(HERE, "docs")
IMG_OUT  = os.path.join(OUT, "assets", "img")

# ---------------------------------------------------------------- business config
# EDIT THESE with the real details, then re-run.  (Placeholders for now.)
BIZ = {
    "name_ka": "ევროდეკორი",
    "name_en": "Eurodecor",
    "sub_ka":  "შპალერების მაღაზია",
    "sub_en":  "Wallpaper Store",
    "phone_display": "+995 599 92 27 49",
    "phone_tel":     "+995599922749",
    "whatsapp":      "995599922749",
    "messenger":     "https://m.me/eurodecorwallpaper",
    "facebook":      "https://www.facebook.com/eurodecorwallpaper",
    "email":         "eurodecor.wallpaper@gmail.com",
    "address_ka": "აკაკი წერეთლის გამზირი 130, თბილისი 0119",
    "address_en": "130 Akaki Tsereteli Ave, Tbilisi 0119",
    "maps": "https://maps.app.goo.gl/YRNTNxxe1ZpKK7hP9",
    "hours_ka": "ყოველდღე 10:00 – 19:00",
    "hours_en": "Every day 10:00 – 19:00",
    "site_url": "https://georgetchitchinadze10.github.io/eurodecor",
}

C_PRIMARY = "#461c5d"
C_SECOND  = "#d1c2b9"

# ---------------------------------------------------------------- type + EN descriptions
TYPE_MAP = {
    "Wallpaper":            {"ka": "შპალერი",           "en": "Wallpaper",          "key": "wallpaper"},
    "შესაღები ფლიზერინი":   {"ka": "შესაღები ფლიზელინი", "en": "Paintable flizeline","key": "paintable"},
    "Glue":                {"ka": "შპალერის წებო",      "en": "Wallpaper glue",     "key": "glue"},
}
EN_DESC = {
    "wallpaper": [
        "Vinyl on non-woven (flizeline) base",
        "Washable & moisture-resistant",
        "Easy to hang",
        "Wide choice of modern & classic designs",
    ],
    "paintable": [
        "Paintable surface",
        "Easy to hang",
    ],
    "glue": [
        "Coverage: 70–75 m²",
        "Universal wallpaper glue",
        "Easy to mix & use",
        "Strong adhesion",
        "Ideal for all wallpaper types",
    ],
}

IMG_EXT = (".png", ".jpg", ".jpeg", ".webp")

# ---------------------------------------------------------------- helpers
def esc(s):
    return html.escape(str(s), quote=True)

def clean_size(s):
    return str(s).replace("m^2", "m²").replace("^2", "²").strip() if s else ""

def bullets_ka(desc):
    if not desc:
        return []
    out = []
    for line in str(desc).split("\n"):
        line = line.strip().lstrip("✅").strip()
        if line:
            out.append(line)
    return out

def price_html(old, new, lang):
    """Return inner HTML for a price block for a given language."""
    def fmt(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return f"{int(v)} ₾"
        # string price (e.g. wholesale)
        s = str(v).replace("₾", "").strip()
        m = re.match(r"(\d+)", s)
        num = f"{m.group(1)} ₾" if m else s
        tag_ka = "საბითუმო" if "საბითუმო" in str(v) else ""
        tag_en = "wholesale" if "საბითუმო" in str(v) else ""
        tag = tag_ka if lang == "ka" else tag_en
        return f'{num} <span class="ptag">{tag}</span>' if tag else num
    new_s = fmt(new)
    old_s = fmt(old)
    parts = []
    if old_s:
        parts.append(f'<span class="old">{old_s}</span>')
    parts.append(f'<span class="new">{new_s}</span>')
    return "".join(parts)

def optimize(src, dst, maxw=1100, quality=80):
    if os.path.exists(dst):
        return
    im = Image.open(src)
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "P", "LA"):
        im = im.convert("RGB")
    w, h = im.size
    if w > maxw:
        im = im.resize((maxw, round(h * maxw / w)), Image.LANCZOS)
    im.save(dst, "WEBP", quality=quality, method=6)

def optimize_logo(src, dst, maxw=420):
    if os.path.exists(dst):
        return
    im = Image.open(src)
    im = ImageOps.exif_transpose(im)
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    w, h = im.size
    if w > maxw:
        im = im.resize((maxw, round(h * maxw / w)), Image.LANCZOS)
    im.save(dst, "WEBP", quality=90, method=6)

# ---------------------------------------------------------------- read catalog
def read_categories():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["Categories"]
    cats = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        no, typ, desc, size, old, new = row[:6]
        if no is None or str(no).strip() == "" or typ is None:
            continue
        no = str(no).strip().zfill(2)
        if not no.isdigit():
            continue
        tinfo = TYPE_MAP.get(str(typ).strip(), {"ka": str(typ), "en": str(typ), "key": "wallpaper"})
        cats.append({
            "no": no,
            "type_ka": tinfo["ka"], "type_en": tinfo["en"], "key": tinfo["key"],
            "desc_ka": bullets_ka(desc),
            "desc_en": EN_DESC.get(tinfo["key"], []),
            "size": clean_size(size),
            "old": old, "new": new,
        })
    return cats

# ---------------------------------------------------------------- scan + build item images
def build_items(cat):
    """Scan the source folder for a category, optimize images, return item list."""
    folder = os.path.join(ROOT, f"category {int(cat['no'])}")
    if not os.path.isdir(folder):
        return []
    files = [f for f in os.listdir(folder) if f.lower().endswith(IMG_EXT)]

    out_dir = os.path.join(IMG_OUT, f"c{cat['no']}")
    os.makedirs(out_dir, exist_ok=True)

    # Numbered files (category 1 style): "#0101.png", "#0101 (2).png"
    numbered = {}
    plain = []
    for f in files:
        m = re.match(r"#?\s*(\d{3,4})\s*(\(2\))?", f)
        if m and f.lstrip().startswith("#"):
            num = m.group(1)
            numbered.setdefault(num, {"main": None, "detail": None})
            if m.group(2):
                numbered[num]["detail"] = f
            else:
                numbered[num]["main"] = f
        else:
            plain.append(f)

    items = []
    if numbered:
        for num in sorted(numbered):
            src_name = numbered[num]["main"] or numbered[num]["detail"]
            src = os.path.join(folder, src_name)
            dst_name = f"{num}.webp"
            optimize(src, os.path.join(out_dir, dst_name))
            items.append({"num": num, "img": f"assets/img/c{cat['no']}/{dst_name}"})
    else:
        for i, f in enumerate(sorted(plain), start=1):
            num = f"{cat['no']}{i:02d}"
            src = os.path.join(folder, f)
            dst_name = f"{num}.webp"
            optimize(src, os.path.join(out_dir, dst_name))
            items.append({"num": num, "img": f"assets/img/c{cat['no']}/{dst_name}"})
    return items

# ---------------------------------------------------------------- HTML fragments
def i18n(ka, en, tag="span"):
    return f'<{tag} class="ka">{ka}</{tag}><{tag} class="en">{en}</{tag}>'

ICON_PATHS = {
    "phone": '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6 19.8 19.8 0 0 1-3.1-8.7A2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1.9.3 1.8.6 2.6a2 2 0 0 1-.5 2.1L8 9.9a16 16 0 0 0 6 6l1.5-1.2a2 2 0 0 1 2.1-.5c.8.3 1.7.5 2.6.6a2 2 0 0 1 1.7 2z"/>',
    "chat": '<path d="M21 11.5a8.4 8.4 0 0 1-9 8.4 8.4 8.4 0 0 1-3.9-.9L3 20l1.3-3.9A8.4 8.4 0 1 1 21 11.5z"/>',
    "send": '<path d="M22 2 11 13"/><path d="M22 2 15 22l-4-9-9-4z"/>',
    "pin": '<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>',
    "mail": '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m2 7 10 6 10-6"/>',
    "chevL": '<polyline points="15 18 9 12 15 6"/>',
    "chevR": '<polyline points="9 18 15 12 9 6"/>',
    "close": '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
    "zoom": '<circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.5" y2="16.5"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/>',
}

def icon(name, size=20):
    return (f'<svg class="ic" width="{size}" height="{size}" viewBox="0 0 24 24" '
            f'fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" '
            f'stroke-linejoin="round" aria-hidden="true">{ICON_PATHS[name]}</svg>')

def head(title_ka, title_en, desc_ka, canonical):
    ld = f'''{{
      "@context":"https://schema.org","@type":"HomeGoodsStore",
      "name":"Eurodecor {esc(BIZ["sub_en"])}",
      "image":"{esc(BIZ["site_url"])}/assets/img/logo.webp",
      "address":{{"@type":"PostalAddress","streetAddress":"{esc(BIZ["address_en"])}","addressLocality":"Tbilisi","addressCountry":"GE"}},
      "telephone":"{esc(BIZ["phone_tel"])}","priceRange":"₾","url":"{esc(BIZ["site_url"])}"
    }}'''
    return f'''<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title_ka)} | {esc(title_en)}</title>
<meta name="description" content="{esc(desc_ka)}">
<link rel="canonical" href="{esc(canonical)}">
<meta property="og:title" content="{esc(title_ka)} | {esc(title_en)}">
<meta property="og:description" content="{esc(desc_ka)}">
<meta property="og:type" content="website">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=Noto+Serif+Georgian:wght@500;600;700&family=Noto+Sans+Georgian:wght@400;500;600&display=swap" rel="stylesheet">
<link rel="stylesheet" href="styles.css">
<link rel="icon" href="assets/img/logo.webp">
<script type="application/ld+json">{ld}</script>'''

def header_html():
    return f'''<header class="site-header">
  <a class="brand" href="index.html">
    <img src="assets/img/logo.webp" alt="Eurodecor" class="brand-logo">
    <span class="brand-text">
      <strong>Eurodecor</strong>
      {i18n("შპალერების მაღაზია","Wallpaper Store", "small")}
    </span>
  </a>
  <button id="langBtn" class="lang-btn" type="button" aria-label="Language">
    <span class="ka">EN</span><span class="en">ქარ</span>
  </button>
</header>'''

def contact_bar():
    b = BIZ
    return f'''<nav class="contact-bar" aria-label="Contact">
  <a href="tel:{esc(b['phone_tel'])}" class="cbtn cbtn-call">{icon('phone')}<span>{i18n("დარეკვა","Call")}</span></a>
  <a href="https://wa.me/{esc(b['whatsapp'])}" class="cbtn cbtn-wa" target="_blank" rel="noopener">{icon('chat')}<span>WhatsApp</span></a>
  <a href="{esc(b['messenger'])}" class="cbtn cbtn-msg" target="_blank" rel="noopener">{icon('send')}<span>Messenger</span></a>
  <a href="{esc(b['maps'])}" class="cbtn cbtn-map" target="_blank" rel="noopener">{icon('pin')}<span>{i18n("მისამართი","Directions")}</span></a>
  <a href="mailto:{esc(b['email'])}" class="cbtn cbtn-mail">{icon('mail')}<span>Email</span></a>
</nav>'''

def footer_html():
    b = BIZ
    return f'''<footer class="site-footer">
  <div class="foot-grid">
    <div>
      <img src="assets/img/logo.webp" alt="Eurodecor" class="foot-logo">
      <p class="foot-tag">{i18n("პირდაპირ ქარხნიდან · 25 წლიანი გამოცდილება","Factory-direct · 25 years of experience","span")}</p>
    </div>
    <div>
      <h4>{i18n("კონტაქტი","Contact","span")}</h4>
      <p>{i18n(b['address_ka'], b['address_en'],"span")}</p>
      <p><a href="tel:{esc(b['phone_tel'])}">{esc(b['phone_display'])}</a></p>
      <p><a href="mailto:{esc(b['email'])}">{esc(b['email'])}</a></p>
      <p>{i18n(b['hours_ka'], b['hours_en'],"span")}</p>
    </div>
    <div>
      <h4>{i18n("გვეწვიეთ","Visit us","span")}</h4>
      <p><a href="{esc(b['maps'])}" target="_blank" rel="noopener">{i18n("რუკაზე ნახვა →","Open in Maps →","span")}</a></p>
      <p><a href="{esc(b['facebook'])}" target="_blank" rel="noopener">Facebook →</a></p>
    </div>
  </div>
  <p class="copyright">© 2026 Eurodecor · {i18n("ყველა უფლება დაცულია","All rights reserved","span")}</p>
</footer>'''

LANG_JS = '''<script>
(function(){
  var root=document.documentElement;
  var saved=localStorage.getItem('lang')||'ka';
  root.setAttribute('data-lang',saved);
  var btn=document.getElementById('langBtn');
  if(btn){btn.addEventListener('click',function(){
    var cur=root.getAttribute('data-lang')==='en'?'ka':'en';
    root.setAttribute('data-lang',cur);localStorage.setItem('lang',cur);
  });}
})();
</script>'''

LIGHTBOX_JS = '''<script>
(function(){
  var figs=[].slice.call(document.querySelectorAll('.item-grid .item'));
  var ov=document.getElementById('lb');
  if(!figs.length||!ov)return;
  var big=document.getElementById('lb-img'),cap=document.getElementById('lb-cap'),
      cnt=document.getElementById('lb-count'),idx=0;
  var data=figs.map(function(f){return {src:f.querySelector('img').getAttribute('src'),num:f.getAttribute('data-num')};});
  function show(i){idx=(i+data.length)%data.length;big.src=data[idx].src;cap.textContent='#'+data[idx].num;cnt.textContent=(idx+1)+' / '+data.length;}
  function open(i){show(i);ov.classList.add('open');ov.setAttribute('aria-hidden','false');document.body.style.overflow='hidden';}
  function close(){ov.classList.remove('open');ov.setAttribute('aria-hidden','true');document.body.style.overflow='';}
  figs.forEach(function(f,i){
    f.addEventListener('click',function(){open(i);});
    f.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();open(i);}});
  });
  document.getElementById('lb-next').addEventListener('click',function(e){e.stopPropagation();show(idx+1);});
  document.getElementById('lb-prev').addEventListener('click',function(e){e.stopPropagation();show(idx-1);});
  document.getElementById('lb-close').addEventListener('click',close);
  ov.addEventListener('click',function(e){if(e.target===ov||e.target.classList.contains('lb-stage'))close();});
  document.addEventListener('keydown',function(e){
    if(!ov.classList.contains('open'))return;
    if(e.key==='Escape')close();else if(e.key==='ArrowRight')show(idx+1);else if(e.key==='ArrowLeft')show(idx-1);
  });
})();
</script>'''

# ---------------------------------------------------------------- page: home
def render_home(cats):
    cards = []
    for c in cats:
        first_img = c["items"][0]["img"] if c["items"] else "assets/img/logo.webp"
        badge = f'<span class="badge">{i18n("ფასდაკლება","Sale")}</span>' if c["old"] else ""
        cards.append(f'''<a class="cat-card" href="category-{c['no']}.html">
      <div class="cat-thumb">{badge}
        <img loading="lazy" src="{first_img}" alt="{esc(c['type_en'])} {c['no']}">
      </div>
      <div class="cat-body">
        <span class="cat-type">{i18n(c['type_ka'], c['type_en'])}</span>
        <div class="price-wrap"><span class="ka">{price_html(c["old"], c["new"], "ka")}</span><span class="en">{price_html(c["old"], c["new"], "en")}</span></div>
        <span class="cat-size">{esc(c['size'])} · #{c['no']}</span>
      </div>
    </a>''')
    grid = "\n".join(cards)
    hero = f'''<section class="hero">
    <div class="hero-inner">
      <span class="hero-over">{i18n("· 25 წლიანი გამოცდილება ·","· EST. 25 YEARS ·","span")}</span>
      <h1>{i18n("ევროდეკორი","Eurodecor","span")}</h1>
      <p class="hero-sub">{i18n("შპალერების მაღაზია","Wallpaper Store","span")}</p>
      <span class="hero-rule"></span>
      <p class="hero-usp">{i18n("პირდაპირ ქარხნიდან · საუკეთესო ფასები თბილისში · დიდი არჩევანი","Factory-direct · best prices in Tbilisi · huge selection","span")}</p>
    </div>
  </section>'''
    body = f'''{header_html()}
{hero}
{contact_bar()}
<main class="container">
  <h2 class="section-title">{i18n("კატეგორიები","Categories","span")}</h2>
  <div class="cat-grid">
    {grid}
  </div>
</main>
{footer_html()}
{LANG_JS}'''
    title_ka = "ევროდეკორი — შპალერების მაღაზია თბილისში"
    title_en = "Eurodecor — Wallpaper Store in Tbilisi"
    desc = "შპალერი, ვინილის შპალერი, ფლიზელინი, შესაღები შპალერი და შპალერის წებო — საუკეთესო ფასებში, პირდაპირ ქარხნიდან. თბილისი, აკაკი წერეთლის 130."
    return f'''<!doctype html><html lang="ka" data-lang="ka"><head>
{head(title_ka, title_en, desc, BIZ["site_url"] + "/")}
</head><body>
{body}
</body></html>'''

# ---------------------------------------------------------------- page: category
def render_category(c):
    items = []
    for it in c["items"]:
        items.append(f'''<figure class="item" data-num="{it['num']}" tabindex="0" role="button" aria-label="#{it['num']}">
      <div class="item-img">
        <img loading="lazy" src="{it['img']}" alt="{esc(c['type_en'])} #{it['num']}">
        <span class="zoom">{icon('zoom')}</span>
      </div>
      <figcaption>#{it['num']}</figcaption>
    </figure>''')
    grid = "\n".join(items) if items else f'<p class="muted">{i18n("ფოტოები მალე","Photos coming soon")}</p>'

    d_ka = "".join(f"<li>{esc(x)}</li>" for x in c["desc_ka"])
    d_en = "".join(f"<li>{esc(x)}</li>" for x in c["desc_en"])
    badge = f'<span class="badge badge-dark">{i18n("ფასდაკლება","Sale")}</span>' if c["old"] else ""

    intro = f'''<section class="cat-head">
    <nav class="crumb"><a href="index.html">{i18n("მთავარი","Home")}</a> <span>/</span> #{c['no']}</nav>
    <span class="cat-over">{i18n(c['type_ka'], c['type_en'])} · #{c['no']}</span>
    <h1 class="cat-title">{i18n(c['type_ka'], c['type_en'],"span")}</h1>
    <div class="cat-price">{badge}
      <span class="ka">{price_html(c['old'], c['new'], 'ka')}</span><span class="en">{price_html(c['old'], c['new'], 'en')}</span>
    </div>
    <p class="cat-size-big">{esc(c['size'])}</p>
    <ul class="feat ka">{d_ka}</ul>
    <ul class="feat en">{d_en}</ul>
    <p class="order-hint">{i18n("მოგწონთ რომელიმე? მოგვწერეთ ნომერი (მაგ. #"+c['no']+"01) Messenger-ში ან WhatsApp-ში.","Like one? Message us the number (e.g. #"+c['no']+"01) on Messenger or WhatsApp.")}</p>
  </section>'''

    lightbox = f'''<div class="lb" id="lb" aria-hidden="true">
  <button class="lb-close" id="lb-close" aria-label="Close">{icon('close', 26)}</button>
  <button class="lb-nav lb-prev" id="lb-prev" aria-label="Previous">{icon('chevL', 30)}</button>
  <figure class="lb-stage"><img id="lb-img" src="" alt=""></figure>
  <button class="lb-nav lb-next" id="lb-next" aria-label="Next">{icon('chevR', 30)}</button>
  <div class="lb-bar"><span id="lb-cap"></span><span id="lb-count"></span></div>
</div>'''

    body = f'''{header_html()}
{contact_bar()}
<main class="container">
  {intro}
  <div class="item-grid">
    {grid}
  </div>
  <p class="back"><a href="index.html">← {i18n("ყველა კატეგორია","All categories")}</a></p>
</main>
{footer_html()}
{lightbox}
{LANG_JS}
{LIGHTBOX_JS}'''
    title_ka = f"{c['type_ka']} #{c['no']} — ევროდეკორი"
    title_en = f"{c['type_en']} #{c['no']} — Eurodecor"
    desc = f"{c['type_ka']} #{c['no']}, {c['size']} — ევროდეკორი, თბილისი. საუკეთესო ფასი."
    canonical = f"{BIZ['site_url']}/category-{c['no']}.html"
    return f'''<!doctype html><html lang="ka" data-lang="ka"><head>
{head(title_ka, title_en, desc, canonical)}
</head><body>
{body}
</body></html>'''

# ---------------------------------------------------------------- CSS
def write_css():
    css = f''':root{{
  --plum:{C_PRIMARY};
  --plum-d:#2c1140;
  --plum-2:#5b2a72;
  --taupe:{C_SECOND};
  --taupe-l:#e6ddd6;
  --ivory:#faf7f4;
  --ink:#2a2230;
  --muted:#8a7f92;
  --gold:#b08d57;
  --sale:#8d2846;
  --line:rgba(70,28,93,.12);
  --serif:'Playfair Display','Noto Serif Georgian',Georgia,serif;
  --sans:'Noto Sans Georgian','Segoe UI',system-ui,-apple-system,sans-serif;
}}
*{{box-sizing:border-box}}
html{{-webkit-text-size-adjust:100%}}
body{{margin:0;font-family:var(--sans);color:var(--ink);background:var(--ivory);line-height:1.55}}
img{{max-width:100%;display:block}}
a{{color:inherit;text-decoration:none}}
.ic{{flex:none}}
.container{{max-width:1160px;margin:0 auto;padding:0 22px}}

/* language toggle */
html[data-lang="ka"] .en{{display:none !important}}
html[data-lang="en"] .ka{{display:none !important}}

/* header */
.site-header{{position:sticky;top:0;z-index:40;display:flex;align-items:center;justify-content:space-between;
  padding:12px 22px;background:rgba(250,247,244,.9);backdrop-filter:blur(10px);
  border-bottom:1px solid var(--line);color:var(--plum)}}
.brand{{display:flex;align-items:center;gap:11px}}
.brand-logo{{height:42px;width:auto;background:#fff;border-radius:9px;padding:4px;box-shadow:0 2px 8px rgba(70,28,93,.08)}}
.brand-text{{display:flex;flex-direction:column;line-height:1.05}}
.brand-text strong{{font-family:var(--serif);font-size:1.28rem;font-weight:700;letter-spacing:.3px;color:var(--plum)}}
.brand-text small{{font-size:.7rem;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-top:2px}}
.lang-btn{{background:transparent;color:var(--plum);border:1px solid var(--plum);
  padding:7px 16px;border-radius:999px;font-weight:600;cursor:pointer;font-size:.82rem;letter-spacing:.05em;transition:.16s}}
.lang-btn:hover{{background:var(--plum);color:#fff}}

/* hero */
.hero{{background:radial-gradient(130% 150% at 50% -30%,var(--plum-2) 0%,var(--plum) 46%,var(--plum-d) 100%);
  color:#fff;text-align:center;padding:78px 20px 68px}}
.hero-over{{display:block;letter-spacing:.34em;text-transform:uppercase;font-size:.72rem;color:var(--taupe);font-weight:600}}
.hero h1{{font-family:var(--serif);font-weight:600;font-size:clamp(2.7rem,6vw,4.1rem);margin:.35rem 0 0;letter-spacing:.5px}}
.hero-sub{{font-family:var(--serif);font-style:italic;color:var(--taupe);font-size:clamp(1.1rem,2.4vw,1.4rem);margin:.35rem 0 0}}
.hero-rule{{display:block;width:72px;height:2px;margin:22px auto;
  background:linear-gradient(90deg,transparent,var(--gold),transparent)}}
.hero-usp{{color:rgba(255,255,255,.82);font-size:.96rem;letter-spacing:.02em;margin:0}}

/* contact bar */
.contact-bar{{display:flex;gap:12px;justify-content:center;flex-wrap:wrap;padding:22px 14px;background:var(--ivory)}}
.cbtn{{display:inline-flex;align-items:center;gap:9px;background:#fff;color:var(--plum);
  padding:11px 20px;border-radius:999px;border:1px solid var(--line);font-weight:600;font-size:.92rem;
  box-shadow:0 2px 10px rgba(70,28,93,.05);transition:.18s}}
.cbtn:hover{{background:var(--plum);color:#fff;border-color:var(--plum);transform:translateY(-2px);box-shadow:0 12px 24px rgba(70,28,93,.18)}}
.cbtn span{{white-space:nowrap}}

/* section */
.section-title{{font-family:var(--serif);text-align:center;margin:50px 0 2px;font-size:2rem;font-weight:600;color:var(--plum)}}
.section-title::after{{content:"";display:block;width:56px;height:2px;margin:14px auto 0;
  background:linear-gradient(90deg,transparent,var(--gold),transparent)}}

/* category grid */
.cat-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:26px;margin:32px 0 64px}}
.cat-card{{background:#fff;border:1px solid var(--line);border-radius:14px;overflow:hidden;
  box-shadow:0 4px 18px rgba(70,28,93,.06);transition:transform .2s,box-shadow .2s}}
.cat-card:hover{{transform:translateY(-4px);box-shadow:0 20px 42px rgba(70,28,93,.15)}}
.cat-thumb{{position:relative;aspect-ratio:4/5;overflow:hidden;background:var(--taupe-l)}}
.cat-thumb img{{width:100%;height:100%;object-fit:cover;transition:transform .55s ease}}
.cat-card:hover .cat-thumb img{{transform:scale(1.06)}}
.badge{{position:absolute;top:12px;left:12px;z-index:2;background:var(--gold);color:#fff;
  font-size:.66rem;letter-spacing:.12em;text-transform:uppercase;font-weight:700;padding:5px 11px;border-radius:999px}}
.cat-body{{padding:16px 18px 20px}}
.cat-type{{display:block;font-weight:600;color:var(--ink);font-size:1.04rem}}
.price-wrap{{margin:9px 0 5px}}
.old{{text-decoration:line-through;color:var(--muted);margin-right:9px;font-size:.95rem}}
.new{{font-family:var(--serif);color:var(--plum);font-weight:700;font-size:1.45rem}}
.ptag{{font-size:.62rem;letter-spacing:.1em;text-transform:uppercase;background:var(--gold);color:#fff;
  padding:3px 8px;border-radius:999px;vertical-align:middle;margin-left:7px}}
.cat-size{{display:block;color:var(--muted);font-size:.82rem;margin-top:6px;letter-spacing:.03em}}

/* category page */
.cat-head{{max-width:760px;margin:36px auto 8px;text-align:center}}
.crumb{{color:var(--muted);font-size:.84rem}}
.crumb a{{color:var(--plum)}}
.crumb span{{margin:0 5px;opacity:.6}}
.cat-over{{display:block;letter-spacing:.28em;text-transform:uppercase;font-size:.72rem;color:var(--gold);font-weight:700;margin-top:16px}}
.cat-title{{font-family:var(--serif);color:var(--plum);font-size:clamp(1.9rem,4vw,2.7rem);margin:.25rem 0 .3rem;font-weight:600}}
.cat-price{{font-family:var(--serif);font-size:1.55rem;margin:.3rem 0;display:flex;gap:12px;align-items:center;justify-content:center;flex-wrap:wrap}}
.badge-dark{{font-family:var(--sans);background:var(--sale);color:#fff;font-size:.64rem;letter-spacing:.1em;
  text-transform:uppercase;padding:5px 11px;border-radius:999px;font-weight:700}}
.cat-size-big{{color:var(--muted);margin:.1rem 0 1.1rem;letter-spacing:.04em}}
.feat{{list-style:none;padding:0;margin:0 auto 1.2rem;display:inline-block;text-align:left}}
.feat li{{position:relative;padding-left:26px;margin:.44rem 0;color:#40354a}}
.feat li::before{{content:"";position:absolute;left:4px;top:.15em;width:6px;height:11px;
  border:solid var(--gold);border-width:0 2px 2px 0;transform:rotate(45deg)}}
.order-hint{{max-width:560px;margin:8px auto 0;background:#fff;border:1px solid var(--taupe);
  border-radius:12px;padding:12px 18px;color:#4a3a52;font-size:.92rem}}

/* item grid + lightbox trigger */
.item-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:18px;margin:28px 0 36px}}
.item{{margin:0;cursor:pointer;outline:none}}
.item-img{{position:relative;aspect-ratio:4/5;border-radius:12px;overflow:hidden;background:var(--taupe-l);
  border:1px solid var(--line);box-shadow:0 4px 16px rgba(70,28,93,.07)}}
.item-img img{{width:100%;height:100%;object-fit:cover;transition:transform .45s}}
.item:hover .item-img img,.item:focus .item-img img{{transform:scale(1.07)}}
.zoom{{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#fff;
  background:rgba(44,17,64,.3);opacity:0;transition:.2s}}
.item:hover .zoom,.item:focus-visible .zoom{{opacity:1}}
.item figcaption{{text-align:center;padding:9px 0 0;font-weight:600;color:var(--plum);font-size:.9rem;letter-spacing:.05em}}
.back{{margin:6px 0 44px}}
.back a{{color:var(--plum);font-weight:600}}
.muted{{color:var(--muted);text-align:center}}

/* footer */
.site-footer{{background:var(--plum-d);color:#e7dcee;margin-top:24px;padding:48px 22px 26px}}
.foot-grid{{max-width:1160px;margin:0 auto;display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:30px}}
.foot-logo{{height:48px;background:#fff;border-radius:9px;padding:5px;width:auto}}
.foot-tag{{color:var(--taupe);margin-top:12px;font-size:.9rem;max-width:230px}}
.site-footer h4{{font-family:var(--serif);color:#fff;margin:0 0 10px;font-weight:600;font-size:1.08rem}}
.site-footer a{{color:var(--taupe);transition:.15s}}
.site-footer a:hover{{color:#fff}}
.site-footer p{{margin:.35rem 0;font-size:.92rem}}
.copyright{{text-align:center;color:#a58cb8;margin-top:32px;font-size:.8rem;letter-spacing:.04em}}

/* lightbox */
.lb{{position:fixed;inset:0;z-index:100;display:none;align-items:center;justify-content:center;
  background:rgba(22,9,31,.95);padding:20px}}
.lb.open{{display:flex}}
.lb-stage{{margin:0;display:flex;align-items:center;justify-content:center}}
.lb-stage img{{max-width:92vw;max-height:82vh;object-fit:contain;border-radius:6px;box-shadow:0 24px 70px rgba(0,0,0,.65)}}
.lb-close{{position:absolute;top:16px;right:18px;width:46px;height:46px}}
.lb-nav{{position:absolute;top:50%;transform:translateY(-50%);width:54px;height:54px}}
.lb-prev{{left:16px}}.lb-next{{right:16px}}
.lb-close,.lb-nav{{display:flex;align-items:center;justify-content:center;border:none;cursor:pointer;
  color:#fff;background:rgba(255,255,255,.1);border-radius:50%;transition:.15s}}
.lb-close:hover,.lb-nav:hover{{background:rgba(255,255,255,.26)}}
.lb-bar{{position:absolute;bottom:18px;left:0;right:0;display:flex;justify-content:center;gap:18px;
  color:#fff;font-size:.95rem;letter-spacing:.05em}}
#lb-cap{{font-weight:700;color:var(--taupe)}}

@media(max-width:560px){{
  .hero{{padding:58px 18px 50px}}
  .cbtn span{{display:none}}
  .cbtn{{padding:12px}}
  .cat-grid{{gap:16px;grid-template-columns:repeat(auto-fill,minmax(150px,1fr))}}
  .item-grid{{grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:12px}}
  .lb-nav{{width:44px;height:44px}}
  .lb-prev{{left:8px}}.lb-next{{right:8px}}
}}
'''
    with open(os.path.join(OUT, "styles.css"), "w", encoding="utf-8") as f:
        f.write(css)

# ---------------------------------------------------------------- main
def main():
    os.makedirs(IMG_OUT, exist_ok=True)
    # .nojekyll so GitHub Pages serves everything as-is
    open(os.path.join(OUT, ".nojekyll"), "w").close()

    optimize_logo(LOGO_SRC, os.path.join(IMG_OUT, "logo.webp"))

    cats = read_categories()
    for c in cats:
        c["items"] = build_items(c)
        print(f"  category {c['no']} ({c['type_en']}): {len(c['items'])} items")

    write_css()
    with open(os.path.join(OUT, "index.html"), "w", encoding="utf-8") as f:
        f.write(render_home(cats))
    for c in cats:
        with open(os.path.join(OUT, f"category-{c['no']}.html"), "w", encoding="utf-8") as f:
            f.write(render_category(c))

    print(f"\nBuilt {len(cats)} categories -> {OUT}")

if __name__ == "__main__":
    main()
