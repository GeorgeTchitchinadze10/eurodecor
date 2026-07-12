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
LOGO_SRC = os.path.join(ROOT, "Eurodecor Logo.png")
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
  <a href="tel:{esc(b['phone_tel'])}" class="cbtn cbtn-call">📞<span>{i18n("დარეკვა","Call")}</span></a>
  <a href="https://wa.me/{esc(b['whatsapp'])}" class="cbtn cbtn-wa" target="_blank" rel="noopener">💬<span>WhatsApp</span></a>
  <a href="{esc(b['messenger'])}" class="cbtn cbtn-msg" target="_blank" rel="noopener">✉️<span>Messenger</span></a>
  <a href="{esc(b['maps'])}" class="cbtn cbtn-map" target="_blank" rel="noopener">📍<span>{i18n("მისამართი","Directions")}</span></a>
</nav>'''

def footer_html():
    b = BIZ
    return f'''<footer class="site-footer">
  <div class="foot-grid">
    <div>
      <img src="assets/img/logo.webp" alt="Eurodecor" class="foot-logo">
      <p class="foot-tag">{i18n("პირდაპირ ქარხნიდან · 25 წლის გამოცდილება","Factory-direct · 25 years of experience","span")}</p>
    </div>
    <div>
      <h4>{i18n("კონტაქტი","Contact","span")}</h4>
      <p>{i18n(b['address_ka'], b['address_en'],"span")}</p>
      <p><a href="tel:{esc(b['phone_tel'])}">{esc(b['phone_display'])}</a></p>
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

# ---------------------------------------------------------------- page: home
def render_home(cats):
    cards = []
    for c in cats:
        first_img = c["items"][0]["img"] if c["items"] else "assets/img/logo.webp"
        price = f'<div class="price">{price_html(c["old"], c["new"], "ka")}</div>'
        price_en = f'<div class="price en-price">{price_html(c["old"], c["new"], "en")}</div>'
        cards.append(f'''<a class="cat-card" href="category-{c['no']}.html">
      <div class="cat-thumb"><img loading="lazy" src="{first_img}" alt="{esc(c['type_en'])} {c['no']}"></div>
      <div class="cat-body">
        <span class="cat-type">{i18n(c['type_ka'], c['type_en'])}</span>
        <span class="cat-no">#{c['no']}</span>
        <div class="price-wrap"><span class="ka">{price_html(c["old"], c["new"], "ka")}</span><span class="en">{price_html(c["old"], c["new"], "en")}</span></div>
        <span class="cat-size">{esc(c['size'])}</span>
      </div>
    </a>''')
    grid = "\n".join(cards)
    hero = f'''<section class="hero">
    <div class="hero-inner">
      <h1>{i18n("ევროდეკორი","Eurodecor","span")}</h1>
      <p class="hero-sub">{i18n("შპალერების მაღაზია — პირდაპირ ქარხნიდან","Wallpaper store — factory direct","span")}</p>
      <p class="hero-usp">{i18n("25 წლის გამოცდილება · საუკეთესო ფასები თბილისში · დიდი არჩევანი","25 years of experience · best prices in Tbilisi · huge selection","span")}</p>
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
        items.append(f'''<figure class="item">
      <img loading="lazy" src="{it['img']}" alt="{esc(c['type_en'])} #{it['num']}">
      <figcaption>#{it['num']}</figcaption>
    </figure>''')
    grid = "\n".join(items) if items else f'<p class="muted">{i18n("ფოტოები მალე","Photos coming soon")}</p>'

    d_ka = "".join(f"<li>{esc(x)}</li>" for x in c["desc_ka"])
    d_en = "".join(f"<li>{esc(x)}</li>" for x in c["desc_en"])

    intro = f'''<section class="cat-head">
    <div class="cat-head-info">
      <span class="crumb"><a href="index.html">{i18n("მთავარი","Home")}</a> / #{c['no']}</span>
      <h1>{i18n(c['type_ka'], c['type_en'],"span")} <span class="hno">#{c['no']}</span></h1>
      <div class="cat-price">{price_html(c['old'], c['new'], 'ka') if True else ''}
        <span class="ka">{price_html(c['old'], c['new'], 'ka')}</span><span class="en">{price_html(c['old'], c['new'], 'en')}</span>
      </div>
      <p class="cat-size-big">{esc(c['size'])}</p>
      <ul class="feat ka">{d_ka}</ul>
      <ul class="feat en">{d_en}</ul>
      <p class="order-hint">{i18n("მოგწონთ რომელიმე? მოგვწერეთ ნომერი (მაგ. #"+c['no']+"01) Messenger-ში ან WhatsApp-ში.","Like one? Message us the number (e.g. #"+c['no']+"01) on Messenger or WhatsApp.")}</p>
    </div>
  </section>'''
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
{LANG_JS}'''
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
  --primary:{C_PRIMARY};
  --primary-d:#33143f;
  --second:{C_SECOND};
  --cream:#f6f1ee;
  --ink:#241a2b;
  --muted:#7a6f80;
  --line:#e7ddd6;
  --radius:16px;
  --shadow:0 6px 24px rgba(70,28,93,.10);
}}
*{{box-sizing:border-box}}
html{{-webkit-text-size-adjust:100%}}
body{{margin:0;font-family:'Segoe UI',system-ui,-apple-system,'Noto Sans Georgian',sans-serif;color:var(--ink);background:var(--cream);line-height:1.5}}
img{{max-width:100%;display:block}}
a{{color:inherit;text-decoration:none}}
.container{{max-width:1120px;margin:0 auto;padding:0 18px}}

/* language toggle */
html[data-lang="ka"] .en{{display:none !important}}
html[data-lang="en"] .ka{{display:none !important}}

/* header */
.site-header{{position:sticky;top:0;z-index:40;display:flex;align-items:center;justify-content:space-between;
  padding:10px 18px;background:var(--primary);color:#fff;box-shadow:var(--shadow)}}
.brand{{display:flex;align-items:center;gap:10px}}
.brand-logo{{height:40px;width:auto;background:#fff;border-radius:8px;padding:3px}}
.brand-text{{display:flex;flex-direction:column;line-height:1.05}}
.brand-text strong{{font-size:1.15rem;letter-spacing:.3px}}
.brand-text small{{font-size:.72rem;opacity:.85}}
.lang-btn{{background:rgba(255,255,255,.15);color:#fff;border:1px solid rgba(255,255,255,.35);
  padding:7px 14px;border-radius:999px;font-weight:600;cursor:pointer;font-size:.85rem}}
.lang-btn:hover{{background:rgba(255,255,255,.28)}}

/* hero */
.hero{{background:linear-gradient(135deg,var(--primary),var(--primary-d));color:#fff;text-align:center;padding:52px 18px 46px}}
.hero h1{{margin:0;font-size:2.5rem;letter-spacing:.5px}}
.hero-sub{{margin:.5rem 0 .4rem;font-size:1.15rem;color:var(--second)}}
.hero-usp{{margin:0;opacity:.9;font-size:.95rem}}

/* contact bar */
.contact-bar{{display:flex;gap:8px;justify-content:center;flex-wrap:wrap;padding:14px 12px;background:var(--second)}}
.cbtn{{display:inline-flex;align-items:center;gap:7px;background:#fff;color:var(--primary);
  padding:10px 16px;border-radius:999px;font-weight:600;font-size:.9rem;box-shadow:var(--shadow)}}
.cbtn:hover{{transform:translateY(-1px)}}
.cbtn span{{white-space:nowrap}}

/* section */
.section-title{{text-align:center;margin:34px 0 6px;font-size:1.5rem;color:var(--primary)}}

/* category grid */
.cat-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:18px;margin:22px 0 40px}}
.cat-card{{background:#fff;border:1px solid var(--line);border-radius:var(--radius);overflow:hidden;
  box-shadow:var(--shadow);transition:transform .15s,box-shadow .15s}}
.cat-card:hover{{transform:translateY(-3px);box-shadow:0 12px 30px rgba(70,28,93,.16)}}
.cat-thumb{{aspect-ratio:3/4;overflow:hidden;background:var(--second)}}
.cat-thumb img{{width:100%;height:100%;object-fit:cover}}
.cat-body{{padding:12px 14px 16px;position:relative}}
.cat-type{{font-weight:700;color:var(--primary);display:block}}
.cat-no{{position:absolute;top:12px;right:14px;font-size:.78rem;color:var(--muted)}}
.cat-size{{display:block;color:var(--muted);font-size:.82rem;margin-top:4px}}
.price-wrap{{margin-top:6px}}
.old{{text-decoration:line-through;color:var(--muted);margin-right:8px;font-size:.9rem}}
.new{{color:#b0163b;font-weight:800;font-size:1.15rem}}
.ptag{{font-size:.7rem;background:var(--primary);color:#fff;padding:2px 7px;border-radius:999px;vertical-align:middle}}

/* category page */
.cat-head{{margin:26px 0 10px}}
.crumb{{color:var(--muted);font-size:.85rem}}
.crumb a{{color:var(--primary)}}
.cat-head h1{{margin:.3rem 0;color:var(--primary);font-size:1.9rem}}
.hno{{color:var(--muted);font-weight:500;font-size:1.1rem}}
.cat-price{{font-size:1.2rem;margin:.2rem 0}}
.cat-size-big{{color:var(--muted);margin:.2rem 0 .6rem}}
.feat{{margin:.4rem 0 .6rem;padding-left:1.1rem;color:#3d2f45}}
.feat li{{margin:.2rem 0}}
.order-hint{{background:#fff;border:1px dashed var(--second);border-radius:12px;padding:10px 14px;color:#4a3a52;font-size:.92rem}}

.item-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:14px;margin:18px 0 30px}}
.item{{margin:0;background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;box-shadow:var(--shadow)}}
.item img{{aspect-ratio:3/4;width:100%;object-fit:cover;background:var(--second)}}
.item figcaption{{text-align:center;padding:8px;font-weight:700;color:var(--primary)}}
.back{{margin:10px 0 40px}}
.back a{{color:var(--primary);font-weight:600}}
.muted{{color:var(--muted)}}

/* footer */
.site-footer{{background:var(--primary-d);color:#e9e2ee;margin-top:30px;padding:34px 18px 20px}}
.foot-grid{{max-width:1120px;margin:0 auto;display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:22px}}
.foot-logo{{height:44px;background:#fff;border-radius:8px;padding:4px;width:auto}}
.foot-tag{{color:var(--second);margin-top:10px;font-size:.9rem}}
.site-footer h4{{color:#fff;margin:0 0 8px}}
.site-footer a{{color:var(--second)}}
.site-footer p{{margin:.3rem 0;font-size:.92rem}}
.copyright{{text-align:center;color:#a892b5;margin-top:24px;font-size:.82rem}}

@media(max-width:560px){{
  .hero h1{{font-size:2rem}}
  .cbtn span{{display:none}}
  .cbtn{{padding:11px 15px;font-size:1.1rem}}
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
