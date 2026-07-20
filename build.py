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
COVER_SRC = os.path.join(ROOT, "Eurodecor cover.png")
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
    "address_ru": "проспект Акакия Церетели 130, Тбилиси 0119",
    "maps": "https://maps.app.goo.gl/YRNTNxxe1ZpKK7hP9",
    "hours_ka": "ყოველდღე 10:00 – 19:00",
    "hours_en": "Every day 10:00 – 19:00",
    "hours_ru": "Ежедневно 10:00 – 19:00",
    "site_url": "https://georgetchitchinadze10.github.io/eurodecor",
}

C_PRIMARY = "#461c5d"
C_SECOND  = "#d1c2b9"

# ---------------------------------------------------------------- type + EN descriptions
TYPE_MAP = {
    "Wallpaper":            {"ka": "შპალერი",           "en": "Wallpaper",          "ru": "Обои",                    "key": "wallpaper"},
    "შესაღები ფლიზერინი":   {"ka": "შესაღები ფლიზელინი", "en": "Paintable flizeline","ru": "Флизелин под покраску",   "key": "paintable"},
    "Glue":                {"ka": "შპალერის წებო",      "en": "Wallpaper glue",     "ru": "Клей для обоев",          "key": "glue"},
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
RU_DESC = {
    "wallpaper": [
        "Винил на флизелиновой основе",
        "Моющиеся и влагостойкие",
        "Легко клеить",
        "Большой выбор современных и классических дизайнов",
    ],
    "paintable": [
        "Поверхность под покраску",
        "Легко клеить",
    ],
    "glue": [
        "Покрытие: 70–75 м²",
        "Универсальный клей для обоев",
        "Легко разводить и использовать",
        "Прочное сцепление",
        "Подходит для всех типов обоев",
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
        has_tag = "საბითუმო" in str(v)
        tag = {"ka": "საბითუმო", "en": "wholesale", "ru": "опт"}.get(lang, "") if has_tag else ""
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

def optimize_logo(src, dst, maxw=420, box=None):
    im = Image.open(src)
    im = ImageOps.exif_transpose(im)
    if box:
        im = im.crop(box)
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    w, h = im.size
    if w > maxw:
        im = im.resize((maxw, round(h * maxw / w)), Image.LANCZOS)
    im.save(dst, "WEBP", quality=90, method=6)

def make_thumb(src, dst, maxw=520, q=76):
    """Downscale an already-optimized image to a small card thumbnail."""
    im = Image.open(src)
    w, h = im.size
    if w > maxw:
        im = im.resize((maxw, round(h * maxw / w)), Image.LANCZOS)
    im.save(dst, "WEBP", quality=q, method=6)

def optimize_cover(src, dst, box=None, maxw=1000, q=82):
    """Crop the product scene out of the Facebook cover for the hero art."""
    im = Image.open(src)
    im = ImageOps.exif_transpose(im)
    if box:
        im = im.crop(box)
    im = im.convert("RGB")
    w, h = im.size
    if w > maxw:
        im = im.resize((maxw, round(h * maxw / w)), Image.LANCZOS)
    im.save(dst, "WEBP", quality=q, method=6)

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
        tinfo = TYPE_MAP.get(str(typ).strip(), {"ka": str(typ), "en": str(typ), "ru": str(typ), "key": "wallpaper"})
        cats.append({
            "no": no,
            "type_ka": tinfo["ka"], "type_en": tinfo["en"], "type_ru": tinfo.get("ru", tinfo["en"]), "key": tinfo["key"],
            "desc_ka": bullets_ka(desc),
            "desc_en": EN_DESC.get(tinfo["key"], []),
            "desc_ru": RU_DESC.get(tinfo["key"], []),
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
def i18n(ka, en, tag="span", ru=None):
    if ru is None:
        ru = en
    return (f'<{tag} class="ka">{ka}</{tag}>'
            f'<{tag} class="en">{en}</{tag}>'
            f'<{tag} class="ru">{ru}</{tag}>')

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
    "ruler": '<path d="M2.5 15.5 15.5 2.5a2 2 0 0 1 2.8 0l3.2 3.2a2 2 0 0 1 0 2.8L8.5 21.5a2 2 0 0 1-2.8 0l-3.2-3.2a2 2 0 0 1 0-2.8z"/><path d="m7.5 10.5 2 2"/><path d="m10.5 7.5 2 2"/><path d="m13.5 4.5 2 2"/><path d="m4.5 13.5 2 2"/>',
}

def icon(name, size=20):
    return (f'<svg class="ic" width="{size}" height="{size}" viewBox="0 0 24 24" '
            f'fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" '
            f'stroke-linejoin="round" aria-hidden="true">{ICON_PATHS[name]}</svg>')

# a single gold leaf (used in the line·leaf·line divider that echoes the logo lockup)
LEAF_SVG = ('<svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">'
            '<path d="M12 2C7.5 6 6 11.5 9 17c.6 1.1 1.7 2.4 3 5 1.3-2.6 2.4-3.9 3-5 3-5.5 1.5-11-3-15z"/>'
            '<path d="M12 6v11" stroke="rgba(0,0,0,.18)" stroke-width="1" fill="none"/></svg>')

def leaf_div(center=False, light=False):
    cls = "leaf-div" + (" center" if center else "") + (" light" if light else "")
    return f'<span class="{cls}"><i></i>{LEAF_SVG}<i></i></span>'

# a laurel sprig used as a faint corner flourish in the hero
def sprig_svg(cls="hero-sprig"):
    stem = '<path d="M100 22V184" stroke="currentColor" stroke-width="2.2" fill="none"/>'
    leaves = []
    for y in (46, 72, 98, 124, 150):
        leaves.append(f'<path d="M100 {y}c14-13 34-14 46-6-12 12-32 12-46 6z"/>')
        leaves.append(f'<path d="M100 {y+12}c-14-13-34-14-46-6 12 12 32 12 46 6z"/>')
    return (f'<svg class="{cls}" viewBox="0 0 200 200" fill="currentColor" aria-hidden="true">'
            f'{stem}{"".join(leaves)}</svg>')

FONT_URL = ("https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700"
            "&family=Noto+Serif+Georgian:wght@600;700"
            "&family=Noto+Sans+Georgian:wght@400;600&display=swap")

def head(title_ka, title_en, desc_ka, canonical, preload_img=None):
    preload = f'\n<link rel="preload" as="image" href="{preload_img}" fetchpriority="high">' if preload_img else ""
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
<meta property="og:type" content="website">{preload}
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="preload" as="style" href="{FONT_URL}">
<link rel="stylesheet" href="{FONT_URL}" media="print" onload="this.media='all'">
<noscript><link rel="stylesheet" href="{FONT_URL}"></noscript>
<link rel="stylesheet" href="styles.css">
<link rel="icon" href="assets/img/logo.webp">
<script type="application/ld+json">{ld}</script>'''

def header_html():
    return f'''<header class="site-header">
  <a class="brand" href="index.html">
    <span class="brand-mark"><img src="assets/img/mark.webp" alt="Eurodecor"></span>
    <span class="brand-text">
      <strong>{i18n("ევროდეკორი","Eurodecor","span",ru="Евродекор")}</strong>
      {i18n("შპალერების მაღაზია","Wallpaper Store","small",ru="Магазин обоев")}
    </span>
  </a>
  <div class="lang-switch" role="group" aria-label="Language">
    <button type="button" class="lang-opt" data-lang-set="ka">ქარ</button>
    <button type="button" class="lang-opt" data-lang-set="en">EN</button>
    <button type="button" class="lang-opt" data-lang-set="ru">РУ</button>
  </div>
</header>'''

def contact_bar():
    b = BIZ
    return f'''<nav class="contact-bar" aria-label="Contact">
  <a href="tel:{esc(b['phone_tel'])}" class="cbtn cbtn-call" aria-label="დარეკვა · Call">{icon('phone')}<span>{i18n("დარეკვა","Call",ru="Позвонить")}</span></a>
  <a href="https://wa.me/{esc(b['whatsapp'])}" class="cbtn cbtn-wa" target="_blank" rel="noopener" aria-label="WhatsApp">{icon('chat')}<span>WhatsApp</span></a>
  <a href="{esc(b['messenger'])}" class="cbtn cbtn-msg" target="_blank" rel="noopener" aria-label="Messenger">{icon('send')}<span>Messenger</span></a>
  <a href="{esc(b['maps'])}" class="cbtn cbtn-map" target="_blank" rel="noopener" aria-label="მისამართი · Directions">{icon('pin')}<span>{i18n("მისამართი","Directions",ru="Маршрут")}</span></a>
  <a href="mailto:{esc(b['email'])}" class="cbtn cbtn-mail" aria-label="Email">{icon('mail')}<span>Email</span></a>
</nav>'''

def calc_section():
    return f'''<section class="calc-section" aria-label="Wallpaper calculator">
  <div class="container">
    <div class="calc-card">
      <div class="calc-head">
        <span class="calc-ic">{icon('ruler', 24)}</span>
        <div>
          <h2 class="calc-title">{i18n("შპალერის კალკულატორი","Wallpaper calculator","span",ru="Калькулятор обоев")}</h2>
          <p class="calc-sub">{i18n("გამოთვალეთ რამდენი რულონი დაგჭირდებათ","Estimate how many rolls you need","span",ru="Рассчитайте, сколько рулонов вам нужно")}</p>
        </div>
      </div>
      <div class="calc-grid">
        <label class="calc-field"><span>{i18n("ოთახის სიგრძე (მ)","Room length (m)","span",ru="Длина комнаты (м)")}</span>
          <input type="number" id="calcL" min="0" step="0.1" inputmode="decimal" placeholder="4"></label>
        <label class="calc-field"><span>{i18n("ოთახის სიგანე (მ)","Room width (m)","span",ru="Ширина комнаты (м)")}</span>
          <input type="number" id="calcW" min="0" step="0.1" inputmode="decimal" placeholder="3"></label>
        <label class="calc-field"><span>{i18n("კედლის სიმაღლე (მ)","Wall height (m)","span",ru="Высота стены (м)")}</span>
          <input type="number" id="calcH" min="0" step="0.1" inputmode="decimal" placeholder="2.7"></label>
        <label class="calc-field"><span>{i18n("რულონის ზომა","Roll size","span",ru="Размер рулона")}</span>
          <select id="calcRW"><option value="1.06">1.06 m × 10 m</option><option value="0.53">0.53 m × 10 m</option></select></label>
      </div>
      <button type="button" id="calcBtn" class="btn btn-gold calc-btn">{icon('ruler')}<span>{i18n("გამოთვლა","Calculate",ru="Рассчитать")}</span></button>
      <div class="calc-result" id="calcResult" hidden></div>
    </div>
  </div>
</section>'''

def map_section():
    b = BIZ
    coords = "41.74599397323277,44.777773328922635"
    embed = f"https://www.google.com/maps?q={coords}&z=17&output=embed&hl=ka"
    return f'''<section class="map-section" aria-label="Location">
  <div class="container map-inner">
    <div class="map-head">
      <div>
        <h2 class="map-title">{i18n("მისამართი","Address","span",ru="Адрес")}</h2>
        <p class="map-addr">{i18n(b['address_ka'], b['address_en'],"span",ru=b['address_ru'])}</p>
        <p class="map-hours">{i18n(b['hours_ka'], b['hours_en'],"span",ru=b['hours_ru'])} <span class="open-status" data-openstatus></span></p>
      </div>
      <a href="{esc(b['maps'])}" class="btn btn-gold map-btn" target="_blank" rel="noopener">{icon('pin')}<span>{i18n("რუკაზე ნახვა","View on map",ru="Смотреть на карте")}</span></a>
    </div>
    <div class="map-frame">
      <iframe src="{embed}" width="100%" height="380" style="border:0" loading="lazy" allowfullscreen referrerpolicy="no-referrer-when-downgrade" title="Eurodecor — რუკა"></iframe>
    </div>
  </div>
</section>'''

def footer_html():
    b = BIZ
    return f'''<footer class="site-footer">
  <div class="foot-grid">
    <div>
      <span class="foot-mark"><img src="assets/img/mark.webp" alt="Eurodecor"></span>
      <p class="foot-tag">{i18n("პირდაპირ ქარხნიდან · 25 წლიანი გამოცდილება","Factory-direct · 25 years of experience","span",ru="Напрямую с фабрики · 25 лет опыта")}</p>
    </div>
    <div>
      <h2>{i18n("კონტაქტი","Contact","span",ru="Контакты")}</h2>
      <p>{i18n(b['address_ka'], b['address_en'],"span",ru=b['address_ru'])}</p>
      <p><a href="tel:{esc(b['phone_tel'])}">{esc(b['phone_display'])}</a></p>
      <p><a href="mailto:{esc(b['email'])}">{esc(b['email'])}</a></p>
      <p>{i18n(b['hours_ka'], b['hours_en'],"span",ru=b['hours_ru'])} <span class="open-status" data-openstatus></span></p>
    </div>
    <div>
      <h2>{i18n("გვეწვიეთ","Visit us","span",ru="Посетите нас")}</h2>
      <p><a href="{esc(b['maps'])}" target="_blank" rel="noopener">{i18n("რუკაზე ნახვა →","Open in Maps →","span",ru="Открыть на карте →")}</a></p>
      <p><a href="{esc(b['facebook'])}" target="_blank" rel="noopener">Facebook →</a></p>
    </div>
  </div>
  <p class="copyright">© 2026 Eurodecor · {i18n("ყველა უფლება დაცულია","All rights reserved","span",ru="Все права защищены")}</p>
</footer>
<div class="sticky-contact" aria-label="Quick contact">
  <a href="https://wa.me/{esc(b['whatsapp'])}" class="sc-btn sc-wa" target="_blank" rel="noopener" aria-label="WhatsApp">{icon('chat', 26)}</a>
  <a href="tel:{esc(b['phone_tel'])}" class="sc-btn sc-call" aria-label="დარეკვა · Call">{icon('phone', 26)}</a>
</div>
{STATUS_JS}'''

LANG_JS = '''<script>
(function(){
  var root=document.documentElement;
  var valid={ka:1,en:1,ru:1};
  var saved=localStorage.getItem('lang');
  if(!valid[saved])saved='ka';
  var opts=document.querySelectorAll('[data-lang-set]');
  function apply(l){
    root.setAttribute('data-lang',l);
    root.setAttribute('lang',l);
    localStorage.setItem('lang',l);
    for(var i=0;i<opts.length;i++){
      opts[i].setAttribute('aria-pressed',opts[i].getAttribute('data-lang-set')===l?'true':'false');
    }
  }
  apply(saved);
  for(var i=0;i<opts.length;i++){
    opts[i].addEventListener('click',function(){apply(this.getAttribute('data-lang-set'));});
  }
})();
</script>'''

STATUS_JS = '''<script>
(function(){
  var els=document.querySelectorAll('[data-openstatus]');
  if(!els.length)return;
  var h;
  try{h=parseInt(new Intl.DateTimeFormat('en-GB',{hour:'2-digit',hour12:false,timeZone:'Asia/Tbilisi'}).format(new Date()),10);}
  catch(e){h=new Date().getHours();}
  h=h%24;
  var open=h>=10&&h<19;
  var html=open
    ?'<span class="ka">ღიაა ახლა</span><span class="en">Open now</span><span class="ru">Открыто</span>'
    :'<span class="ka">დაკეტილია</span><span class="en">Closed</span><span class="ru">Закрыто</span>';
  for(var i=0;i<els.length;i++){
    els[i].className='open-status '+(open?'open-status--open':'open-status--closed');
    els[i].innerHTML=html;
  }
})();
</script>'''

CALC_JS = '''<script>
(function(){
  var btn=document.getElementById('calcBtn');
  if(!btn)return;
  var res=document.getElementById('calcResult');
  function num(id){var v=parseFloat(document.getElementById(id).value);return isNaN(v)?0:v;}
  function warn(ka,en,ru){res.hidden=false;res.className='calc-result calc-result--warn';
    res.innerHTML='<span class="ka">'+ka+'</span><span class="en">'+en+'</span><span class="ru">'+ru+'</span>';}
  function ruRolls(n){var m=n%10,h=n%100;
    if(m===1&&h!==11)return 'рулон';
    if(m>=2&&m<=4&&(h<10||h>=20))return 'рулона';
    return 'рулонов';}
  function calc(){
    var L=num('calcL'),W=num('calcW'),H=num('calcH');
    var rw=parseFloat(document.getElementById('calcRW').value)||1.06;
    var rollLen=10;
    if(L<=0||W<=0||H<=0){return warn('შეავსეთ ყველა ველი სწორად.','Please fill in all fields correctly.','Заполните все поля правильно.');}
    if(H>rollLen){return warn('კედელი ძალიან მაღალია ამ რულონისთვის.','Wall is too tall for this roll.','Стена слишком высокая для этого рулона.');}
    var perimeter=2*(L+W);
    var strips=Math.ceil(perimeter/rw);
    var stripsPerRoll=Math.floor(rollLen/H);
    var rolls=Math.ceil(strips/stripsPerRoll);
    var withExtra=rolls+1;
    res.hidden=false;res.className='calc-result';
    res.innerHTML=
      '<div class="calc-big"><span class="calc-num">'+rolls+'</span>'
      +'<span class="ka">&nbsp;რულონი</span><span class="en">&nbsp;rolls</span><span class="ru">&nbsp;'+ruRolls(rolls)+'</span></div>'
      +'<p class="calc-rec"><span class="ka">გირჩევთ აიღოთ '+withExtra+' რულონი — მარაგი ჭრისა და ნახატის დასამთხვევად.</span>'
      +'<span class="en">We recommend buying '+withExtra+' rolls — one spare for cutting and pattern matching.</span>'
      +'<span class="ru">Рекомендуем купить '+withExtra+' '+ruRolls(withExtra)+' — один запасной для подрезки и подгонки рисунка.</span></p>'
      +'<p class="calc-hint"><span class="ka">პერიმეტრი ≈ '+perimeter.toFixed(1)+' მ · '+strips+' ზოლი · '+stripsPerRoll+' ზოლი / რულონი</span>'
      +'<span class="en">Perimeter ≈ '+perimeter.toFixed(1)+' m · '+strips+' strips · '+stripsPerRoll+' strips / roll</span>'
      +'<span class="ru">Периметр ≈ '+perimeter.toFixed(1)+' м · '+strips+' полос · '+stripsPerRoll+' полос / рулон</span></p>';
  }
  btn.addEventListener('click',calc);
  ['calcL','calcW','calcH'].forEach(function(id){
    document.getElementById(id).addEventListener('keydown',function(e){if(e.key==='Enter')calc();});
  });
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
        first_img = c.get("thumb", "assets/img/logo.webp")
        badge = f'<span class="badge">{i18n("ფასდაკლება","Sale",ru="Скидка")}</span>' if c["old"] else ""
        cards.append(f'''<a class="cat-card" href="category-{c['no']}.html">
      <div class="cat-thumb">{badge}
        <img loading="lazy" decoding="async" width="520" height="693" src="{first_img}" alt="{esc(c['type_en'])} {c['no']}">
      </div>
      <div class="cat-body">
        <span class="cat-type">{i18n(c['type_ka'], c['type_en'],ru=c['type_ru'])}</span>
        <div class="price-wrap"><span class="ka">{price_html(c["old"], c["new"], "ka")}</span><span class="en">{price_html(c["old"], c["new"], "en")}</span><span class="ru">{price_html(c["old"], c["new"], "ru")}</span></div>
        <span class="cat-size">{esc(c['size'])} · #{c['no']}</span>
      </div>
    </a>''')
    grid = "\n".join(cards)
    hero = f'''<section class="hero">
    {sprig_svg()}
    <div class="hero-grid">
      <div class="hero-copy">
        <span class="hero-over">{i18n("· 25 წლიანი გამოცდილება ·","· EST. 25 YEARS ·","span",ru="· 25 ЛЕТ ОПЫТА ·")}</span>
        <h1>{i18n("ევროდეკორი","Eurodecor","span",ru="Евродекор")}</h1>
        {leaf_div(light=True)}
        <p class="hero-sub">{i18n("შპალერების მაღაზია","Wallpaper Store","span",ru="Магазин обоев")}</p>
        <p class="hero-usp">{i18n("პირდაპირ ქარხნიდან · საუკეთესო ფასები თბილისში · დიდი არჩევანი","Factory-direct · best prices in Tbilisi · huge selection","span",ru="Напрямую с фабрики · лучшие цены в Тбилиси · огромный выбор")}</p>
        <div class="hero-cta">
          <a class="btn btn-gold" href="https://wa.me/{esc(BIZ['whatsapp'])}" target="_blank" rel="noopener">{icon('chat')}<span>{i18n("მოგვწერეთ","Message us",ru="Напишите нам")}</span></a>
          <a class="btn btn-ghost" href="#categories">{i18n("კატალოგის ნახვა","Browse catalog",ru="Смотреть каталог")}</a>
        </div>
      </div>
      <div class="hero-art"><img src="assets/img/hero.webp" width="1000" height="762" fetchpriority="high" decoding="async" alt="Eurodecor — ევროდეკორის შპალერები"></div>
    </div>
  </section>'''
    body = f'''{header_html()}
{hero}
{contact_bar()}
<main class="container">
  <h2 class="section-title" id="categories">{i18n("კატეგორიები","Categories","span",ru="Категории")}</h2>
  {leaf_div(center=True)}
  <div class="cat-grid">
    {grid}
  </div>
</main>
{calc_section()}
{map_section()}
{footer_html()}
{LANG_JS}
{CALC_JS}'''
    title_ka = "ევროდეკორი — შპალერების მაღაზია თბილისში"
    title_en = "Eurodecor — Wallpaper Store in Tbilisi"
    desc = "შპალერი, ვინილის შპალერი, ფლიზელინი, შესაღები შპალერი და შპალერის წებო — საუკეთესო ფასებში, პირდაპირ ქარხნიდან. თბილისი, აკაკი წერეთლის 130."
    return f'''<!doctype html><html lang="ka" data-lang="ka"><head>
{head(title_ka, title_en, desc, BIZ["site_url"] + "/", preload_img="assets/img/hero.webp")}
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
    grid = "\n".join(items) if items else f'<p class="muted">{i18n("ფოტოები მალე","Photos coming soon",ru="Фото скоро")}</p>'

    d_ka = "".join(f"<li>{esc(x)}</li>" for x in c["desc_ka"])
    d_en = "".join(f"<li>{esc(x)}</li>" for x in c["desc_en"])
    d_ru = "".join(f"<li>{esc(x)}</li>" for x in c["desc_ru"])
    badge = f'<span class="badge badge-dark">{i18n("ფასდაკლება","Sale",ru="Скидка")}</span>' if c["old"] else ""

    intro = f'''<section class="cat-head">
    <nav class="crumb"><a href="index.html">{i18n("მთავარი","Home",ru="Главная")}</a> <span>/</span> #{c['no']}</nav>
    <span class="cat-over">{i18n(c['type_ka'], c['type_en'],ru=c['type_ru'])} · #{c['no']}</span>
    <h1 class="cat-title">{i18n(c['type_ka'], c['type_en'],"span",ru=c['type_ru'])}</h1>
    {leaf_div(center=True)}
    <div class="cat-price">{badge}
      <span class="ka">{price_html(c['old'], c['new'], 'ka')}</span><span class="en">{price_html(c['old'], c['new'], 'en')}</span><span class="ru">{price_html(c['old'], c['new'], 'ru')}</span>
    </div>
    <p class="cat-size-big">{esc(c['size'])}</p>
    <ul class="feat ka">{d_ka}</ul>
    <ul class="feat en">{d_en}</ul>
    <ul class="feat ru">{d_ru}</ul>
    <p class="order-hint">{i18n("მოგწონთ რომელიმე? მოგვწერეთ ნომერი (მაგ. #"+c['no']+"01) Messenger-ში ან WhatsApp-ში.","Like one? Message us the number (e.g. #"+c['no']+"01) on Messenger or WhatsApp.",ru="Понравился какой-то? Напишите нам номер (напр. #"+c['no']+"01) в Messenger или WhatsApp.")}</p>
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
  <p class="back"><a href="index.html">← {i18n("ყველა კატეგორია","All categories",ru="Все категории")}</a></p>
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
  --royal:#5a2288;
  --royal-2:#712fa6;
  --plum-d:#2a0f3e;
  --taupe:{C_SECOND};
  --taupe-l:#e6ddd6;
  --cream:#f5efe7;
  --cream-2:#efe6da;
  --ink:#2c2333;
  --muted:#6e6377;
  --gold:#c2a15c;
  --gold-l:#e4cf95;
  --sale:#8d2846;
  --line:rgba(70,28,93,.12);
  --serif:'Playfair Display','Noto Serif Georgian',Georgia,serif;
  --sans:'Noto Sans Georgian','Segoe UI',system-ui,-apple-system,sans-serif;
}}
*{{box-sizing:border-box}}
html{{-webkit-text-size-adjust:100%;scroll-behavior:smooth}}
body{{margin:0;font-family:var(--sans);color:var(--ink);background:var(--cream);line-height:1.55}}
img{{max-width:100%;display:block}}
a{{color:inherit;text-decoration:none}}
.ic{{flex:none}}
.container{{max-width:1160px;margin:0 auto;padding:0 22px}}

/* language toggle */
html[data-lang="ka"] .en,html[data-lang="ka"] .ru{{display:none !important}}
html[data-lang="en"] .ka,html[data-lang="en"] .ru{{display:none !important}}
html[data-lang="ru"] .ka,html[data-lang="ru"] .en{{display:none !important}}

/* gold leaf divider (echoes the logo lockup) */
.leaf-div{{display:flex;align-items:center;gap:12px;margin:16px 0}}
.leaf-div.center{{justify-content:center;margin:14px 0 4px}}
.leaf-div i{{display:block;height:1px;width:44px;background:linear-gradient(90deg,transparent,var(--gold))}}
.leaf-div i:last-child{{background:linear-gradient(90deg,var(--gold),transparent)}}
.leaf-div svg{{width:16px;height:16px;color:var(--gold);flex:none}}
.leaf-div.light svg{{color:var(--gold-l)}}
.leaf-div.light i{{background:linear-gradient(90deg,transparent,rgba(228,207,149,.9))}}
.leaf-div.light i:last-child{{background:linear-gradient(90deg,rgba(228,207,149,.9),transparent)}}

/* header */
.site-header{{position:sticky;top:0;z-index:40;display:flex;align-items:center;justify-content:space-between;
  padding:11px 22px;background:rgba(245,239,231,.88);backdrop-filter:blur(10px);
  border-bottom:1px solid var(--line);color:var(--plum)}}
.brand{{display:flex;align-items:center;gap:12px}}
.brand-mark{{display:block;width:46px;height:46px;border-radius:12px;overflow:hidden;flex:none;
  box-shadow:0 4px 12px rgba(70,28,93,.22)}}
.brand-mark img{{width:100%;height:100%;object-fit:cover}}
.brand-text{{display:flex;flex-direction:column;line-height:1.05}}
.brand-text strong{{font-family:var(--serif);font-size:1.3rem;font-weight:700;letter-spacing:.3px;color:var(--plum)}}
.brand-text small{{font-size:.68rem;letter-spacing:.16em;text-transform:uppercase;color:var(--muted);margin-top:3px}}
.lang-switch{{display:inline-flex;border:1px solid var(--plum);border-radius:999px;overflow:hidden}}
.lang-opt{{background:transparent;color:var(--plum);border:0;padding:6px 11px;font-weight:600;cursor:pointer;
  font-size:.78rem;letter-spacing:.04em;transition:.15s;font-family:inherit;line-height:1.2}}
.lang-opt+.lang-opt{{border-left:1px solid rgba(70,28,93,.25)}}
.lang-opt:hover{{background:rgba(70,28,93,.09)}}
.lang-opt[aria-pressed="true"]{{background:var(--plum);color:#fff}}

/* hero */
.hero{{position:relative;overflow:hidden;color:#fff;
  background:linear-gradient(120deg,#39144f 0%,var(--royal) 52%,#45206c 100%)}}
.hero-grid{{position:relative;z-index:2;max-width:1220px;margin:0 auto;
  display:grid;grid-template-columns:1.02fr .98fr;align-items:stretch;min-height:clamp(380px,50vw,540px)}}
.hero-copy{{align-self:center;padding:54px clamp(22px,4vw,60px)}}
.hero-over{{display:block;letter-spacing:.32em;text-transform:uppercase;font-size:.72rem;color:var(--gold-l);font-weight:600}}
.hero h1{{font-family:var(--serif);font-weight:600;font-size:clamp(2.6rem,5.4vw,4rem);margin:.4rem 0 0;letter-spacing:.5px;line-height:1.05}}
.hero-sub{{font-family:var(--serif);font-style:italic;color:var(--taupe);font-size:clamp(1.1rem,2.3vw,1.5rem);margin:.2rem 0 0}}
.hero-usp{{color:rgba(255,255,255,.85);font-size:1rem;letter-spacing:.02em;margin:14px 0 0;max-width:30em}}
.hero-cta{{display:flex;flex-wrap:wrap;gap:12px;margin-top:26px}}
.btn{{display:inline-flex;align-items:center;gap:9px;padding:13px 26px;border-radius:999px;
  font-weight:600;font-size:.96rem;transition:.18s;cursor:pointer}}
.btn-gold{{background:linear-gradient(135deg,var(--gold-l),var(--gold));color:#3a2410;
  box-shadow:0 10px 26px rgba(0,0,0,.28)}}
.btn-gold:hover{{transform:translateY(-2px);box-shadow:0 16px 34px rgba(0,0,0,.34)}}
.btn-ghost{{border:1px solid rgba(255,255,255,.55);color:#fff}}
.btn-ghost:hover{{background:rgba(255,255,255,.12);border-color:#fff}}
.hero-art{{position:relative;align-self:stretch;min-height:260px}}
.hero-art img{{position:absolute;inset:0;width:100%;height:100%;object-fit:cover;object-position:center;
  -webkit-mask-image:linear-gradient(90deg,transparent 0,#000 10%);
  mask-image:linear-gradient(90deg,transparent 0,#000 10%)}}
.hero-sprig{{position:absolute;left:-26px;bottom:-40px;width:230px;z-index:1;color:var(--gold-l);opacity:.13;pointer-events:none}}

/* contact bar */
.contact-bar{{display:flex;gap:12px;justify-content:center;flex-wrap:wrap;padding:22px 14px;background:var(--cream-2)}}
.cbtn{{display:inline-flex;align-items:center;gap:9px;background:#fff;color:var(--plum);
  padding:11px 20px;border-radius:999px;border:1px solid var(--line);font-weight:600;font-size:.92rem;
  box-shadow:0 2px 10px rgba(70,28,93,.05);transition:.18s}}
.cbtn:hover{{background:var(--plum);color:#fff;border-color:var(--plum);transform:translateY(-2px);box-shadow:0 12px 24px rgba(70,28,93,.18)}}
.cbtn span{{white-space:nowrap}}

/* section */
.section-title{{font-family:var(--serif);text-align:center;margin:52px 0 0;font-size:2.1rem;font-weight:600;color:var(--plum);scroll-margin-top:80px}}

/* category grid */
.cat-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:28px;margin:30px 0 66px}}
.cat-card{{background:#fff;border:1px solid var(--line);border-radius:16px 16px 14px 14px;overflow:hidden;
  box-shadow:0 6px 22px rgba(70,28,93,.07);transition:transform .22s,box-shadow .22s}}
.cat-card:hover{{transform:translateY(-5px);box-shadow:0 24px 48px rgba(70,28,93,.17)}}
.cat-thumb{{position:relative;aspect-ratio:3/4;overflow:hidden;background:var(--royal);border-radius:15px 15px 0 0}}
.cat-thumb img{{width:100%;height:100%;object-fit:cover;transition:transform .55s ease}}
.cat-card:hover .cat-thumb img{{transform:scale(1.06)}}
.cat-thumb::after{{content:"";position:absolute;inset:10px;z-index:2;pointer-events:none;
  border:1px solid rgba(226,207,149,.85);border-radius:125px 125px 6px 6px;
  box-shadow:0 0 0 1px rgba(0,0,0,.05) inset}}
.badge{{position:absolute;top:14px;left:14px;z-index:3;background:linear-gradient(135deg,var(--gold-l),var(--gold));color:#3a2410;
  font-size:.64rem;letter-spacing:.12em;text-transform:uppercase;font-weight:700;padding:5px 12px;border-radius:999px;
  box-shadow:0 3px 10px rgba(0,0,0,.18)}}
.cat-body{{padding:16px 18px 20px;text-align:center}}
.cat-type{{display:block;font-weight:600;color:var(--ink);font-size:1.06rem}}
.price-wrap{{margin:9px 0 5px}}
.old{{text-decoration:line-through;color:var(--muted);margin-right:9px;font-size:.95rem}}
.new{{font-family:var(--serif);color:var(--plum);font-weight:700;font-size:1.5rem}}
.ptag{{font-size:.62rem;letter-spacing:.1em;text-transform:uppercase;background:var(--gold);color:#3a2410;
  padding:3px 8px;border-radius:999px;vertical-align:middle;margin-left:7px}}
.cat-size{{display:block;color:var(--muted);font-size:.82rem;margin-top:6px;letter-spacing:.03em}}

/* category page */
.cat-head{{max-width:760px;margin:36px auto 8px;text-align:center}}
.crumb{{color:var(--muted);font-size:.84rem}}
.crumb a{{color:var(--plum)}}
.crumb span{{margin:0 5px;opacity:.6}}
.cat-head .leaf-div{{margin-top:6px}}
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
.item-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:20px;margin:28px 0 36px}}
.item{{margin:0;cursor:pointer;outline:none}}
.item-img{{position:relative;aspect-ratio:3/4;border-radius:14px 14px 8px 8px;overflow:hidden;background:var(--royal);
  box-shadow:0 6px 20px rgba(70,28,93,.09)}}
.item-img img{{width:100%;height:100%;object-fit:cover;transition:transform .45s}}
.item-img::after{{content:"";position:absolute;inset:8px;z-index:2;pointer-events:none;
  border:1px solid rgba(226,207,149,.8);border-radius:110px 110px 5px 5px}}
.item:hover .item-img img,.item:focus .item-img img{{transform:scale(1.07)}}
.zoom{{position:absolute;inset:0;z-index:3;display:flex;align-items:center;justify-content:center;color:#fff;
  background:rgba(44,17,64,.32);opacity:0;transition:.2s}}
.item:hover .zoom,.item:focus-visible .zoom{{opacity:1}}
.item figcaption{{text-align:center;padding:9px 0 0;font-weight:600;color:var(--plum);font-size:.9rem;letter-spacing:.05em}}
.back{{margin:6px 0 44px}}
.back a{{color:var(--plum);font-weight:600}}
.muted{{color:var(--muted);text-align:center}}

/* calculator */
.calc-section{{padding:26px 0 8px}}
.calc-card{{background:#fff;border:1px solid var(--taupe);border-radius:20px;padding:26px 24px;
  box-shadow:0 14px 34px rgba(70,28,93,.10);position:relative}}
.calc-card::after{{content:"";position:absolute;inset:0;border-radius:20px;pointer-events:none;
  box-shadow:inset 0 0 0 2px rgba(194,161,92,.28)}}
.calc-head{{display:flex;align-items:center;gap:14px;margin-bottom:20px}}
.calc-ic{{display:flex;align-items:center;justify-content:center;width:46px;height:46px;flex:0 0 auto;
  border-radius:12px;background:var(--plum);color:var(--gold-l,#e4cf95)}}
.calc-title{{font-family:var(--serif);color:var(--plum);font-weight:700;font-size:1.4rem;margin:0}}
.calc-sub{{color:var(--muted);font-size:.92rem;margin:2px 0 0}}
.calc-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px 16px;margin-bottom:20px}}
.calc-field{{display:flex;flex-direction:column;gap:6px;font-size:.86rem;font-weight:600;color:var(--ink)}}
.calc-field input,.calc-field select{{font:inherit;font-weight:500;padding:11px 12px;border:1px solid var(--taupe);
  border-radius:10px;background:var(--cream);color:var(--ink);width:100%}}
.calc-field input:focus,.calc-field select:focus{{outline:none;border-color:var(--plum);
  box-shadow:0 0 0 3px rgba(90,34,136,.15)}}
.calc-btn{{width:100%;justify-content:center}}
.calc-result{{margin-top:20px;padding:20px;border-radius:14px;background:var(--cream);
  border:1px dashed var(--gold);text-align:center}}
.calc-result--warn{{border-style:solid;border-color:#c14b4b;background:rgba(193,75,75,.06);color:#a24141;font-weight:600}}
.calc-big{{font-family:var(--serif);color:var(--plum);font-weight:700;font-size:1.5rem}}
.calc-num{{font-size:2.6rem;line-height:1;color:var(--royal,#5a2288)}}
.calc-rec{{color:var(--ink);font-weight:600;margin:10px 0 0;font-size:.95rem}}
.calc-hint{{color:var(--muted);font-size:.82rem;margin:8px 0 0}}

/* map */
.map-section{{padding:10px 0 6px}}
.map-inner{{margin-top:8px}}
.map-head{{display:flex;align-items:flex-end;justify-content:space-between;gap:18px;flex-wrap:wrap;margin-bottom:16px}}
.map-title{{font-family:var(--serif);color:var(--plum);font-weight:700;font-size:1.5rem;margin:0}}
.map-addr{{color:var(--ink);font-weight:600;margin:6px 0 2px}}
.map-hours{{color:var(--muted);font-size:.92rem;margin:0}}
.map-btn{{flex:0 0 auto}}
.map-frame{{position:relative;border-radius:16px;overflow:hidden;border:1px solid var(--taupe);
  box-shadow:0 14px 34px rgba(70,28,93,.16)}}
.map-frame::after{{content:"";position:absolute;inset:0;border-radius:16px;pointer-events:none;
  box-shadow:inset 0 0 0 3px rgba(194,161,92,.35)}}
.map-frame iframe{{display:block}}
@media(max-width:560px){{.map-head{{align-items:flex-start}}.map-btn{{width:100%;justify-content:center}}.map-title{{font-size:1.3rem}}}}

/* open/closed status badge */
.open-status{{display:none;align-items:center;gap:6px;margin-left:8px;padding:2px 10px;border-radius:999px;
  font-size:.78rem;font-weight:700;letter-spacing:.02em;vertical-align:middle;white-space:nowrap}}
.open-status--open,.open-status--closed{{display:inline-flex}}
.open-status::before{{content:"";width:7px;height:7px;border-radius:50%}}
.open-status--open{{background:rgba(46,140,74,.14);color:#1f7a3d}}
.open-status--open::before{{background:#25a04c;box-shadow:0 0 0 3px rgba(37,160,76,.22)}}
.open-status--closed{{background:rgba(150,60,60,.12);color:#a24141}}
.open-status--closed::before{{background:#c14b4b}}
.site-footer .open-status--open{{background:rgba(255,255,255,.12);color:#8fe6a6}}
.site-footer .open-status--closed{{background:rgba(255,255,255,.1);color:#e8a9a9}}

/* sticky mobile contact */
.sticky-contact{{display:none}}
@media(max-width:820px){{
  .sticky-contact{{display:flex;flex-direction:column;gap:12px;position:fixed;right:16px;
    bottom:calc(16px + env(safe-area-inset-bottom));z-index:200}}
  .sc-btn{{display:flex;align-items:center;justify-content:center;width:54px;height:54px;border-radius:50%;
    color:#fff;box-shadow:0 8px 20px rgba(0,0,0,.28)}}
  .sc-btn svg{{width:26px;height:26px}}
  .sc-wa{{background:#25d366}}
  .sc-call{{background:var(--plum)}}
}}

/* footer */
.site-footer{{position:relative;background:var(--plum-d);color:#e7dcee;margin-top:24px;padding:46px 22px 26px;
  border-top:2px solid transparent;border-image:linear-gradient(90deg,transparent,var(--gold),transparent) 1}}
.foot-grid{{max-width:1160px;margin:0 auto;display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:30px}}
.foot-mark{{display:block;width:56px;height:56px;border-radius:14px;overflow:hidden;
  box-shadow:0 6px 16px rgba(0,0,0,.35)}}
.foot-mark img{{width:100%;height:100%;object-fit:cover}}
.foot-tag{{color:var(--taupe);margin-top:12px;font-size:.9rem;max-width:230px}}
.site-footer h2{{font-family:var(--serif);color:#fff;margin:0 0 10px;font-weight:600;font-size:1.08rem}}
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

@media(max-width:820px){{
  .hero-grid{{grid-template-columns:1fr;min-height:0}}
  .hero-copy{{order:2;text-align:center;padding:38px 22px 34px}}
  .hero-over,.hero-usp{{margin-left:auto;margin-right:auto}}
  .leaf-div{{justify-content:center}}
  .hero-cta{{justify-content:center}}
  .hero-art{{order:1;height:230px;min-height:230px}}
  .hero-art img{{-webkit-mask-image:linear-gradient(180deg,#000 78%,transparent);
    mask-image:linear-gradient(180deg,#000 78%,transparent)}}
  .hero-sprig{{display:none}}
}}
@media(max-width:560px){{
  .cbtn span{{display:none}}
  .cbtn{{padding:12px}}
  .cat-grid{{gap:16px;grid-template-columns:repeat(auto-fill,minmax(150px,1fr))}}
  .cat-thumb::after{{inset:7px;border-radius:90px 90px 5px 5px}}
  .item-grid{{grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:14px}}
  .item-img::after{{inset:6px;border-radius:80px 80px 4px 4px}}
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

    # full logo (favicon / schema / social) + cropped emblem chip for the header & footer
    optimize_logo(LOGO_SRC, os.path.join(IMG_OUT, "logo.webp"))
    optimize_logo(LOGO_SRC, os.path.join(IMG_OUT, "mark.webp"), maxw=200, box=(348, 244, 852, 748))
    # hero art = the product scene cropped from the Facebook cover
    optimize_cover(COVER_SRC, os.path.join(IMG_OUT, "hero.webp"), box=(880, 0, 1942, 809))

    cats = read_categories()
    for c in cats:
        c["items"] = build_items(c)
        # small thumbnail for the home grid (the full item photo is far too big for a card)
        if c["items"]:
            src = os.path.join(OUT, c["items"][0]["img"])
            thumb_rel = f"assets/img/card-{c['no']}.webp"
            make_thumb(src, os.path.join(OUT, thumb_rel))
            c["thumb"] = thumb_rel
        else:
            c["thumb"] = "assets/img/logo.webp"
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
