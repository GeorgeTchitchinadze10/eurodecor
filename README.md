# Eurodecor — Wallpaper Store website

Static catalog site for **Eurodecor** (შპალერების მაღაზია), Tbilisi.
Live: https://georgetchitchinadze10.github.io/eurodecor

## How it works
`build.py` reads the product catalog (`D:\Eurodecor\Eurodecor - Catalog.xlsx`)
and the category photo folders (`D:\Eurodecor\category N`), optimizes every
image to WebP, and generates the static site into `docs/` — which GitHub Pages
serves directly. No Node, no build tools.

## Rebuild after changing the Excel or adding photos
```
py build.py
git add -A && git commit -m "update catalog" && git push
```

## Edit business details (phone, hours, links)
See the `BIZ = {...}` block at the top of `build.py`, then rebuild.
