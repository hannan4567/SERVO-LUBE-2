"""update-here/*.xlsx -> data/products.json (run automatically by GitHub Actions)."""
import glob, json, os, re, sys, datetime as dt
import openpyxl

files = sorted(glob.glob("update-here/*.xlsx"), key=os.path.getmtime)
if not files:
    print("No Excel in update-here/ - nothing to do"); sys.exit(0)   # no red cross when you only delete
ws = openpyxl.load_workbook(files[-1], data_only=True).active
rows = list(ws.iter_rows())
head = [str(c.value or "").strip().lower() for c in rows[0]]
def ix(*keys, need=True):
    for i, h in enumerate(head):
        if any(h.startswith(k) for k in keys): return i
    if need: sys.exit(f"Excel column missing: {keys[0]}")
I = dict(code=ix("material code"), sku=ix("sku"), unit=ix("eun"), seg=ix("segment"), pack=ix("pack"),
         size=ix("pack size"), mrp=ix("mrp"), dlp=ix("dlp"), disc=ix("prod discount"), free=ix("premium free"),
         offer=ix("offer", need=False), valid=ix("valid", need=False))

SEG = {"CVO": "Commercial Vehicle", "PCMO": "Passenger Car", "MCO": "Motorcycle", "2T": "2-Stroke",
       "GREASE": "Grease", "AGRI": "Agri & Tractor"}
GUESS = [(r"GEAR|TRANS|ATF|MGGO|AXLE", "Gear & Transmission"), (r"COOL|BRAKE|KOOL", "Coolant & Brake"),
         (r"\b(2T|4T)\b|SCOOT", "Motorcycle"), (r"TRACTOR|GREENMILE|RAFTAAR", "Agri & Tractor"),
         (r"SYSTEM|HYDREX|MESH|CUT|TURB|COMP|THERM|QUENCH", "Industrial")]
def segment(s, name):
    s = str(s or "").strip()
    if s.upper() in SEG: return SEG[s.upper()]
    if s: return s
    return next((l for p, l in GUESS if re.search(p, name.upper())), "Other")

PACK = re.compile(r"\s*[-–]?\s*(\d+\s*[xX]\s*)?[\d./]+\s*(ML|L|LTR|KG|G|GM)\b.*$", re.I)
num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
out = []
for r in rows[1:]:
    code, name = r[I["code"]].value, r[I["sku"]].value
    if not code or not name: continue
    name = re.sub(r"\s+", " ", str(name)).strip()
    mrp, dlp, size = num(r[I["mrp"]].value), num(r[I["dlp"]].value), num(r[I["size"]].value)
    disc, free = num(r[I["disc"]].value), num(r[I["free"]].value)
    ex = dlp / 1.18 - disc * size - free           # dealer rate before GST, after all schemes
    offers = []
    if I["offer"] is not None and r[I["offer"]].value: offers.append(str(r[I["offer"]].value).strip())
    for c in r:                                    # Excel cell comments also become offers
        if c.comment:
            t = re.sub(r"^\s*[\w .]+:\s*\n", "", c.comment.text).strip()
            if t and t not in offers: offers.append(t)
    valid = r[I["valid"]].value if I["valid"] is not None else None
    if hasattr(valid, "strftime"): valid = valid.strftime("%d %b %Y")
    cans = round(num(r[I["pack"]].value) / size) if size else 1
    bucket = str(r[I["unit"]].value).upper() == "BUC"
    out.append(dict(code=str(int(code)) if isinstance(code, (int, float)) else str(code), name=name,
        fam=PACK.sub("", name).strip(), seg=segment(r[I["seg"]].value, name), sz=size,
        pack=(f"{round(size*1000)} ml" if size < 1 else f"{size:g} L") + (" bucket" if bucket else ""),
        per=("per bucket" if bucket else f"carton of {cans}"), mrp=round(mrp, 2), dlp=round(dlp, 2),
        ex=round(ex, 2), net=round(ex * 1.18, 2), offers=offers, valid=str(valid) if valid else "",
        promo="PROMO" in name.upper()))
ist = dt.timezone(dt.timedelta(hours=5, minutes=30))
json.dump({"updated": dt.datetime.now(ist).isoformat(timespec="minutes"), "products": out},
          open("data/products.json", "w", encoding="utf-8"), separators=(",", ":"), ensure_ascii=False)
print(f"Wrote {len(out)} products from {files[-1]}")
