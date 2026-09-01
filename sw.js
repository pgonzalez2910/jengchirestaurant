from pathlib import Path
src = Path("/mnt/data/jc_new_order_site/sw.js")
dst = Path("/mnt/data/sw.js")
dst.write_text(src.read_text(encoding="utf-8"), encoding="utf-8")
print("Ready:", dst)
