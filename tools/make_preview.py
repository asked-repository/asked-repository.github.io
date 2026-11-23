"""Copy the CSVs into data/preview-data.js so index.html also works when opened from disk."""
import json, os, sys
d = sys.argv[1] if len(sys.argv) > 1 else "data"
out = {n: open(os.path.join(d, n + ".csv"), encoding="utf-8-sig").read() for n in ["surveys", "concepts", "themes", "variables"]}
open(os.path.join(d, "preview-data.js"), "w", encoding="utf-8").write("window.PREVIEW_DATA=" + json.dumps(out, ensure_ascii=False) + ";\n")
