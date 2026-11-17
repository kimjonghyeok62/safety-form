# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from xlconnect import get_workbook

wb = get_workbook()
vbproj = wb.VBProject

os.makedirs("live_dump", exist_ok=True)
for comp in vbproj.VBComponents:
    code = comp.CodeModule
    n = code.CountOfLines
    text = code.Lines(1, n) if n > 0 else ""
    safe_name = comp.Name
    with open(f"live_dump/{safe_name}.txt", "w", encoding="utf-8") as f:
        f.write(text)
print("dumped", vbproj.VBComponents.Count, "modules")
