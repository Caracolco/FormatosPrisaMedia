#!/usr/bin/env python3
"""Regenera todo: data/formatos.json, index.html y Guia_Formatos_PrisaMedia.xlsx."""
import subprocess, sys, pathlib
b = pathlib.Path(__file__).parent
for s in ("datos.py", "build_html.py", "build_excel.py"):
    subprocess.run([sys.executable, str(b / s)], check=True)
