"""Genera Guia_Formatos_PrisaMedia.xlsx desde data/formatos.json.

Mismo contenido, mismos nombres y mismos colores que el HTML. Las hojas de consulta
usan listas desplegables en cascada y fórmulas (sin macros), de modo que funcionan en
Excel 2013+ y LibreOffice. Los datos viven en hojas Datos_* editables.
"""
import json
import pathlib
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

RAIZ = pathlib.Path(__file__).resolve().parent.parent
D = json.loads((RAIZ / "data" / "formatos.json").read_text(encoding="utf-8"))

# ---- paleta (idéntica al HTML) ----
INK, INK2, INK3 = "15181C", "444A53", "5E636C"
BG, SURF, SURF2, LINE = "F3F2ED", "FFFFFF", "FAF9F5", "D9D7CE"
ACC, ACC_SOFT = "07794A", "E2F1E9"
WARN_SOFT, WARN = "FBF0DA", "8A5200"
FONT = "Calibri"

def fill(c):
    c = c.lstrip("#")
    return PatternFill("solid", fgColor=c, bgColor=c)
def font(sz=11, b=False, c=INK, i=False): return Font(name=FONT, size=sz, bold=b, color=c, italic=i)
thin = Side(style="thin", color=LINE)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFTC = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

EMISORAS = D["emisoras"]
EM_NOMBRES = [e["nombre"] for e in EMISORAS]
FAM_LABEL = {"hablada": "Radio hablada", "musical": "Radio musical"}

wb = Workbook()

# ----------------------------------------------------------------- utilidades
def base_sheet(ws, tab, widths, zoom=100):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.sheet_properties.tabColor = tab
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    for row in ws.iter_rows(min_row=1, max_row=1):
        pass

def paint_bg(ws, rows, cols):
    f = fill(BG)
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(r, c).fill = f

def title(ws, text, sub, col_to):
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=col_to)
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=col_to)
    c = ws.cell(2, 2, text); c.font = font(22, True); c.alignment = Alignment(vertical="center")
    s = ws.cell(3, 2, sub); s.font = font(11, c=INK2); s.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[1].height = 34
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 34
    logo = XLImage(str(RAIZ / "assets/logos/prisamedia.png")); logo.width, logo.height = 96, 28
    ws.add_image(logo, "B1")

def label(ws, r, text, n=None, col=2):
    c = ws.cell(r, col, (f"{n}   " if n else "") + text)
    c.font = font(11, True); c.alignment = LEFTC; c.fill = fill(SURF); c.border = BOX

def inp(ws, r, c1, c2, value=None):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    for c in range(c1, c2 + 1):
        ws.cell(r, c).fill = fill("FFFFFF"); ws.cell(r, c).border = Border(left=Side(style="medium", color=INK) if c == c1 else None,
                                                                        right=Side(style="medium", color=INK) if c == c2 else None,
                                                                        top=Side(style="medium", color=INK), bottom=Side(style="medium", color=INK))
    cell = ws.cell(r, c1, value); cell.font = font(12, True); cell.alignment = LEFTC
    ws.row_dimensions[r].height = 28
    return cell

def band(ws, r, text, c1, c2, color=INK):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    c = ws.cell(r, c1, text); c.font = font(12, True, "FFFFFF"); c.fill = fill(color); c.alignment = LEFTC
    for k in range(c1, c2 + 1): ws.cell(r, k).fill = fill(color)
    ws.row_dimensions[r].height = 26

def out_row(ws, r, lab, formula, c1=3, c2=8, h=None, big=False, soft=None):
    label(ws, r, lab)
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    for k in range(c1, c2 + 1):
        ws.cell(r, k).fill = fill(soft or SURF); ws.cell(r, k).border = BOX
    c = ws.cell(r, c1, formula); c.font = font(14 if big else 11, big, INK); c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
    if h: ws.row_dimensions[r].height = h
    return c

def dv_list(ws, cell, formula, prompt):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True,
                        errorTitle="Opción no válida", error="Elige una opción de la lista desplegable.",
                        showInputMessage=True, promptTitle="Elige una opción", prompt=prompt)
    ws.add_data_validation(dv); dv.add(cell)

def hide_cols(ws, c1, c2):
    for c in range(c1, c2 + 1):
        ws.column_dimensions[L(c)].hidden = True

def height_for(texts, chars_per_line, min_h=30, cap=330):
    """Altura de fila para texto que cambia con la selección: percentil 80 de líneas (el usuario puede ampliar la fila)."""
    cnt = sorted(sum(max(1, -(-len(p) // chars_per_line)) for p in str(t).split("\n")) for t in texts if str(t).strip())
    if not cnt: return min_h
    n = cnt[min(len(cnt) - 1, int(len(cnt) * 0.8))]
    return min(cap, max(min_h, n * 15.5 + 8))

def em_color_rules(ws, rng, first_cell):
    for e in EMISORAS:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first_cell}="{e["nombre"]}"'],
                                                       fill=fill(e["color"]), font=Font(name=FONT, bold=True, color="FFFFFF")))

# ----------------------------------------------------------------- datos aplanados
def txt_comp(f): return "\n".join("• " + c for c in f["componentes"])
def txt_spec(f): return "\n".join(f"• {s['t']}: {s['d']}" for s in f["specs"])
def franja_texto(f): return f["franjaTexto"] or ""
def txt_estr(f): return "\n".join(f"• {x['t']}: {x['d']}" for x in f.get("estrategia", []))
FRJ = {fam: {x["id"]: x for x in lst} for fam, lst in D["franjas"].items()}

radio = sorted(D["radio"], key=lambda f: 0)  # ya viene ordenado por familia/tipo
# orden estable: familia (hablada primero), tipo en orden de aparición
fam_order = ["hablada", "musical"]
radio = sorted(D["radio"], key=lambda f: (fam_order.index(f["familia"]),))
tipos_orden = {}
for f in D["radio"]:
    tipos_orden.setdefault(f["familia"], [])
    if f["tipo"] not in tipos_orden[f["familia"]]: tipos_orden[f["familia"]].append(f["tipo"])
radio = sorted(D["radio"], key=lambda f: (fam_order.index(f["familia"]), tipos_orden[f["familia"]].index(f["tipo"])))

# ================================================================= HOJAS DE DATOS
# ---- Datos_Radio
wr = wb.active; wr.title = "Datos_Radio"
hdr = ["Clave tipo", "Clave formato", "Familia", "Tipo de formato", "Formato", "Duración", "# de palabras", "Palabras máx.",
       "Descripción", "Qué incluye / estructura", "Nota de franja", "Especificaciones", "Uso recomendado", "Formato protegido", "Franjas (n)", "Recomendación estratégica"]
wr.append(hdr)
for f in radio:
    pal = f["palabras"] if f["palabras"] else "N/A"
    wr.append([f"{f['familia']}|{f['tipo']}", f"{f['familia']}|{f['tipo']}|{f['nombre']}", f["familia"], f["tipo"], f["nombre"], f["duracion"],
               pal, f["maxPalabras"] or "", f["descripcion"] or "", txt_comp(f), franja_texto(f), txt_spec(f), f["consejo"] or "",
               "Sí" if f["protegido"] else "No", len(f["franjas"]), txt_estr(f)])
NR = len(radio) + 1

# ---- Datos_Franjas  (una fila por formato × franja)
wf = wb.create_sheet("Datos_Franjas")
wf.append(["Clave formato", "Franja de pauta", "Horario", "Clave completa"])
for f in radio:
    for fid in f["franjas"]:
        fr = FRJ[f["familia"]][fid]
        hor = (f.get("franjaHorarios") or {}).get(fid) or fr["horario"]
        wf.append([f"{f['familia']}|{f['nombre']}", fr["nombre"], hor, f"{f['familia']}|{f['nombre']}|{fr['nombre']}"])
NF = wf.max_row

# ---- Datos_Digital
dd = wb.create_sheet("Datos_Digital")
dd.append(["Clave", "Tipo de pieza", "Formato", "Pieza", "Tamaño o duración (inicial)", "Expandido", "Peso máximo", "Archivo", "Observaciones"])
for f in D["digital"]:
    for p in f["piezas"]:
        dd.append([f"{f['categoria']}|{f['nombre']}", f["categoria"], f["nombre"], p["pieza"], p["medida"], p["expande"] or "N/A", p["peso"] or "—", p["archivo"], p["obs"] or "—"])
ND = dd.max_row

# ---- Datos_Audiencia
da = wb.create_sheet("Datos_Audiencia")
bandas = D["bandasEdad"]
head = ["Id", "Emisora", "Formato", "Público (mapa de afinidad)"] + [f"Edad {b['label']}" for b in bandas] + [f"NSE {n}" for n in D["nse"]] + \
       ["Oyentes AM/FM", "Oyentes streaming", "Visitantes web + app", "Seguidores redes", "Reproducciones video", "Oyentes AM/FM (valor)"] + [p["label"] for p in D["perfiles"]]
da.append(head)
for e in EMISORAS:
    row = [e["id"], e["nombre"], e["formato"], e["rango"]] + [1 if i in e["bandas"] else 0 for i in range(5)] + [1 if n in e["nse"] else 0 for n in D["nse"]]
    row += [e["metricas"][k] for k in ("am", "stream", "web", "redes", "video")] + [e["valores"]["am"]]
    por = {p["id"]: p["pct"] for p in e["perfiles"]}
    row += [(por[p["id"]] / 100 if p["id"] in por else None) for p in D["perfiles"]]
    da.append(row)
C_AGE0, C_NSE0, C_AM, C_P0 = 5, 10, 15, 21   # columnas de inicio (1-based)
for r in range(2, 7):
    for c in range(C_P0, C_P0 + len(D["perfiles"])): da.cell(r, c).number_format = "0%"
# bloque de cálculo (filas 10-15)
da["A9"] = "CÁLCULO (no editar: lo usa la hoja Audiencia)"; da["A9"].font = font(11, True)
for j, h in enumerate(["Emisora", "Edades cubiertas", "Niveles NSE cubiertos", "Ajuste", "Hábito %", "Clave de orden", "Veredicto"], 1):
    da.cell(10, j, h)

# ---- Listas
ls = wb.create_sheet("Listas")
ls["A1"], ls["B1"] = "Emisora", "Familia"
for i, e in enumerate(EMISORAS, 2):
    ls.cell(i, 1, e["nombre"]); ls.cell(i, 2, e["familia"])
ls["D1"], ls["E1"] = "Familia", "Tipo de formato"
r = 2
for fam in fam_order:
    for t in tipos_orden[fam]:
        ls.cell(r, 4, fam); ls.cell(r, 5, t); r += 1
NT = r - 1
cats = []
for f in D["digital"]:
    if f["categoria"] not in cats: cats.append(f["categoria"])
ls["G1"] = "Tipo de pieza (digital)"
for i, c in enumerate(cats, 2): ls.cell(i, 7, c)
hd = ["Tipo de pieza", "Formato", "Clave", "Aplica para", "Texto del post", "Ten en cuenta", "Pasos", "Fanpages", "# piezas"]
for j, h in enumerate(hd): ls.cell(1, 9 + j, h)
dig_sorted = sorted(D["digital"], key=lambda f: cats.index(f["categoria"]))
for i, f in enumerate(dig_sorted, 2):
    ls.cell(i, 9, f["categoria"]); ls.cell(i, 10, f["nombre"]); ls.cell(i, 11, f"{f['categoria']}|{f['nombre']}")
    ls.cell(i, 12, f["aplica"] if f["emisoras"] else "Todas las emisoras")
    ls.cell(i, 13, f["texto"] or "")
    ls.cell(i, 14, "\n".join("• " + n for n in f["notas"]))
    ls.cell(i, 15, "\n".join(f"{k}. {n}" for k, n in enumerate(f["pasos"], 1)))
    ls.cell(i, 16, "\n".join(f"{l['t']}: {l['u']}" for l in f["enlaces"]))
    ls.cell(i, 17, len(f["piezas"]))
    ls.cell(i, 36, txt_estr(f))
NL = len(dig_sorted) + 1
ls["S1"] = "Hábito o interés"
ls.cell(1, 36, "Recomendación estratégica")
ls.cell(2, 19, "(ninguno)")
for i, p in enumerate(D["perfiles"], 3): ls.cell(i, 19, p["label"])
NP = len(D["perfiles"]) + 2
# opciones simples del recomendador: edad y nivel socioeconómico
EDAD_OPC = [("Todas las edades", [0, 1, 2, 3, 4]), ("12 a 17 (Alpha)", [0]), ("18 a 24 (Gen Z)", [1]), ("25 a 34 (Millennials)", [2]),
            ("35 a 44 (Gen X)", [3]), ("45 y más (Baby Boomers)", [4]), ("Jóvenes 12 a 24", [0, 1]), ("Adultos 25 a 44", [2, 3]),
            ("Adultos 35 y más", [3, 4]), ("Adultos 25 y más", [2, 3, 4])]
NSE_OPC = [("Todos los niveles", [2, 3, 4, 5, 6])] + [(f"NSE {n}", [n]) for n in D["nse"]] + \
          [("NSE 2 y 3 (medio-bajo)", [2, 3]), ("NSE 3 a 6", [3, 4, 5, 6]), ("NSE 4 a 6 (medio-alto)", [4, 5, 6])]
ls["U1"] = "Edad"
for j in range(5): ls.cell(1, 22 + j, f"Edad {j}")
for i, (lab, idx) in enumerate(EDAD_OPC, 2):
    ls.cell(i, 21, lab)
    for j in range(5): ls.cell(i, 22 + j, 1 if j in idx else 0)
NA = len(EDAD_OPC) + 1
ls["AB1"] = "Nivel socioeconómico"
for j, n in enumerate(D["nse"]): ls.cell(1, 29 + j, f"NSE {n}")
for i, (lab, idx) in enumerate(NSE_OPC, 2):
    ls.cell(i, 28, lab)
    for j, n in enumerate(D["nse"]): ls.cell(i, 29 + j, 1 if n in idx else 0)
NN = len(NSE_OPC) + 1
for sh in (wr, wf, dd, da, ls):
    sh.sheet_properties.tabColor = "9A9A92"
    for c in sh[1]:
        c.font = font(11, True, "FFFFFF"); c.fill = fill(INK); c.alignment = Alignment(wrap_text=True, vertical="center")
    sh.freeze_panes = "A2"
for sh, ws_ in ((wr, [24, 40, 10, 18, 34, 28, 11, 11, 50, 60, 50, 60, 40, 11, 9, 80]), (wf, [40, 30, 50, 60]), (dd, [34, 26, 30, 36, 26, 16, 18, 22, 36]),
                (da, [14, 16, 16, 28] + [9] * 10 + [12] * 6 + [12] * 30), (ls, [18, 10, 3, 10, 22, 3, 30, 3, 28, 36, 42, 26, 36, 50, 40, 40, 9, 3, 30, 3, 28] + [11] * 5 + [3, 28] + [11] * 5)):
    for i, w in enumerate(ws_, 1): sh.column_dimensions[L(i)].width = w
for row in wr.iter_rows(min_row=2):
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
for sh in (wr, ls):
    for row in sh.iter_rows(min_row=2):
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")

# ================================================================= RADIO (consulta)
rs = wb.create_sheet("Radio", 0)
base_sheet(rs, "D92027", [2, 34, 15, 15, 15, 15, 15, 15, 2, 2, 2])
paint_bg(rs, 40, 12)
title(rs, "Formatos de radio", "Elige en orden los 4 pasos (celdas con borde negro) y abajo aparecen la duración, las palabras, la franja y lo que debes entregar.", 8)
label(rs, 5, "Emisora", 1); label(rs, 6, "Tipo de formato", 2); label(rs, 7, "Formato y duración o # de palabras", 3); label(rs, 8, "Franja de pauta", 4)
c5 = inp(rs, 5, 3, 8, "Caracol Radio"); c6 = inp(rs, 6, 3, 8, "Leídos"); c7 = inp(rs, 7, 3, 8, 'Mención leída de 30"'); c8 = inp(rs, 8, 3, 8, "La Luciérnaga")
# ayudantes (columnas ocultas K,L)
rs["K5"] = '=IFERROR(INDEX(Listas!$B$2:$B$6,MATCH($C$5,Listas!$A$2:$A$6,0)),"")'
rs["K6"] = '=$K$5&"|"&$C$6'
rs["K7"] = '=$K$5&"|"&$C$6&"|"&$C$7'
rs["K8"] = '=$K$5&"|"&$C$7'
rs["K9"] = '=$K$8&"|"&$C$8'
rs["K10"] = f'=IFERROR(MATCH($K$7,Datos_Radio!$B$2:$B${NR},0),0)'
rs["K11"] = f'=IFERROR(MATCH($K$9,Datos_Franjas!$D$2:$D${NF},0),0)'
rs["K12"] = f'=COUNTIF(Datos_Franjas!$A$2:$A${NF},$K$8)'
hide_cols(rs, 11, 11)
dv_list(rs, "C5", "=Listas!$A$2:$A$6", "Emisora: define si es radio hablada o musical.")
dv_list(rs, "C6", f'=OFFSET(Listas!$E$2,MATCH($K$5,Listas!$D$2:$D${NT},0)-1,0,COUNTIF(Listas!$D$2:$D${NT},$K$5),1)', "Tipo de formato disponible para la emisora.")
dv_list(rs, "C7", f'=OFFSET(Datos_Radio!$E$2,MATCH($K$5&"|"&$C$6,Datos_Radio!$A$2:$A${NR},0)-1,0,COUNTIF(Datos_Radio!$A$2:$A${NR},$K$5&"|"&$C$6),1)', "Formato y duración o # de palabras.")
dv_list(rs, "C8", f'=OFFSET(Datos_Franjas!$B$2,MATCH($K$8,Datos_Franjas!$A$2:$A${NF},0)-1,0,COUNTIF(Datos_Franjas!$A$2:$A${NF},$K$8),1)', "Franja o programa en el que se emite.")
em_color_rules(rs, "C5:H5", "$C$5")
rs.row_dimensions[4].height = 8; rs.row_dimensions[9].height = 22
rs.cell(9, 2, "¿No sabes qué emisora elegir? Te recomendamos una según tu público  →").hyperlink = "#Recomendar!A1"
rs.cell(9, 2).font = Font(name=FONT, size=10, color=ACC, underline="single")
band(rs, 10, "FICHA TÉCNICA", 2, 8)
R = lambda col: f'INDEX(Datos_Radio!${col}$2:${col}${NR},$K$10)'
pend = '"Completa los pasos 1 a 3 para ver la ficha."'
rows = [
    ("Formato", f'=IF($K$10=0,{pend},{R("E")}&IF({R("N")}="Sí","   ·   FORMATO PROTEGIDO",""))', 34, True),
    ("Duración", f'=IF($K$10=0,"",{R("F")})', 24, False),
    ("# de palabras", f'=IF($K$10=0,"",IF(ISNUMBER({R("G")}),{R("G")}&" palabras",IF({R("H")}<>"","Hasta "&{R("H")}&" palabras","No aplica")))', 24, False),
    ("Franja de pauta", f'=IF($K$10=0,"",IF($K$11>0,$C$8&": "&INDEX(Datos_Franjas!$C$2:$C${NF},$K$11),IF($K$12=1,INDEX(Datos_Franjas!$B$2:$B${NF},MATCH($K$8,Datos_Franjas!$A$2:$A${NF},0))&": "&INDEX(Datos_Franjas!$C$2:$C${NF},MATCH($K$8,Datos_Franjas!$A$2:$A${NF},0)),"Elige una franja en el paso 4 ("&$K$12&" opciones)."))&IF({R("K")}="","",CHAR(10)&{R("K")}))', height_for([franja_texto(f) for f in radio], 90, 40), False),
    ("Qué incluye", f'=IF($K$10=0,"",IF({R("J")}="",{R("I")},{R("J")}))', height_for([txt_comp(f) or (f["descripcion"] or "") for f in radio], 90, 30), False),
    ("Especificaciones", f'=IF($K$10=0,"",IF({R("L")}="","—",{R("L")}))', height_for([txt_spec(f) for f in radio], 90, 30), False),
    ("Uso recomendado", f'=IF($K$10=0,"",IF({R("M")}="","—",{R("M")}))', 30, False),
    ("Recomendación estratégica", f'=IF($K$10=0,"",{R("P")})', height_for([txt_estr(f) for f in radio], 90, 60), False),
]
r0 = 11
for i, (lab, fm, h, big) in enumerate(rows):
    out_row(rs, r0 + i, lab, fm, h=h, big=big, soft=(ACC_SOFT if lab == "Recomendación estratégica" else None))
rr = r0 + len(rows) + 1
band(rs, rr, "VERIFICA TU TEXTO (menciones leídas y formatos con palabras)", 2, 8)
label(rs, rr + 1, "Pega aquí tu texto")
rs.merge_cells(start_row=rr + 1, start_column=3, end_row=rr + 1, end_column=8)
for k in range(3, 9): rs.cell(rr + 1, k).border = Border(top=Side(style="medium", color=INK), bottom=Side(style="medium", color=INK),
                                                       left=Side(style="medium", color=INK) if k == 3 else None, right=Side(style="medium", color=INK) if k == 8 else None)
rs.cell(rr + 1, 3).alignment = WRAP; rs.cell(rr + 1, 3).font = font(11); rs.row_dimensions[rr + 1].height = 86
rs.cell(rr + 1, 2).alignment = Alignment(vertical="top", indent=1, wrap_text=True)
txt = f"$C${rr+1}"
cnt = f'IF(LEN(TRIM({txt}))=0,0,LEN(TRIM({txt}))-LEN(SUBSTITUTE(TRIM({txt})," ",""))+1)'
mx = f'IFERROR(IF(INDEX(Datos_Radio!$H$2:$H${NR},$K$10)="",0,INDEX(Datos_Radio!$H$2:$H${NR},$K$10)),0)'
out_row(rs, rr + 2, "Resultado", f'=IF({mx}=0,"Este formato no tiene un límite de palabras.",{cnt}&" de "&{mx}&" palabras · "&IF({cnt}>{mx},"te pasas por "&({cnt}-{mx}),IF({cnt}={mx},"justo","te quedan "&({mx}-{cnt}))))', h=28, big=True)
rs.conditional_formatting.add(f"C{rr+2}:H{rr+2}", FormulaRule(formula=[f'AND({mx}>0,{cnt}>{mx})'], fill=fill("FBE7E4"), font=Font(name=FONT, bold=True, color="B3261E")))
rs.conditional_formatting.add(f"C{rr+2}:H{rr+2}", FormulaRule(formula=[f'AND({mx}>0,{cnt}<={mx},{cnt}>0)'], fill=fill(ACC_SOFT)))
rs.cell(rr + 4, 2, "Fuente: especificaciones de formatos Prisa Media. Guía de consulta; las condiciones comerciales las confirma tu ejecutivo.").font = font(9, c=INK3, i=True)
radio_cnt_row = rr + 1

# ================================================================= DIGITAL (consulta)
ds = wb.create_sheet("Digital", 1)
base_sheet(ds, "15181C", [2, 34, 22, 20, 16, 14, 18, 40, 2, 2])
paint_bg(ds, 40, 11)
title(ds, "Formatos digitales", "Medidas, pesos y archivos para web, mobile, audio y video en streaming, redes sociales y contenidos.", 8)
label(ds, 5, "Tipo de pieza", 1); label(ds, 6, "Formato", 2)
inp(ds, 5, 3, 8, "Brand Day"); inp(ds, 6, 3, 8, "Brand Day Caracol Radio")
ds["K5"] = '=$C$5&"|"&$C$6'
ds["K6"] = f'=IFERROR(MATCH($K$5,Listas!$K$2:$K${NL},0),0)'
ds["K7"] = f'=IFERROR(MATCH($K$5,Datos_Digital!$A$2:$A${ND},0),0)'
ds["K8"] = f'=IF($K$6=0,0,INDEX(Listas!$Q$2:$Q${NL},$K$6))'
hide_cols(ds, 11, 11)
dv_list(ds, "C5", f"=Listas!$G$2:$G${len(cats)+1}", "Tipo de pieza digital.")
dv_list(ds, "C6", f'=OFFSET(Listas!$J$2,MATCH($C$5,Listas!$I$2:$I${NL},0)-1,0,COUNTIF(Listas!$I$2:$I${NL},$C$5),1)', "Formato dentro del tipo de pieza.")
ds.row_dimensions[4].height = 8
band(ds, 8, "FICHA TÉCNICA", 2, 8)
LD = lambda col: f'INDEX(Listas!${col}$2:${col}${NL},$K$6)'
dpend = '"Completa los pasos 1 y 2 para ver la ficha."'
out_row(ds, 9, "Formato", f'=IF($K$6=0,{dpend},$C$6)', h=34, big=True)
out_row(ds, 10, "Aplica para", f'=IF($K$6=0,"",{LD("L")})', h=24)
out_row(ds, 11, "Texto del post", f'=IF($K$6=0,"",IF({LD("M")}="","—",{LD("M")}))', h=height_for([f["texto"] or "" for f in D["digital"]], 92, 24), soft=WARN_SOFT)
# tabla de piezas
hr = 13
for j, h in enumerate(["Pieza", "Tamaño o duración", "Expandido", "Peso máximo", "Archivo", "Observaciones"]):
    c = ds.cell(hr, 2 + j, h); c.font = font(11, True, "FFFFFF"); c.fill = fill(INK2); c.alignment = CEN; c.border = BOX
ds.row_dimensions[hr].height = 26
maxp = max(len(f["piezas"]) for f in D["digital"])
colmap = ["D", "E", "F", "G", "H", "I"]
for k in range(1, maxp + 1):
    r = hr + k
    for j, col in enumerate(colmap):
        f_ = f'=IF(AND($K$7>0,{k}<=$K$8),INDEX(Datos_Digital!${col}$2:${col}${ND},$K$7+{k-1}),"")'
        c = ds.cell(r, 2 + j, f_); c.font = font(11, j == 0); c.alignment = WRAP; c.border = BOX; c.fill = fill(SURF)
    ds.row_dimensions[r].height = 46 if k == 1 else 34
rn = hr + maxp + 2
out_row(ds, rn, "Ten en cuenta", f'=IF($K$6=0,"",IF({LD("N")}="","—",{LD("N")}))', h=height_for(["\n".join("• " + n for n in f["notas"]) for f in D["digital"]], 92, 24))
out_row(ds, rn + 1, "Pasos", f'=IF($K$6=0,"",IF({LD("O")}="","—",{LD("O")}))', h=height_for(["\n".join(f["pasos"]) for f in D["digital"]], 88, 24))
out_row(ds, rn + 3, "Recomendación estratégica", f'=IF($K$6=0,"",{LD("AJ")})', h=height_for([txt_estr(f) for f in D["digital"]], 90, 60), soft=ACC_SOFT)
out_row(ds, rn + 2, "Fanpages", f'=IF($K$6=0,"",IF({LD("P")}="","—",{LD("P")}))', h=height_for(["\n".join(l["u"] for l in f["enlaces"]) for f in D["digital"]], 92, 24))
for r in range(hr + 1, hr + maxp + 1):
    ds.conditional_formatting.add(f"B{r}:G{r}", FormulaRule(formula=[f'AND($K$7>0,{r-hr}<=$K$8)'], fill=fill(SURF)))
ds.cell(rn + 5, 2, "Fuente: especificaciones de formatos digitales Prisa Media. Las recomendaciones estratégicas son pautas generales de planificación (ajústalas con tu ejecutivo).").font = font(9, c=INK3, i=True)

# ================================================================= RECOMENDAR EMISORA (opcional)
au = wb.create_sheet("Recomendar", 0)
base_sheet(au, "9A9A92", [2, 32, 17, 17, 17, 17, 17, 17, 2, 2])
paint_bg(au, 40, 11)
title(au, "Recomendar emisora (opcional)", "Si quieres segmentar, elige a quién quieres llegar y te sugerimos la emisora. Si ya sabes dónde pautar, ve directo a la hoja Radio o Digital.", 8)
label(au, 5, "Edad", 1); label(au, 6, "Nivel socioeconómico", 2); label(au, 7, "Hábito o interés (opcional)", 3)
inp(au, 5, 3, 8, "Adultos 35 y más"); inp(au, 6, 3, 8, "NSE 4 a 6 (medio-alto)"); inp(au, 7, 3, 8, "(ninguno)")
dv_list(au, "C5", f"=Listas!$U$2:$U${NA}", "Para total personas elige «Todas las edades».")
dv_list(au, "C6", f"=Listas!$AB$2:$AB${NN}", "Para total personas elige «Todos los niveles».")
dv_list(au, "C7", f"=Listas!$S$2:$S${NP}", "Muestra el % de oyentes de cada emisora con ese hábito (TGI Choices 2026).")
au.row_dimensions[4].height = 8
for j in range(5):
    au[f"{L(11+j)}3"] = f'=IFERROR(INDEX(Listas!${L(22+j)}$2:${L(22+j)}${NA},MATCH($C$5,Listas!$U$2:$U${NA},0)),1)'
    au[f"{L(11+j)}4"] = f'=IFERROR(INDEX(Listas!${L(29+j)}$2:${L(29+j)}${NN},MATCH($C$6,Listas!$AB$2:$AB${NN},0)),1)'
au["Q1"] = "=SUM($K$3:$O$3)"; au["Q2"] = "=SUM($K$4:$O$4)"
au["Q3"] = f'=IFERROR(MATCH($C$7,Datos_Audiencia!${L(C_P0)}$1:${L(C_P0+len(D["perfiles"])-1)}$1,0),0)'
hide_cols(au, 11, 17)
for i, e in enumerate(EMISORAS):
    r = 11 + i; src = 2 + i
    da.cell(r, 1, f"=Datos_Audiencia!B{src}")
    da.cell(r, 2, f"=SUMPRODUCT(Recomendar!$K$3:$O$3,Datos_Audiencia!$E{src}:$I{src})")
    da.cell(r, 3, f"=SUMPRODUCT(Recomendar!$K$4:$O$4,Datos_Audiencia!$J{src}:$N{src})")
    da.cell(r, 4, f"=(B{r}/Recomendar!$Q$1)*(C{r}/Recomendar!$Q$2)"); da.cell(r, 4).number_format = "0%"
    da.cell(r, 5, f"=IF(Recomendar!$Q$3=0,0,N(INDEX(Datos_Audiencia!${L(C_P0)}{src}:${L(C_P0+len(D['perfiles'])-1)}{src},1,Recomendar!$Q$3)))"); da.cell(r, 5).number_format = "0%"
    da.cell(r, 6, f"=D{r}*1000+E{r}*100+Datos_Audiencia!T{src}/1000000000")
    da.cell(r, 7, f'=IF(D{r}>=0.999,"Cubre todo tu público",IF(D{r}>=0.5,"Cubre gran parte",IF(D{r}>0,"Cubre una parte","Fuera de su público")))')
band(au, 9, "EMISORAS RECOMENDADAS", 2, 8, ACC)
cntfit = 'COUNTIF(Datos_Audiencia!$D$11:$D$15,">0")'
nm = lambda k: f'INDEX(Datos_Audiencia!$A$11:$A$15,MATCH(LARGE(Datos_Audiencia!$F$11:$F$15,{k}),Datos_Audiencia!$F$11:$F$15,0))'
rec = (f'=IF({cntfit}=0,"Ninguna emisora cubre esta combinación: prueba con otra edad o nivel.",'
       f'IF(AND($Q$1=5,$Q$2=5),"Para llegar a total personas, recomendamos todas las emisoras: cada una cubre una parte distinta del público.",'
       f'"Te recomendamos "&{nm(1)}&IF({cntfit}>=2,IF({cntfit}>=3,", "&{nm(2)}&" y "&{nm(3)}," y "&{nm(2)}),"")&"."))')
out_row(au, 10, "Recomendación", rec, h=44, big=True, soft=ACC_SOFT)
for j, h in enumerate(["Emisora", "Ajuste", "Edades que cubre", "Niveles NSE que cubre", "Hábito elegido", "Oyentes AM/FM", "Tipo de radio"]):
    c = au.cell(12, 2 + j, h); c.font = font(11, True, "FFFFFF"); c.fill = fill(INK2); c.alignment = CEN; c.border = BOX
au.row_dimensions[12].height = 28
for k in range(1, 6):
    r = 12 + k
    m = f'MATCH(LARGE(Datos_Audiencia!$F$11:$F$15,{k}),Datos_Audiencia!$F$11:$F$15,0)'
    vals = [f'=INDEX(Datos_Audiencia!$A$11:$A$15,{m})', f'=INDEX(Datos_Audiencia!$G$11:$G$15,{m})',
            f'=INDEX(Datos_Audiencia!$B$11:$B$15,{m})&" de "&$Q$1', f'=INDEX(Datos_Audiencia!$C$11:$C$15,{m})&" de "&$Q$2',
            f'=IF($Q$3=0,"—",IF(INDEX(Datos_Audiencia!$E$11:$E$15,{m})=0,"sin dato",TEXT(INDEX(Datos_Audiencia!$E$11:$E$15,{m}),"0%")))',
            f'=INDEX(Datos_Audiencia!$O$2:$O$6,{m})', f'=INDEX(Datos_Audiencia!$C$2:$C$6,{m})']
    for j, v in enumerate(vals):
        c = au.cell(r, 2 + j, v); c.font = font(12 if j == 0 else 11, j == 0); c.alignment = CEN if j else LEFTC; c.border = BOX; c.fill = fill(SURF)
    au.row_dimensions[r].height = 30
em_color_rules(au, "B13:B17", "$B13")
au.conditional_formatting.add("C13:C17", FormulaRule(formula=['C13="Fuera de su público"'], font=Font(name=FONT, color="8A8F97")))
au.conditional_formatting.add("C13:C17", FormulaRule(formula=['C13="Cubre todo tu público"'], fill=fill(ACC_SOFT), font=Font(name=FONT, bold=True, color="055B38")))
au.cell(18, 2, "Con la emisora elegida, ve a la hoja Radio o Digital y selecciónala en el paso 1.").font = font(10, c=INK3, i=True)
band(au, 20, "MAPA DE AFINIDAD  ·  edad (abajo) × nivel socioeconómico (izquierda)", 2, 8, INK)
au.cell(21, 2, "Cada casilla lista las emisoras cuyo público la cubre. La zona verde es el público que elegiste.").font = font(10, c=INK2, i=True)
for i, n in enumerate([6, 5, 4, 3, 2]):
    r = 22 + i
    c = au.cell(r, 2, f"NSE {n}"); c.font = font(12, True); c.alignment = Alignment(horizontal="right", vertical="center", indent=1); c.fill = fill(SURF); c.border = BOX
    au.row_dimensions[r].height = 62
    for j in range(5):
        parts = []
        for s_, e in enumerate(EMISORAS):
            src = 2 + s_
            parts.append(f'IF(AND(Datos_Audiencia!${L(C_AGE0+j)}${src}=1,Datos_Audiencia!${L(C_NSE0+D["nse"].index(n))}${src}=1),"{e["nombre"]}"&CHAR(10),"")')
        c = au.cell(r, 3 + j, "=" + "&".join(parts)); c.font = font(10, True, INK); c.alignment = CEN
        c.border = Border(left=Side(style="thin", color="FFFFFF"), right=Side(style="thin", color="FFFFFF"), top=Side(style="thin", color="FFFFFF"), bottom=Side(style="thin", color="FFFFFF"))
        c.fill = fill("E4E2DA")
        au.conditional_formatting.add(f"{L(3+j)}{r}", FormulaRule(formula=[f"AND(${L(11+j)}$3=1,INDEX($K$4:$O$4,{D['nse'].index(n)+1})=1)"], fill=fill("BFE0CF"), font=Font(name=FONT, bold=True, color="07482C")))
for j, b in enumerate(bandas):
    c = au.cell(27, 3 + j, f"{b['label']}\n{b['gen']}"); c.font = font(10, True); c.alignment = CEN; c.fill = fill(SURF); c.border = BOX
au.row_dimensions[27].height = 34
au.cell(29, 2, "Fuente: mapa de afinidad TGI R2 2024 · perfiles de oyentes TGI Choices 2026.").font = font(9, c=INK3, i=True)

# ================================================================= MI PLAN
mp = wb.create_sheet("Mi plan", 3)
base_sheet(mp, "6B6F76", [2, 5, 20, 20, 34, 24, 34, 70, 2, 12, 12, 12, 12])
paint_bg(mp, 40, 14)
title(mp, "Mi plan de pauta", "Arma tu selección fila por fila: elige emisora, tipo, formato y franja, y las especificaciones aparecen a la derecha. Las filas de ejemplo se pueden cambiar.", 8)
band(mp, 5, "RADIO", 2, 8, "D92027")
for j, h in enumerate(["#", "Emisora", "Tipo de formato", "Formato", "Franja de pauta", "Duración, palabras y horario", "Qué incluye y especificaciones"]):
    c = mp.cell(6, 2 + j, h); c.font = font(11, True, "FFFFFF"); c.fill = fill(INK2); c.alignment = CEN; c.border = BOX
mp.row_dimensions[6].height = 28
NRAD = 8
ejemplos = [("Caracol Radio", "Leídos", 'Mención leída de 30"', "La Luciérnaga"), ("Tropicana", "Remotos", "Remoto de 2 horas", "Programación musical")]
h_row = height_for([txt_comp(f) + "\n" + txt_spec(f) + "\n" + txt_estr(f) for f in radio], 66, 40, 260)
INBOX = Border(left=Side(style="medium", color=INK), right=Side(style="medium", color=INK), top=Side(style="medium", color=INK), bottom=Side(style="medium", color=INK))
for k in range(NRAD):
    r = 7 + k
    c = mp.cell(r, 2, k + 1); c.font = font(11, True, INK3); c.alignment = CEN; c.border = BOX; c.fill = fill(SURF)
    ex = ejemplos[k] if k < len(ejemplos) else ("", "", "", "")
    for j in range(4):
        c = mp.cell(r, 3 + j, ex[j] or None); c.font = font(11, True); c.alignment = WRAP; c.fill = fill("FFFFFF"); c.border = INBOX
    mp[f"J{r}"] = f'=IFERROR(INDEX(Listas!$B$2:$B$6,MATCH($C{r},Listas!$A$2:$A$6,0)),"")'
    mp[f"K{r}"] = f'=IFERROR(MATCH($J{r}&"|"&$D{r}&"|"&$E{r},Datos_Radio!$B$2:$B${NR},0),0)'
    mp[f"L{r}"] = f'=IFERROR(MATCH($J{r}&"|"&$E{r}&"|"&$F{r},Datos_Franjas!$D$2:$D${NF},0),0)'
    mp[f"M{r}"] = f'=$J{r}&"|"&$E{r}'
    dv_list(mp, f"C{r}", "=Listas!$A$2:$A$6", "Emisora")
    dv_list(mp, f"D{r}", f'=OFFSET(Listas!$E$2,MATCH($J{r},Listas!$D$2:$D${NT},0)-1,0,COUNTIF(Listas!$D$2:$D${NT},$J{r}),1)', "Tipo de formato")
    dv_list(mp, f"E{r}", f'=OFFSET(Datos_Radio!$E$2,MATCH($J{r}&"|"&$D{r},Datos_Radio!$A$2:$A${NR},0)-1,0,COUNTIF(Datos_Radio!$A$2:$A${NR},$J{r}&"|"&$D{r}),1)', "Formato")
    dv_list(mp, f"F{r}", f'=OFFSET(Datos_Franjas!$B$2,MATCH($M{r},Datos_Franjas!$A$2:$A${NF},0)-1,0,COUNTIF(Datos_Franjas!$A$2:$A${NF},$M{r}),1)', "Franja de pauta")
    I_ = lambda col: f'INDEX(Datos_Radio!${col}$2:${col}${NR},$K{r})'
    resumen = f'=IF($K{r}=0,"",{I_("F")}&IF(ISNUMBER({I_("G")})," · "&{I_("G")}&" palabras","")&IF($L{r}=0,"",CHAR(10)&INDEX(Datos_Franjas!$C$2:$C${NF},$L{r})))'
    detalle = f'=IF($K{r}=0,"",IF({I_("J")}="",{I_("I")},{I_("J")})&IF({I_("L")}="","",CHAR(10)&{I_("L")})&IF({I_("P")}="","",CHAR(10)&"RECOMENDACIÓN ESTRATÉGICA"&CHAR(10)&{I_("P")}))'
    for j, f_ in enumerate([resumen, detalle]):
        c = mp.cell(r, 7 + j, f_); c.font = font(10 if j else 11); c.alignment = WRAP; c.border = BOX; c.fill = fill(SURF)
    mp.row_dimensions[r].height = h_row if k < 2 else 44
em_color_rules(mp, f"C7:C{6+NRAD}", "$C7")
hide_cols(mp, 10, 13)
rd = 7 + NRAD + 1
band(mp, rd, "DIGITAL", 2, 8, INK)
for col, h in ((2, "#"), (3, "Tipo de pieza"), (5, "Formato"), (6, "Aplica para"), (8, "Medidas, pesos y archivos")):
    mp.cell(rd + 1, col, h)
for col in range(2, 9):
    c = mp.cell(rd + 1, col); c.font = font(11, True, "FFFFFF"); c.fill = fill(INK2); c.alignment = CEN; c.border = BOX
mp.merge_cells(start_row=rd + 1, start_column=3, end_row=rd + 1, end_column=4)
mp.merge_cells(start_row=rd + 1, start_column=6, end_row=rd + 1, end_column=7)
mp.row_dimensions[rd + 1].height = 28
NDIG = 6
piezas_txt = {}
for f in D["digital"]:
    lines = [f"• {p['pieza']}: {p['medida']}" + (f" → {p['expande']}" if p["expande"] else "") + (f" · {p['peso']}" if p["peso"] else "") + f" · {p['archivo']}" + (f" · {p['obs']}" if p["obs"] else "") for p in f["piezas"]]
    if f["texto"]: lines.insert(0, f"Texto: {f['texto']}")
    lines += ["• " + n for n in f["notas"]] + [f"{k}. {n}" for k, n in enumerate(f["pasos"], 1)]
    if f.get("estrategia"): lines += ["RECOMENDACIÓN ESTRATÉGICA", txt_estr(f)]
    piezas_txt[(f["categoria"], f["nombre"])] = "\n".join(lines)
ls.cell(1, 18, "Resumen de piezas")
for i, f in enumerate(dig_sorted, 2): ls.cell(i, 18, piezas_txt[(f["categoria"], f["nombre"])]).alignment = Alignment(wrap_text=True, vertical="top")
ls.column_dimensions["R"].width = 60
dej = [("Brand Day", "Brand Day Caracol Radio")]
hd_row = height_for(piezas_txt.values(), 66, 44, 260)
for k in range(NDIG):
    r = rd + 2 + k
    c = mp.cell(r, 2, k + 1); c.font = font(11, True, INK3); c.alignment = CEN; c.border = BOX; c.fill = fill(SURF)
    ex = dej[k] if k < len(dej) else ("", "")
    mp.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    for cc, v in ((3, ex[0]), (5, ex[1])):
        c = mp.cell(r, cc, v or None); c.font = font(11, True); c.alignment = WRAP
    for cc in (3, 4, 5):
        mp.cell(r, cc).fill = fill("FFFFFF"); mp.cell(r, cc).border = INBOX
    mp[f"J{r}"] = f'=IFERROR(MATCH($C{r}&"|"&$E{r},Listas!$K$2:$K${NL},0),0)'
    dv_list(mp, f"C{r}", f"=Listas!$G$2:$G${len(cats)+1}", "Tipo de pieza")
    dv_list(mp, f"E{r}", f'=OFFSET(Listas!$J$2,MATCH($C{r},Listas!$I$2:$I${NL},0)-1,0,COUNTIF(Listas!$I$2:$I${NL},$C{r}),1)', "Formato")
    mp.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    c = mp.cell(r, 6, f'=IF($J{r}=0,"",INDEX(Listas!$L$2:$L${NL},$J{r}))'); c.font = font(11); c.alignment = WRAP
    for cc in (6, 7): mp.cell(r, cc).border = BOX; mp.cell(r, cc).fill = fill(SURF)
    c = mp.cell(r, 8, f'=IF($J{r}=0,"",INDEX(Listas!$R$2:$R${NL},$J{r}))'); c.font = font(10); c.alignment = WRAP; c.border = BOX; c.fill = fill(SURF)
    mp.row_dimensions[r].height = hd_row if k < 1 else 44
mp.cell(rd + NDIG + 3, 2, "Para guardar o enviar tu plan: Archivo > Imprimir > Guardar como PDF. La hoja ya está configurada en horizontal.").font = font(10, c=INK3, i=True)
mp.page_setup.orientation = "landscape"; mp.page_setup.fitToWidth = 1; mp.page_setup.fitToHeight = 0
mp.sheet_properties.pageSetUpPr.fitToPage = True
for sh in (rs, ds, au):
    sh.page_setup.orientation = "portrait"; sh.page_setup.fitToWidth = 1; sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True

# orden: formatos primero; recomendación opcional al final; hojas de datos ocultas
order = ["Radio", "Digital", "Mi plan", "Recomendar", "Datos_Radio", "Datos_Franjas", "Datos_Digital", "Datos_Audiencia", "Listas"]
wb._sheets = [wb[n] for n in order]
for n in order[4:]: wb[n].sheet_state = "hidden"
wb.active = 0
for ws in wb: ws.sheet_view.tabSelected = (ws.title == "Radio")
def unlock(ws, coords):
    for co in coords: ws[co].protection = Protection(locked=False)
unlock(rs, ["C5", "C6", "C7", "C8", f"C{radio_cnt_row}"])
unlock(ds, ["C5", "C6"])
unlock(au, ["C5", "C6", "C7"])
for k in range(NRAD): unlock(mp, [f"{L(3+j)}{7+k}" for j in range(4)])
for k in range(NDIG): unlock(mp, [f"C{rd+2+k}", f"E{rd+2+k}"])
for ws in (rs, ds, au, mp):
    ws.protection.sheet = True
    ws.protection.formatRows = False; ws.protection.formatColumns = False
    ws.protection.selectLockedCells = False; ws.protection.selectUnlockedCells = False

out = RAIZ / "Guia_Formatos_PrisaMedia.xlsx"
wb.save(out)
print("OK", out, f"{out.stat().st_size/1024:.0f} KB")
