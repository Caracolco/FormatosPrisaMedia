"""Genera index.html (un solo archivo) a partir de src/index.html + data/formatos.json."""
import base64, json, pathlib
RAIZ = pathlib.Path(__file__).resolve().parent.parent

def b64(p): return base64.b64encode(p.read_bytes()).decode()

def construir():
    data = json.loads((RAIZ / "data" / "formatos.json").read_text(encoding="utf-8"))
    logos = {p.stem: "data:image/png;base64," + b64(p) for p in sorted((RAIZ / "assets" / "logos").glob("*.png"))}
    font = ("@font-face{font-family:\"Schibsted Grotesk\";font-style:normal;font-weight:400 900;font-display:swap;"
            "src:url(data:font/woff2;base64," + b64(RAIZ / "assets/fonts/schibsted-grotesk-latin.woff2") + ") format(\"woff2\");"
            "unicode-range:U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD}")
    html = (RAIZ / "src" / "index.html").read_text(encoding="utf-8")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = (html.replace("/*__FONT__*/", font)
                .replace("/*__DATA__*/", payload)
                .replace("/*__LOGOS__*/", json.dumps(logos))
                .replace("__LOGO_PRISA__", logos["prisamedia"]))
    out = RAIZ / "index.html"
    out.write_text(html, encoding="utf-8")
    print("OK", out, f"{out.stat().st_size/1024:.0f} KB")

if __name__ == "__main__":
    construir()
