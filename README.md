# Guía de formatos Prisa Media

Herramienta de consulta para anunciantes: eligen emisora, tipo de formato, duración o # de palabras y franja de pauta, y ven las especificaciones que deben tener en cuenta. En digital consultan medidas, pesos y archivos. Si quieren segmentar, indican edad y nivel socioeconómico (mapa de afinidad) y la guía recomienda la emisora.

Hay dos versiones con el mismo contenido, los mismos nombres y los mismos colores:

| Archivo | Qué es |
|---|---|
| `index.html` | Plataforma web de un solo archivo (funciona sin internet). |
| `Guia_Formatos_PrisaMedia.xlsx` | Excel con listas desplegables en cascada y fórmulas (sin macros). |

## Cómo usarlas

**HTML:** abre `index.html` en el navegador, o publícalo con GitHub Pages (Settings → Pages → Deploy from a branch → raíz del repositorio).

**Excel:** abre el archivo y empieza por la hoja *Inicio*. Las celdas con borde negro son listas desplegables. Las hojas de consulta (*Audiencia, Radio, Digital, Mi plan*) están protegidas sin contraseña para no borrar fórmulas: *Revisar → Desproteger hoja*.

Ambas permiten armar un **plan** con varias selecciones y guardarlo en PDF.

## Cómo actualizar el contenido

Todo el contenido (formatos, franjas, piezas digitales, audiencias) está en `build/datos.py`. Después de editarlo:

```bash
pip install openpyxl pillow
python3 build/build.py
```

Esto regenera `data/formatos.json`, `index.html` y el Excel, así los dos quedan siempre iguales.

## Estructura

- `build/datos.py`: fuente editorial de los datos.
- `build/build_html.py`, `build/build_excel.py`: generadores.
- `src/index.html`: plantilla de la plataforma (CSS y JS).
- `assets/`: logos (extraídos de la presentación de audiencias) y tipografía.
- `fuente/`: el Excel original de especificaciones.

## Fuentes

Especificaciones de formatos Prisa Media (Grupo Nutresa) · Audiencias Prisa Media: mapa de afinidad TGI R2 2024 y perfiles TGI Choices 2026.
