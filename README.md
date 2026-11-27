# Guía de formatos Prisa Media

Herramienta de consulta para anunciantes. Lo principal son los **formatos**: eligen emisora, tipo de formato, duración o # de palabras y franja de pauta, y ven las especificaciones. En digital consultan medidas, pesos y archivos. De forma **opcional**, si quieren segmentar, indican edad y nivel socioeconómico (mapa de afinidad) y la guía recomienda la emisora.

Hay dos versiones con el mismo contenido, los mismos nombres y los mismos colores:

| Archivo | Qué es |
|---|---|
| `index.html` | Plataforma web de un solo archivo (funciona sin internet). |
| `Guia_Formatos_PrisaMedia.xlsx` | Excel con listas desplegables en cascada y fórmulas (sin macros). |

Cada formato incluye además una **recomendación estratégica**: cuándo mencionar la marca (inicio, inicio y cierre, o cierre), cómo plantear el mensaje, el llamado a la acción y qué evitar. Son pautas generales de planificación de medios, no hacen parte del archivo de especificaciones; se editan en `build/datos.py` (`_estrategia_radio` y `_estrategia_digital`).

## Cómo usarlas

**HTML:** abre `index.html` en el navegador, o publícalo con GitHub Pages (Settings → Pages → Deploy from a branch → raíz del repositorio).

**Excel:** abre el archivo; hay cuatro hojas: *Radio*, *Digital*, *Mi plan* y *Recomendar* (opcional). Las celdas con borde negro son listas desplegables. Las hojas de consulta están protegidas sin contraseña para no borrar fórmulas (*Revisar → Desproteger hoja*). Las hojas de datos (*Datos_…*, *Listas*) están ocultas: clic derecho sobre una pestaña → *Mostrar*.

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
