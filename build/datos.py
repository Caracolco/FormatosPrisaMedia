#!/usr/bin/env python3
"""Fuente editorial de la guía de formatos Prisa Media.

Todo el contenido (radio, digital, audiencias) vive aquí y se exporta a
data/formatos.json. Tanto el HTML como el Excel se generan desde ese JSON,
de modo que ambos siempre muestran exactamente lo mismo.

Origen: FORMATOS_PRISA_MEDIA - GRUPO NUTRESA (xlsx) y AUDIENCIAS_PRISA_MEDIA (pptx).
Editar aquí y ejecutar:  python3 build/build.py
"""
import json
import pathlib

RAIZ = pathlib.Path(__file__).resolve().parent.parent

# --------------------------------------------------------------------------
# AUDIENCIAS
# --------------------------------------------------------------------------
# Generaciones / rangos de edad del mapa de afinidad (diapositiva 1 del PPT)
BANDAS_EDAD = [
    {"id": "alpha", "label": "12 a 17", "gen": "Alpha", "min": 12, "max": 17},
    {"id": "genz", "label": "18 a 24", "gen": "Gen Z", "min": 18, "max": 24},
    {"id": "mill", "label": "25 a 34", "gen": "Millennials", "min": 25, "max": 34},
    {"id": "genx", "label": "35 a 44", "gen": "Gen X", "min": 35, "max": 44},
    {"id": "boom", "label": "45 y más", "gen": "Baby Boomers", "min": 45, "max": 99},
]
NSE = [2, 3, 4, 5, 6]

# Atajos de audiencia: (id, nombre, descripción, índices de bandas, niveles NSE)
ATAJOS = [
    {"id": "total", "label": "Total personas", "ayuda": "Todas las edades y niveles",
     "edad": [0, 1, 2, 3, 4], "nse": [2, 3, 4, 5, 6]},
    {"id": "jovenes", "label": "Jóvenes 12 a 24", "ayuda": "Alpha y Gen Z, todos los NSE",
     "edad": [0, 1], "nse": [2, 3, 4, 5, 6]},
    {"id": "adultos-jov", "label": "Adultos 25 a 44", "ayuda": "Millennials y Gen X",
     "edad": [2, 3], "nse": [2, 3, 4, 5, 6]},
    {"id": "adultos35", "label": "Adultos 35+", "ayuda": "Gen X y Baby Boomers",
     "edad": [3, 4], "nse": [2, 3, 4, 5, 6]},
    {"id": "adultos-alto", "label": "Adultos en estratos medio-alto", "ayuda": "25 años en adelante, NSE 4 a 6",
     "edad": [2, 3, 4], "nse": [4, 5, 6]},
    {"id": "medio-bajo", "label": "Estratos medio-bajo", "ayuda": "Todas las edades, NSE 2 y 3",
     "edad": [0, 1, 2, 3, 4], "nse": [2, 3]},
]

# Perfiles / hábitos (barras "% concentraciones de oyentes", TGI Choices 2026).
# id canónico para poder comparar emisoras; el rótulo original se conserva por emisora.
PERFILES = [
    # (id, rótulo, grupo)
    ("trabajadores", "Trabajadores", "Trabajo y estudio"),
    ("profesionales", "Profesionales", "Trabajo y estudio"),
    ("directivos", "Directivos", "Trabajo y estudio"),
    ("emprendedores", "Emprendedores", "Trabajo y estudio"),
    ("tenderos", "Tenderos / ventas", "Trabajo y estudio"),
    ("estudiantes", "Estudiantes", "Trabajo y estudio"),
    ("pensionados", "Pensionados", "Trabajo y estudio"),
    ("familia", "Padres y jefes de familia", "Hogar y familia"),
    ("dueno-casa", "Dueños de casa", "Hogar y familia"),
    ("mascotas", "Dueños de mascotas", "Hogar y familia"),
    ("mujeres", "Mujeres", "Hogar y familia"),
    ("decisor", "Decisores de compra", "Consumo"),
    ("compras", "Compradores / ir de compras", "Consumo"),
    ("gastronomia", "Gastronomía", "Consumo"),
    ("bienestar", "Bienestar", "Consumo"),
    ("conductores", "Conductores y dueños de vehículo", "Movilidad"),
    ("moteros", "Moteros", "Movilidad"),
    ("viajeros", "Viajeros", "Movilidad"),
    ("deporte", "Fans del deporte", "Ocio y entretenimiento"),
    ("deportistas", "Deportistas", "Ocio y entretenimiento"),
    ("music", "Amantes de la música", "Ocio y entretenimiento"),
    ("conciertos", "Conciertos", "Ocio y entretenimiento"),
    ("farandula", "Farándula", "Ocio y entretenimiento"),
    ("entretenimiento", "Entretenimiento", "Ocio y entretenimiento"),
    ("apostadores", "Apostadores", "Ocio y entretenimiento"),
    ("conectados", "Conectados", "Digital y generaciones"),
    ("gamers", "Gamers", "Digital y generaciones"),
    ("geek", "Geek", "Digital y generaciones"),
    ("genz", "Generación Z", "Digital y generaciones"),
    ("millennials", "Millennials", "Digital y generaciones"),
]

EMISORAS = [
    {
        "id": "caracol", "nombre": "Caracol Radio", "familia": "hablada", "formato": "Radio hablada",
        "color": "#D92027", "tinta": "#FFFFFF", "logo": "caracol", "logoBlanco": "caracol-w",
        "bandas": [3, 4], "nse": [2, 3, 4, 5, 6],
        "rango": "+35 años · NSE 2 al 6",
        "slogan": "Juntos, más compañía.",
        "metricas": {"am": "2,3 M", "stream": "1,7 M", "web": "853 K", "redes": "9,8 M", "video": "202 M"},
        "valores": {"am": 2300000, "stream": 1700000, "web": 853000, "redes": 9800000, "video": 202000000},
        "perfiles": [
            ("mascotas", "Dueños de mascotas", 84), ("deporte", "Fans del deporte", 83),
            ("trabajadores", "Trabajadores", 74), ("familia", "Jefe de familia", 64),
            ("dueno-casa", "Dueño de casa", 55), ("profesionales", "Profesionales", 39),
            ("compras", "Compras", 30), ("deportistas", "Deportistas", 26),
            ("directivos", "Directivos", 18), ("pensionados", "Pensionados", 5),
        ],
    },
    {
        "id": "tropicana", "nombre": "Tropicana", "familia": "musical", "formato": "Radio musical",
        "color": "#0B6FC2", "tinta": "#FFFFFF", "logo": "tropicana", "logoBlanco": "tropicana-w",
        "bandas": [0, 1, 2, 3], "nse": [2, 3],
        "rango": "12 a 44 años · NSE 2 a 3",
        "slogan": "",
        "metricas": {"am": "2,3 M", "stream": "1,7 M", "web": "853 K", "redes": "9,8 M", "video": "83 M"},
        "valores": {"am": 2300000, "stream": 1700000, "web": 853000, "redes": 9800000, "video": 83000000},
        "perfiles": [
            ("trabajadores", "Trabajadores", 85), ("mascotas", "Dueños de mascotas", 71),
            ("decisor", "Decisor de compra", 68), ("familia", "Padres de familia", 61),
            ("profesionales", "Profesionales", 46), ("conductores", "Conductores", 43),
            ("moteros", "Moteros", 28), ("emprendedores", "Emprendedor", 26),
            ("genz", "Generación Z", 23), ("tenderos", "Tenderos / ventas", 18),
        ],
    },
    {
        "id": "besame", "nombre": "Bésame", "familia": "musical", "formato": "Radio musical",
        "color": "#C8274F", "tinta": "#FFFFFF", "logo": "besame", "logoBlanco": "besame-w",
        "bandas": [3, 4], "nse": [2, 3, 4],
        "rango": "+35 años · NSE 2 al 4",
        "slogan": "Te enamora",
        "metricas": {"am": "2,4 M", "stream": "1,8 M", "web": "239 K", "redes": "2,8 M", "video": "43 M"},
        "valores": {"am": 2400000, "stream": 1800000, "web": 239000, "redes": 2800000, "video": 43000000},
        "perfiles": [
            ("music", "Music lovers", 77), ("trabajadores", "Trabajadores", 74),
            ("mujeres", "Mujeres", 61), ("farandula", "Farándula", 47),
            ("familia", "Padres", 45), ("bienestar", "Bienestar", 39),
            ("compras", "Ir de compras", 32), ("gastronomia", "Gastronomía", 32),
            ("emprendedores", "Emprendedor", 21),
        ],
    },
    {
        "id": "radioacktiva", "nombre": "Radioacktiva", "familia": "musical", "formato": "Radio musical",
        "color": "#5E3FB0", "tinta": "#FFFFFF", "logo": "radioacktiva", "logoBlanco": "radioacktiva-w",
        "bandas": [2, 3], "nse": [3, 4, 5, 6],
        "rango": "25 a 44 años · NSE 3 al 6",
        "slogan": "",
        "metricas": {"am": "304 K", "stream": "483 K", "web": "334 K", "redes": "3,5 M", "video": "29 M"},
        "valores": {"am": 304000, "stream": 483000, "web": 334000, "redes": 3500000, "video": 29000000},
        "perfiles": [
            ("music", "Music lovers", 94), ("trabajadores", "Trabajadores", 81),
            ("deporte", "Fans del deporte", 71), ("conductores", "Dueños de vehículo", 58),
            ("millennials", "Millennials", 57), ("viajeros", "Viajeros", 47),
            ("apostadores", "Apostadores", 23), ("conciertos", "Conciertos", 11),
        ],
    },
    {
        "id": "los40", "nombre": "Los40", "familia": "musical", "formato": "Radio musical",
        "color": "#111111", "tinta": "#FFFFFF", "logo": "los40", "logoBlanco": "los40",
        "bandas": [0, 1, 2], "nse": [3, 4, 5, 6],
        "rango": "12 a 34 años · NSE 3 a 6",
        "slogan": "",
        "metricas": {"am": "409 K", "stream": "315 K", "web": "374 K", "redes": "8,1 M", "video": "160 M"},
        "valores": {"am": 409000, "stream": 315000, "web": 374000, "redes": 8100000, "video": 160000000},
        "perfiles": [
            ("conectados", "Conectados", 98), ("entretenimiento", "Entretenimiento", 65),
            ("viajeros", "Viajeros", 59), ("estudiantes", "Estudiantes", 44),
            ("geek", "Geek", 43), ("gamers", "Gamers", 42),
            ("compras", "Compras", 32), ("genz", "Genz", 24), ("conciertos", "Conciertos", 23),
        ],
    },
]

METRICAS_ROTULOS = [
    ("am", "Oyentes", "AM/FM"),
    ("stream", "Oyentes", "Streaming"),
    ("web", "Visitantes", "Web + App"),
    ("redes", "Seguidores", "Redes sociales"),
    ("video", "Reproducciones", "de video"),
]

# --------------------------------------------------------------------------
# RADIO
# --------------------------------------------------------------------------
FRANJAS = {
    "musical": [
        {"id": "morning", "nombre": "Morning", "horario": "Lunes a viernes, de 6:00 a. m. a 10:00 a. m."},
        {"id": "musical", "nombre": "Programación musical", "horario": "Lunes a domingo, de 10:00 a. m. a 8:00 p. m."},
    ],
    "hablada": [
        {"id": "6am-w", "nombre": "6AM – W", "horario": "Lunes a viernes, de 5:00 a. m. a 12:00 m."},
        {"id": "noticiero", "nombre": "Noticiero ½ día", "horario": "Lunes a viernes, de 12:00 m. a 1:00 p. m."},
        {"id": "pulso", "nombre": "El Pulso del Fútbol", "horario": "Lunes a viernes, de 1:00 a 2:00 p. m."},
        {"id": "dos-puntos", "nombre": "Dos Puntos", "horario": "Lunes a viernes, de 2:00 a 4:00 p. m."},
        {"id": "luciernaga", "nombre": "La Luciérnaga", "horario": "Lunes a viernes, de 4:00 a 7:00 p. m."},
        {"id": "hora20", "nombre": "Hora 20", "horario": "Lunes a viernes, de 7:00 a 8:00 p. m."},
        {"id": "polemica", "nombre": "La Polémica", "horario": "Lunes a viernes, de 8:00 a 9:00 p. m."},
        {"id": "as-mundial", "nombre": "AS Mundial", "horario": "Lunes a viernes, de 9:00 a 10:00 p. m."},
        {"id": "de-pelicula", "nombre": "De Película", "horario": "Fin de semana, de 4:00 a 5:00 a. m."},
        {"id": "al-campo", "nombre": "Al Campo", "horario": "Fin de semana, de 5:00 a 6:00 p. m."},
        {"id": "planeta-caracol", "nombre": "Planeta Caracol", "horario": "Fin de semana, de 6:00 a 7:00 p. m."},
        {"id": "noticieros-fds", "nombre": "Noticieros fin de semana", "horario": "Sábados y domingos"},
        {"id": "a-vivir", "nombre": "A Vivir Que Son Dos Días", "horario": "Sábados y domingos, de 8:00 a 11:00 a. m."},
        {"id": "noche-fantastica", "nombre": "Noche Fantástica", "horario": "Viernes y sábados, de 10:00 p. m. a 4:00 a. m."},
        {"id": "caracol-fds", "nombre": "Caracol Fin de Semana", "horario": "Sábado o domingo (no aplica en la franja de 12:00 m. a 1:00 p. m.)"},
    ],
}
PROGRAMAS_HABLADA = [f["id"] for f in FRANJAS["hablada"]]

_dur_cuna = ['5"', '10"', '15"', '20"', '30"', '35"', '40"', '45"', '60"']
_leidas = [('5"', 17), ('10"', 32), ('15"', 46), ('20"', 64), ('30"', 95), ('40"', 115), ('60"', 180)]


def _consejo_cuna(d):
    if d in ('5"', '10"'):
        return "Duración ideal para reforzar la recordación de marca, jingles y mensajes de posicionamiento."
    if d == '30"':
        return "Las cuñas de 30 segundos permiten desarrollar mensajes con mayor dinamismo y contenido."
    return None


def _consejo_leida(d):
    if d in ('20"', '30"'):
        return "Las menciones de 20 y 30 segundos permiten desarrollar mensajes con mayor dinamismo y contenido."
    return None


UBICACION_BREAK = "Se emiten dentro del break comercial, compartiendo los espacios comerciales de cada hora con otras marcas."


def _radio():
    f = []

    def add(**k):
        base = {"palabras": None, "maxPalabras": None, "componentes": [], "specs": [], "consejo": None,
                "protegido": False, "franjaTexto": None, "descripcion": None, "duracion": None}
        base.update(k)
        base["id"] = f"{base['familia']}-{base['tipo_id']}-{len(f)+1}"
        f.append(base)

    # ---------------- MUSICALES ----------------
    fr_mus = ["morning", "musical"]
    for d in _dur_cuna:
        nombre = 'Cuña hasta de 5"' if d == '5"' else f'Cuña de {d}'
        add(familia="musical", tipo="Grabados", tipo_id="grabados", nombre=nombre, duracion=d,
            descripcion='N/A. El material debe entregarse previamente grabado y ajustado a la duración exacta establecida para el formato.',
            franjas=fr_mus, franjaTexto="Su ejecución puede realizarse en los siguientes espacios:",
            specs=[{"t": "Formato de entrega", "d": "Archivos en MP3."},
                   {"t": "Ubicación", "d": UBICACION_BREAK}],
            consejo=_consejo_cuna(d))
    for d, p in _leidas:
        nombre = f'Mención leída de {d}'
        add(familia="musical", tipo="Leídos", tipo_id="leidos", nombre=nombre, duracion=d, palabras=p, maxPalabras=p,
            descripcion=f'{p} palabras como máximo para la duración contratada.',
            franjas=fr_mus, franjaTexto="Su ejecución puede realizarse en los siguientes espacios:",
            specs=[{"t": "Texto", "d": "Se entrega un texto ajustado a la cantidad de palabras correspondiente a la duración contratada."},
                   {"t": "Interpretación", "d": "Son interpretadas por el locutor de turno, quien desarrolla el mensaje de acuerdo con el material y las indicaciones entregadas."},
                   {"t": "Ubicación", "d": UBICACION_BREAK}],
            consejo=_consejo_leida(d))
    add(familia="musical", tipo="Secciones", tipo_id="secciones", nombre='Sección de 60"', duracion='60"', maxPalabras=98,
        descripcion="Estructura del formato en tres partes.",
        componentes=["Presentación de la franja: de 5 a 10 segundos (texto).",
                     "Desarrollo del contenido: texto de máximo 98 palabras o presentación en formato de bullets.",
                     "Cierre: cuña grabada o texto leído de máximo 30 segundos."],
        franjas=["morning"],
        franjaTexto="Este formato se realiza exclusivamente en el Morning Show de cada emisora, de lunes a viernes, entre las 6:00 a. m. y las 10:00 a. m.",
        protegido=True,
        specs=[{"t": "Formato protegido", "d": "Espacio exclusivo dentro del Morning, en el que los locutores de la mesa de trabajo generan conversación, interacción y participación alrededor de un tema alineado con los objetivos y el mensaje de la campaña."}])
    add(familia="musical", tipo="Individual", tipo_id="individual", nombre='Individual de 30"', duracion='30"', maxPalabras=64,
        descripcion="Dos opciones de armado; la pieza completa no puede superar los 30 segundos.",
        componentes=["Opción 1: material completamente grabado, con una duración máxima de 30 segundos.",
                     "Opción 2: combinación de contenido leído y grabado, sin superar los 30 segundos en total. Se recomienda distribuirlo así: 20 segundos de texto leído (aproximadamente 64 palabras) + 10 segundos de material grabado para cerrar el formato.",
                     "Importante: la duración total de la pieza no debe superar los 30 segundos."],
        franjas=["musical"],
        franjaTexto="Este formato se realiza exclusivamente en la programación musical de cada emisora, de lunes a domingo, entre las 10:00 a. m. y las 9:00 p. m.",
        franjaHorarios={"musical": "Lunes a domingo, de 10:00 a. m. a 9:00 p. m."},
        protegido=True,
        specs=[{"t": "Formato protegido", "d": "Espacio exclusivo dentro de la programación musical: suena entre canción y canción. El locutor o los locutores del horario correspondiente realizan la mención de manera individual o conjunta, según la dinámica y las características del espacio."}])
    add(familia="musical", tipo="Remotos", tipo_id="remotos", nombre="Remoto de 2 horas", duracion="2 horas", maxPalabras=300,
        descripcion="Un talento de la emisora musical acompaña al anunciante en el sitio que defina durante 2 horas.",
        componentes=["Presencia de un talento de la emisora musical durante 2 horas (de acuerdo con disponibilidad) en el sitio definido por el anunciante.",
                     "Total 6 salidas de hasta 2 minutos: 4 salidas comerciales hechas por el talento + 2 entrevistas presenciales (la entrevista es opcional).",
                     "10 promociones en la programación musical de la emisora seleccionada, convocando dos días antes del remoto."],
        franjas=["musical"],
        franjaTexto="Se realiza de lunes a domingo, a partir de las 10:00 a. m., mediante una reserva de 2 horas de actividad en el lugar indicado por el cliente.",
        franjaHorarios={"musical": "Lunes a domingo, desde las 10:00 a. m. (reserva de 2 horas)"},
        protegido=True,
        specs=[{"t": "Formato protegido", "d": "Exclusivo para la marca durante la activación."},
               {"t": "Palabras de las promociones", "d": "Promoción de máximo 20 segundos: 54 palabras. La marca se vincula dentro del texto, como menciones leídas por el talento de la emisora."},
               {"t": "Palabras de las salidas del remoto", "d": "Cada transmisión tiene una interacción de máximo 2 minutos, equivalente a un promedio de 300 palabras. Ten presente que el locutor también hace preámbulos y estos aumentan la duración: para mayor efectividad, el material no debe sobrepasar 300 o 360 palabras."}])

    # ---------------- HABLADA (Caracol Radio) ----------------
    todas = PROGRAMAS_HABLADA
    nota_prog = "Se realiza dentro de los programas de Caracol Radio, de acuerdo con la programación y los horarios establecidos en la parrilla de cada emisora."
    for d in _dur_cuna:
        nombre = 'Cuña hasta de 5"' if d == '5"' else f'Cuña de {d}'
        add(familia="hablada", tipo="Grabados", tipo_id="grabados", nombre=nombre, duracion=d,
            descripcion='N/A. El material debe entregarse previamente grabado y ajustado a la duración exacta establecida para el formato.',
            franjas=todas, franjaTexto=nota_prog,
            specs=[{"t": "Formato de entrega", "d": "Archivos en MP3."}],
            consejo=_consejo_cuna(d))
    for d, p in _leidas:
        add(familia="hablada", tipo="Leídos", tipo_id="leidos", nombre=f'Mención leída de {d}', duracion=d, palabras=p, maxPalabras=p,
            descripcion=f'{p} palabras como máximo para la duración contratada.',
            franjas=todas,
            franjaTexto="Modalidad de producción: este formato no se realiza en vivo desde la emisora. Se produce previamente mediante una grabación en voz en off, utilizando una voz autorizada por la emisora.",
            specs=[{"t": "Texto", "d": "Se entrega un texto ajustado a la cantidad de palabras correspondiente a la duración contratada."},
                   {"t": "Interpretación", "d": "Son interpretadas por el locutor de turno, quien desarrolla el mensaje de acuerdo con el material y las indicaciones entregadas."},
                   {"t": "Ubicación", "d": UBICACION_BREAK}],
            consejo=_consejo_leida(d))
    add(familia="hablada", tipo="Secciones", tipo_id="secciones", nombre='Sección de 60"', duracion='60"',
        descripcion="Patrocinio de una noticia o sección dentro de los programas de Caracol Radio.",
        componentes=["Presentación de la franja: de 5 a 10 segundos (grabado por una voz de Caracol).",
                     "Cierre: cuña grabada o texto leído de máximo 30 segundos."],
        franjas=todas,
        franjaTexto="Formato de patrocinio de una noticia o sección dentro de los programas de Caracol Radio, cuyo enfoque temático se alinea con el concepto de la campaña. Cada programa cuenta con una franja destinada a este tipo de espacios. Importante: el contenido editorial no puede ser intervenido por el cliente. Los programas La Luciérnaga (de lunes a viernes) y A Vivir Que Son Dos Días (sábados y domingos) ofrecen mayores posibilidades de integración y participación de marca.",
        specs=[{"t": "Presentación de marca", "d": "Apertura del espacio con una mención que identifica y presenta la marca."},
               {"t": "Desarrollo del contenido", "d": "El noticiero o programa desarrolla su propio contenido informativo y editorial, de acuerdo con su línea temática. Este contenido es de manejo exclusivo de la emisora y no puede ser intervenido ni modificado por el cliente."},
               {"t": "Cierre comercial", "d": "Finaliza con la emisión de una cuña grabada de 30 segundos de la marca."}])

    sem = "Una semana (5 días continuos), lunes a viernes"
    mes = "Mes completo (8 días: 4 sábados y 4 domingos continuos)"
    pres15 = "Presentación diaria del programa pregrabada con duración de 15\" durante 5 días."
    desp15 = "Despedida diaria del programa pregrabada con duración de 15\" durante 5 días."
    pres15f = "Presentación diaria del programa con duración de 15\" durante 8 días (sábados y domingos del mes)."
    desp15f = "Despedida diaria del programa, grabada por la W, con duración de 15\" durante 8 días (sábados y domingos del mes)."
    seccion5 = "1 sección diaria durante 5 días. Cada sección incluye presentación de 10\" + cierre de sección con cuña hasta de 30\"."
    seccion8 = "1 sección diaria durante 8 días (sábados y domingos del mes). Cada sección incluye presentación de 10\" + cierre de sección con cuña hasta de 30\"."
    pat_nota = "Patrocinio de programa: la marca acompaña el programa completo durante el periodo contratado."

    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios", nombre="Patrocinio programa nacional 6AM – W", duracion=sem,
        descripcion="Programa nacional de 5:00 a. m. a 12:00 m.", franjas=["6am-w"], franjaTexto=None,
        componentes=[pres15, desp15, seccion5, "1 cuña diaria de hasta 30\" durante 5 días.", "1 Momento CRC de hasta 60\"."])
    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios",
        nombre="Patrocinio de programas nacionales: Noticiero ½ día, Pulso del Fútbol, Hora 20, La Polémica o AS Mundial", duracion=sem,
        descripcion="Elige uno de los cinco programas en la franja de pauta.",
        franjas=["noticiero", "pulso", "hora20", "polemica", "as-mundial"],
        componentes=[pres15, desp15, "1 cuña diaria de hasta 30\" durante 5 días."])
    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios", nombre="Patrocinio programa nacional Dos Puntos", duracion=sem,
        descripcion="Programa nacional de 2:00 a 4:00 p. m.", franjas=["dos-puntos"],
        componentes=[pres15, desp15, seccion5, "1 Momento CRC de hasta 60\"."])
    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios", nombre="Patrocinio programa nacional La Luciérnaga", duracion=sem,
        descripcion="Programa nacional de 4:00 a 7:00 p. m.", franjas=["luciernaga"],
        componentes=[pres15, desp15, seccion5, "1 Momento CRC de hasta 60\"."])
    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios",
        nombre="Patrocinio de programas nacionales de fin de semana: De Película, Al Campo, Planeta Caracol o Noticieros FDS", duracion=mes,
        descripcion="Elige uno de los programas en la franja de pauta.",
        franjas=["de-pelicula", "al-campo", "planeta-caracol", "noticieros-fds"],
        componentes=[pres15f, desp15f, seccion8])
    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios", nombre="Patrocinio programa nacional A Vivir Que Son Dos Días", duracion=mes,
        descripcion="Programa nacional de 8:00 a 11:00 a. m.", franjas=["a-vivir"],
        componentes=[pres15f, desp15f, seccion8, "2 formatos de Momentos CRC de hasta 60\"."])
    add(familia="hablada", tipo="Patrocinios", tipo_id="patrocinios", nombre="Patrocinio programa nacional Noche Fantástica",
        duracion="Viernes y sábados, de 10:00 p. m. a 4:00 a. m.",
        descripcion="Programa nacional de fin de semana en horario nocturno.", franjas=["noche-fantastica"],
        componentes=[pres15, desp15, "5 cuñas diarias de hasta 30\" durante 5 días."])

    lun_vie = "Lunes a viernes"
    add(familia="hablada", tipo="Vitrinas y esquinas", tipo_id="vitrinas", nombre="Vitrina nacional 6AM – W", duracion=f"{lun_vie}, 7:00 a. m. a 12:00 m. (5 horas)",
        descripcion="Un talento Caracol acompaña a la marca durante 5 horas en el sitio que defina el anunciante, o desde cabina.", franjas=["6am-w"],
        componentes=["Presencia de un talento Caracol durante 5 horas (7:00 a. m. a 12:00 m.) en el sitio definido por el anunciante, o desde cabina.",
                     "5 salidas comerciales de hasta 1 minuto hechas por el talento el día de la vitrina + 1 entrevista presencial de hasta 2 minutos.",
                     "3 promociones en cada uno de los siguientes programas: 6AM – W, Noticiero ½ día, El Pulso del Fútbol, Dos Puntos y La Luciérnaga.",
                     "5 contenidos de hasta 30\" en cada programa: 6AM – W, Dos Puntos y La Luciérnaga.",
                     "2 Instagram Stories de convocatoria que se publican antes de la actividad, elaborados por nuestro equipo digital."])
    add(familia="hablada", tipo="Vitrinas y esquinas", tipo_id="vitrinas", nombre="Vitrina nacional «Plus» 6AM – W", duracion=f"{lun_vie}, 7:00 a. m. a 12:00 m.",
        descripcion="Contenido exclusivo de la marca generado por el talento Caracol desde la cabina de 6AM – W.", franjas=["6am-w"],
        componentes=["Desde la cabina de Caracol, nuestro talento genera contenido exclusivo de la marca en 6AM – W.",
                     "3 salidas comerciales de hasta 1 minuto hechas por el talento Caracol el día de la vitrina Plus.",
                     "1 entrevista digital a un vocero de la marca de hasta 10 minutos, desarrollada desde los estudios digitales de Prisa Media o vía Zoom, y posteriormente amplificada en las redes sociales de la emisora.",
                     "3 promociones en cada uno de los siguientes programas: 6AM – W, Noticiero ½ día, El Pulso del Fútbol, Dos Puntos y La Luciérnaga."])
    add(familia="hablada", tipo="Vitrinas y esquinas", tipo_id="vitrinas", nombre="Esquina Caracol nacional Dos Puntos", duracion=f"{lun_vie}, 2:00 a 4:00 p. m. (2 horas)",
        descripcion="Un talento Caracol durante las 2 horas del programa, en el sitio del anunciante o desde cabina.", franjas=["dos-puntos"],
        componentes=["Presencia de un talento Caracol durante las 2 horas del programa en el sitio definido por el anunciante, o desde cabina.",
                     "Total 4 salidas al aire de hasta 1 minuto: 3 salidas hechas por el talento + 1 entrevista.",
                     "2 promociones en 6AM – W.",
                     "3 promociones en cada uno de los siguientes programas: El Pulso del Fútbol, Dos Puntos, La Luciérnaga y Hora 20."])
    add(familia="hablada", tipo="Vitrinas y esquinas", tipo_id="vitrinas", nombre="Esquina nacional La Luciérnaga", duracion=f"{lun_vie}, 4:00 a 7:00 p. m. (3 horas)",
        descripcion="Un talento Caracol durante 3 horas en el sitio del anunciante o desde cabina.", franjas=["luciernaga"],
        componentes=["Presencia de un talento Caracol durante 3 horas (4:00 p. m. a 7:00 p. m.) en el sitio definido por el anunciante, o desde cabina.",
                     "Total 6 salidas de hasta 1 minuto: 4 salidas comerciales hechas por el talento + 2 entrevistas presenciales.",
                     "3 promociones en cada uno de los siguientes programas: 6AM – W, Dos Puntos, La Luciérnaga y Hora 20."])
    add(familia="hablada", tipo="Vitrinas y esquinas", tipo_id="vitrinas", nombre="Esquina nacional Caracol Fin de Semana", duracion="Sábado o domingo (2 horas)",
        descripcion="No aplica en la franja de 12:00 m. a 1:00 p. m.", franjas=["caracol-fds"],
        componentes=["Presencia de un talento Caracol durante 2 horas en el sitio del anunciante o desde cabina, programadas en Caracol Fin de Semana.",
                     "Total 4 salidas al aire de hasta 1 minuto: 3 salidas hechas por el talento + 1 entrevista.",
                     "2 promociones en 6AM – W, Dos Puntos, La Luciérnaga, Noche Fantástica y A Vivir Que Son Dos Días."])
    return f


# --------------------------------------------------------------------------
# DIGITAL
# --------------------------------------------------------------------------
DEMO = ("Recomendamos al anunciante tener en cuenta el demo de la pieza para que se ajuste a las medidas del portal "
        "(en caso de no tener el demo, por favor pídaselo a su ejecutivo).")
ARCH_STD = "HTML5 · JPG · GIF · PNG"


def P(pieza, medida, peso, archivo, obs="", expande=""):
    return {"pieza": pieza, "medida": medida, "expande": expande, "peso": peso, "archivo": archivo, "obs": obs}


def _digital():
    d = []

    def add(cat, nombre, piezas, notas=None, emisoras=None, aplica="Todas las emisoras", enlaces=None, pasos=None, texto=None):
        d.append({"id": f"d{len(d)+1}", "categoria": cat, "nombre": nombre, "piezas": piezas, "notas": notas or [],
                  "emisoras": emisoras or [], "aplica": aplica, "enlaces": enlaces or [], "pasos": pasos or [], "texto": texto})

    W = "Display web · Tradicionales"
    add(W, "Robapágina", [P("Robapágina", "300 × 250 px", "100 KB", ARCH_STD)])
    add(W, "Robapágina doble", [P("Robapágina doble", "300 × 600 px", "100 KB", ARCH_STD)])
    add(W, "Top Banner", [P("Top Banner", "728 × 90 px", "100 KB", ARCH_STD, "No aplica para el primer scroll de la página.")])
    add(W, "Gigabanner", [P("Gigabanner", "980 × 90 px", "100 KB", ARCH_STD)])

    G = "Display web · Grandes formatos"
    add(G, "Billboard", [P("Billboard", "980 × 250 px", "100 KB", ARCH_STD)])
    add(G, "Sync ad con expander", [
        P("Sync ad 1 (inicial + expandida)", "300 × 250 px", "100 KB entre las dos piezas", ARCH_STD, "Aplica para internas.", "600 × 300 px"),
        P("Sync ad 2 (inicial + expandida)", "728 × 90 px", "100 KB entre las dos piezas", ARCH_STD, "", "728 × 300 px"),
    ], notas=["Se compone de 4 piezas: dos tamaños iniciales y sus dos versiones expandidas."])
    add(G, "Toma de internas Caracol", [
        P("Billboard (1 pieza)", "980 × 250 px", "100 KB", ARCH_STD),
        P("Laterales (2 piezas)", "120 × 600 px", "100 KB", "JPG · PNG"),
    ], notas=["Toma de páginas internas. Se compone de 1 Billboard y 2 laterales."], emisoras=["caracol"], aplica="Caracol Radio")

    BD = "Brand Day"
    add(BD, "Brand Day Caracol Radio", [
        P("Backing Caracol Desktop", "Skin U invertida", "100 KB", "JPG · GIF", DEMO),
        P("Laterales Desktop (apoyo en internas)", "120 × 600 px", "100 KB", "JPG · GIF"),
        P("Banner mobile", "320 × 100 px", "100 KB", "JPG · HTML5 · GIF"),
        P("Roba página mobile", "320 × 100 px ó 300 × 250 px", "100 KB", "JPG · HTML5 · GIF"),
    ], emisoras=["caracol"], aplica="Caracol Radio", notas=[DEMO])
    add(BD, "Brand Day W Radio y Radioacktiva", [
        P("Lateral derecho", "422 × 900 px", "100 KB", "JPG · GIF", DEMO),
        P("Lateral izquierdo", "421 × 900 px", "100 KB", "JPG · GIF"),
        P("Banner superior", "980 × 250 px", "100 KB", "JPG · GIF"),
        P("Laterales Desktop (apoyo en internas)", "120 × 600 px", "100 KB", "JPG · GIF"),
        P("Banner mobile", "320 × 100 px", "100 KB", "JPG · HTML5 · GIF"),
        P("Roba página mobile", "320 × 100 px ó 300 × 250 px", "100 KB", "JPG · HTML5 · GIF"),
    ], emisoras=["radioacktiva"], aplica="W Radio y Radioacktiva", notas=[DEMO])
    add(BD, "Brand Day Tropicana · Bésame · Oxígeno · Los40", [
        P("Backing musicales Desktop", "Skin U invertida", "100 KB", "JPG · GIF", DEMO),
        P("Laterales Desktop (apoyo en internas)", "120 × 600 px", "100 KB", "JPG · GIF"),
        P("Banner mobile", "320 × 100 px", "100 KB", "JPG · HTML5 · GIF"),
        P("Roba página mobile", "320 × 100 px ó 300 × 250 px", "100 KB", "JPG · HTML5 · GIF"),
    ], emisoras=["tropicana", "besame", "los40"], aplica="Tropicana, Bésame, Oxígeno y Los40", notas=[DEMO])

    add("Audio y video en streaming", "Preroll de audio en el player de streaming", [
        P('Preroll de audio 30"', '30"', "2 MB", "MP3", "Debe llevar la pata de inicio."),
        P('Preroll de audio 20"', '20"', "2 MB", "MP3", "Debe llevar la pata de inicio."),
        P('Preroll de audio 15"', '15"', "2 MB", "MP3", "Debe llevar la pata de inicio."),
    ])
    add("Audio y video en streaming", "Preroll de video en el player de streaming", [
        P('Preroll de video 30"', '30"', "2 MB", "MP4"),
        P('Preroll de video 20"', '20"', "2 MB", "MP4"),
        P('Preroll de video 15"', '15"', "2 MB", "MP4"),
    ])

    M = "Mobile"
    add(M, "Banner inferior mobile", [P("Banner inferior", "320 × 50 px", "40 KB", "HTML5 · JPG · GIF")])
    add(M, "Robapágina doble mobile", [P("Robapágina doble", "300 × 600 px", "40 KB", "HTML5 · JPG · GIF")])
    add(M, "Robapágina mobile", [P("Robapágina", "300 × 250 px", "40 KB", "HTML5 · JPG · GIF")])
    add(M, "Banner mobile grande", [P("Banner mobile grande", "320 × 100 px", "40 KB", "HTML5 · JPG · GIF")])

    R = "Redes sociales"
    reglas = ["El post no puede tener URL que direccione al portal del anunciante.",
              "Una marca solo puede tener 1 publicación diaria en el fanpage de la emisora."]
    add(R, "Facebook", [
        P("Post", "900 × 900 px", "", "Imagen"),
        P("Post video", "900 × 900 px", "", "MP4", "Máximo 5 minutos."),
        P("Historia", "750 × 1334 px", "", "Imagen"),
        P("Historia video", "750 × 1334 px", "", "MP4", "Máximo 15 segundos."),
    ], notas=reglas, texto="240 caracteres con espacios. Debe incluir el hashtag #Aliados y la @ oficial del cliente.")
    add(R, "Twitter (X)", [
        P("Tweet", "900 × 900 px", "", "Imagen"),
        P("Tweet video", "900 × 900 px", "", "MP4", "Máximo 2 minutos."),
    ], notas=reglas, texto="140 caracteres con espacios. Debe incluir el hashtag #AliadosW o #AliadosCaracol.")
    add(R, "Instagram", [
        P("Post", "900 × 900 px", "", "Imagen"),
        P("Post video", "1080 × 1350 px", "", "MP4", "Máximo 1 minuto."),
        P("Historia", "750 × 1334 px", "", "Imagen"),
        P("Historia video", "1080 × 1350 px", "", "MP4", "Máximo 15 segundos."),
        P("IGTV", "Video 750 × 1334 px + imagen 900 × 900 px", "", "MP4 + imagen"),
    ], notas=reglas, texto="240 caracteres con espacios. Debe incluir el hashtag #AliadosW o #AliadosCaracol.")

    C = "Contenidos en Facebook"
    brief = "Diligenciar el formato brief desarrollando el tema del que va a tratar el contenido."
    add(C, "Contenido patrocinado", [
        P("Imagen para emisoras habladas", "1080 × 750 px", "", "Imagen"),
        P("Imagen para emisoras musicales", "720 × 300 px", "", "Imagen"),
    ], notas=[brief, "Tiempo de desarrollo: 3 días hábiles.",
              "Procura que las imágenes no tengan textos: los cortes automáticos de la plataforma pueden dejarlas cortadas."])
    add(C, "Video TXT", [], notas=[brief, "Logo en vectores + manual de marca.", "Recursos gráficos de la marca (fotos o videos).",
                                    "Tiempo de desarrollo: 6 días hábiles."])
    add(C, "Infografía", [], notas=[brief, "Logo en vectores + manual de marca.", "Recursos gráficos de la marca.",
                                     "Tiempo de desarrollo: 6 días hábiles.",
                                     "Si es el cliente quien desarrolla la infografía, debe medir 1100 px de ancho."])
    add(C, "Agrupador de contenidos", [], notas=["Nombre del agrupador.", "Descripción de dos renglones de la temática del agrupador."],
        enlaces=[{"t": "Caracol Radio", "u": "https://www.facebook.com/caracolradio"},
                 {"t": "W Radio", "u": "https://www.facebook.com/WRadioCo"},
                 {"t": "Tropicana", "u": "https://www.facebook.com/TropicanaColombia"},
                 {"t": "Bésame", "u": "https://www.facebook.com/BesameRadioColombia"},
                 {"t": "Los40", "u": "https://www.facebook.com/LOS40Colombia"},
                 {"t": "Radioacktiva", "u": "https://www.facebook.com/radioacktiva"}])
    add(C, "Retransmisiones en vivo por Facebook", [], pasos=[
        "Enviar la solicitud de crossposting al fan page de la emisora en la cual se va a retransmitir.",
        "Compartir el link de la invitación con Brand Estudio.",
        "Al lanzar su live, el cliente debe seleccionar en configuraciones «retransmitir en nuestro Fan Page» para que salga automáticamente."])
    return d


# --------------------------------------------------------------------------
# RECOMENDACIONES ESTRATÉGICAS (criterio de planificación de medios)
# No vienen del archivo de especificaciones: son pautas generales de buenas
# prácticas ATL/digital. Cada ítem es {t: rótulo, d: texto}. Ajustar con el equipo comercial.
# --------------------------------------------------------------------------
POR = [("Una sola mención", "En piezas de menos de 30 segundos no hay tiempo para repetir la marca sin sacrificar el mensaje: desarrolla la idea y cierra con el nombre, idealmente con un sonido o jingle de marca."),
       ("Marca nueva", "Si la marca es nueva o poco conocida, considera nombrarla también al inicio: los estudios de la industria recomiendan mencionarla pronto.")]


def _e(marca=None, mensaje=None, cta=None, evita=None, extra=None):
    out = []
    if marca: out.append({"t": "Mención de la marca", "d": marca})
    if mensaje: out.append({"t": "Mensaje", "d": mensaje})
    if cta: out.append({"t": "Llamado a la acción", "d": cta})
    if extra:
        for x in ([extra] if isinstance(extra[0], str) else extra): out.append({"t": x[0], "d": x[1]})
    if evita: out.append({"t": "Evita", "d": evita})
    return out


_CUNA = {
    '5"': _e("Una sola vez, al cierre. En 5 segundos la marca es el mensaje: queda como lo último que escucha el oyente, apoyada en un sonido o jingle distintivo.",
             "Una sola idea (marca + beneficio o frase de posicionamiento). Úsala para mantener la recordación, no para explicar.",
             evita="Precios, direcciones o teléfonos: no alcanzan a retenerse.", extra=POR),
    '10"': _e("Una sola vez, al cierre: la pieza desarrolla la idea y termina nombrando la marca.",
              "Un beneficio y, si cabe, un llamado muy corto. Ideal para jingle o frase de posicionamiento.",
              evita="Meter dos ideas en una sola pieza.", extra=POR),
    '15"': _e("Una sola vez, al cierre, junto al llamado a la acción.",
              "Una idea central con su beneficio.", "Uno simple y fácil de recordar, por ejemplo «búscalo en…».",
              "Datos largos como URL o teléfonos.", extra=POR),
    '20"': _e("Una sola vez, al cierre, junto al llamado a la acción.",
              "Beneficio principal, una razón para creerlo y cierre con la marca.", "Al final, junto a la marca.",
              "Saturar con más de un beneficio.", extra=POR),
    '30"': _e("Al inicio (primeros 3 a 5 segundos) y al cierre. Si la pieza es narrativa, una mención intermedia refuerza el recuerdo.",
              "Gancho, desarrollo de un solo beneficio y cierre. Es la duración con más espacio para contar una historia o un lanzamiento.",
              "Claro y único, en el cierre.", "Dejar la marca solo al final: si el oyente se distrae, no la recibe."),
    '35"': _e("Al inicio, a la mitad y al cierre (tres menciones).",
              "Permite apoyarse en un ejemplo o testimonio. Mantén una sola promesa central.",
              "En el cierre.", "Alargar el texto sin ritmo: genera fatiga. Rota 2 o 3 versiones durante la campaña."),
    '60"': _e("Al inicio, una vez durante el desarrollo y en el cierre.",
              "Estructura de gancho, historia o diálogo y cierre. Recomendado para lanzamientos, servicios o productos que requieren explicación.",
              "En el cierre, con la marca.", "Repetir la misma versión muchas veces: alterna versiones para evitar el desgaste."),
}
_CUNA['40"'] = _CUNA['45"'] = _CUNA['35"']

_LEIDA = {
    "corta": _e("Una sola vez, al cierre: que el locutor desarrolle la idea y termine nombrando la marca.",
                "Un solo dato clave en lenguaje conversacional: el locutor pone la voz y tú pones la idea.",
                evita="Redactar como aviso publicitario: se nota que es un texto leído.",
                extra=[POR[0], ("Brief al locutor", "Comparte la pronunciación de la marca y las 2 o 3 ideas que no pueden faltar.")]),
    "media": _e("Al inicio y al cierre; en 30\" puedes añadir una mención intermedia.",
                "Beneficio, razón para creer y llamado a la acción; escribe como se habla, con frases cortas.",
                "Una sola acción, en el cierre.", "Pasarte del número de palabras: el locutor acelera y pierde claridad.",
                extra=("Brief al locutor", "Comparte la pronunciación de la marca y las 2 o 3 ideas que no pueden faltar.")),
    "larga": _e("Al inicio, a la mitad y al cierre.",
                "Cuenta una situación cotidiana o un caso: el locutor aporta credibilidad personal.",
                evita="Listas de datos o precios: se olvidan al escucharlos.",
                extra=("Brief al locutor", "Comparte la pronunciación de la marca y las 2 o 3 ideas que no pueden faltar.")),
}


def _estrategia_radio(f):
    t, n = f["tipo"], f["nombre"]
    if t == "Grabados":
        return _CUNA[f["duracion"]]
    if t == "Leídos":
        return _LEIDA["corta" if f["duracion"] in ('5"', '10"', '15"', '20"') else "media" if f["duracion"] == '30"' else "larga"]
    if t == "Secciones" and f["familia"] == "musical":
        return _e("En la presentación de la franja (apertura) y en el cierre.",
                  "El tema debe interesar al oyente por sí solo (consejo, dato útil, tendencia) y la marca aparece como quien lo hace posible.",
                  "Resérvalo para el cierre de 30\", no para el desarrollo.",
                  "Convertir la sección en un anuncio largo: pierde credibilidad con la mesa de locutores.")
    if t == "Secciones":
        return _e("En la presentación grabada y de nuevo en la cuña de cierre.",
                  "Elige un tema afín al concepto de campaña. El contenido editorial es de la emisora: tú aportas el ángulo, no el texto.",
                  "En la cuña de cierre de 30\".",
                  "Intentar intervenir el contenido editorial: no es posible y afectaría la credibilidad.",
                  extra=("Más integración", "La Luciérnaga y A Vivir Que Son Dos Días ofrecen más posibilidades de participación de marca."))
    if t == "Individual":
        return _e("El locutor la nombra al inicio y la repite al cierre.",
                  "Conversacional y específico para ese horario; sirve para posicionamiento, ofertas o recordatorios.",
                  "Uno solo, simple y al final.",
                  "Un texto sin ritmo: entra entre canción y canción y debe sonar con la energía de la emisora.")
    if t == "Remotos":
        return _e("En la primera frase de cada salida y en las promociones, siempre junto al lugar de la activación.",
                  "Antes: convoca con lugar, hora y beneficio (las 10 promociones, desde dos días antes). Durante: salidas que inviten a pasar por el punto.",
                  "«Ven hoy a…» con un incentivo claro: dinámica, premio o descuento.",
                  "Superar 300 a 360 palabras por salida o llegar sin una dinámica para el público presente.")
    if t == "Patrocinios":
        return _e("En la presentación y en la despedida diaria: la marca abre y cierra el programa. Asocia el nombre al programa, no solo a un producto.",
                  "Elige un programa afín a tu público (puedes apoyarte en Recomendar emisora) y una idea que se pueda ir contando durante el periodo.",
                  "En la cuña o el momento diario de la marca, no en la presentación.",
                  "Repetir la misma cuña todos los días: cambia el mensaje o su orden a lo largo de la semana.")
    if t == "Vitrinas y esquinas":
        return _e("El talento la menciona al comenzar cada salida y al despedirla.",
                  "Aprovecha el contacto con la gente: demostración, entrevista a un vocero, historias reales de clientes.",
                  "Convoca a la activación en las promociones previas: lugar, hora y beneficio.",
                  "Salidas sin tema nuevo: prepara puntos de conversación distintos para cada una.")
    return []


def _estrategia_digital(f):
    c, n = f["categoria"], f["nombre"]
    if c.startswith("Display web"):
        if n in ("Top Banner", "Gigabanner", "Billboard"):
            msg = "Formato horizontal: ideal para mensajes cortos (3 a 5 palabras) con el logo siempre a la vista."
        elif n == "Robapágina doble":
            msg = "Formato vertical: da espacio para producto y beneficio. Muy bueno para lanzamientos."
        elif n == "Robapágina":
            msg = "Formato versátil: sirve para marca y para respuesta. Imagen sencilla y un botón claro."
        elif n.startswith("Sync"):
            msg = "La pieza inicial debe funcionar sola; usa la versión expandida para el detalle o la promoción."
        else:
            msg = "Presencia dominante en páginas internas: usa un solo concepto creativo en las tres piezas."
        return _e("Logo visible desde el primer segundo y en el último cuadro de la animación.", msg,
                  "Un botón o frase con verbo claro (Conoce, Descubre, Compra).",
                  "Texto pequeño o varias ideas. Cuida el peso máximo optimizando las imágenes.")
    if c == "Brand Day":
        return _e("Domina el portal ese día: mismo logo y mismo concepto en todas las piezas (desktop y mobile).",
                  "Una sola oferta o noticia; úsalo para lanzamientos o fechas clave.",
                  "Sincronízalo con pauta de radio el mismo día para reforzar el mensaje.",
                  "Diseños distintos entre piezas. Revisa el demo para que no se corten elementos del skin.")
    if c == "Audio y video en streaming":
        if "audio" in n:
            return _e("En 15\" y 20\": una sola vez, al cierre. En 30\": al inicio (después de la «pata» de inicio) y al cierre.",
                      "Cercano y directo: el oyente de streaming está atento y normalmente con audífonos. En 15\" una idea; en 30\", una mini historia.",
                      "Uno solo, simple y fácil de recordar.",
                      "Copiar la cuña de radio sin adaptarla: el contexto es más íntimo.")
        return _e("Si tu objetivo es recordación, logo y marca en los primeros 5 segundos y refuerzo al final. Si es notoriedad o consideración, puede entrar un poco después priorizando la conexión con el espectador.",
                  "Cuenta la idea de forma visual y agrega subtítulos: muchos usuarios ven video sin sonido. 15\" para recordación, 30\" para contar una historia.",
                  "Cierra con logo y llamado visibles.",
                  "Dejar el mensaje clave para el final.")
    if c == "Mobile":
        return _e("Logo grande y legible desde el primer momento.",
                  "Pantalla pequeña: máximo 5 palabras y una imagen simple.",
                  "Un botón visible y fácil de tocar.",
                  "Texto largo o detalles finos que no se leen en el celular.")
    if c == "Redes sociales":
        if n == "Twitter (X)":
            return _e("Nómbrala en la primera frase; en 140 caracteres no hay espacio para esperar.",
                      "Frase corta con gancho, el hashtag obligatorio y la @ del cliente.",
                      "No se permite URL al portal del anunciante: invita a interactuar o a seguir la @.",
                      "Publicar más de una vez al día: la regla es 1 publicación por marca.")
        return _e("Nómbrala en la primera línea del texto, que es la que más se ve.",
                  "Texto breve con un solo mensaje, hashtag obligatorio y la @ del cliente. En video, lo importante en los primeros 3 segundos y con subtítulos; en historias, formato vertical.",
                  "No se permite URL al portal del anunciante: invita a comentar, compartir o seguir la @.",
                  "Publicar más de una vez al día: la regla es 1 publicación por marca.")
    if n == "Contenido patrocinado":
        return _e("En el título o el primer párrafo, como quien hace posible el contenido.",
                  "Aporta valor primero: un tema que interese a la audiencia y en el que tu marca sea parte de la solución, no un aviso.",
                  "Al final del contenido, no al principio.",
                  "Imágenes con texto: los cortes automáticos de la plataforma pueden dejarlas cortadas.")
    if n == "Video TXT":
        return _e("Al inicio y al cierre del video, con el logo en vectores.",
                  "Guion corto y texto grande y legible: se ve sin audio.",
                  "Una sola acción, al final.",
                  "Frases largas en pantalla o demasiados mensajes.")
    if n == "Infografía":
        return _e("Marca arriba y abajo de la infografía.",
                  "Un dato por bloque, de lo más relevante a lo menos.",
                  evita="Saturarla de texto; respeta el ancho de 1100 px si la desarrollas tú.")
    if n == "Agrupador de contenidos":
        return _e("Con el nombre del agrupador, que ancla la marca a una temática.",
                  "Elige una temática continua y relevante para tu público; el nombre debe ser claro y durable.",
                  evita="Cambiar de tema a mitad de campaña: se pierde la asociación.")
    return _e("Anuncia el live con anticipación y menciona la marca al inicio y cuando lleguen nuevos espectadores.",
              "Un tema con valor para la audiencia, no solo un anuncio de la marca.",
              evita="Lanzar el live sin haber enviado la solicitud de crossposting a tiempo.")


def construir():
    perfiles_meta = [{"id": i, "label": l, "grupo": g} for i, l, g in PERFILES]
    emisoras = []
    for e in EMISORAS:
        e = dict(e)
        e["perfiles"] = [{"id": i, "label": l, "pct": p} for i, l, p in e["perfiles"]]
        emisoras.append(e)
    formatos = _radio()
    for fmt in formatos:
        fmt.pop("tipo_id", None)
        fmt["estrategia"] = _estrategia_radio(fmt)
    digital = _digital()
    for fd in digital:
        fd["estrategia"] = _estrategia_digital(fd)
    return {
        "meta": {
            "titulo": "Guía de formatos Prisa Media",
            "nota_estrategia": "Pautas generales de planificación de medios basadas en estudios y guías de la industria (Nielsen, Spotify, Westwood One, Radiocentre, Google). No hacen parte de las especificaciones técnicas: valídalas con tu ejecutivo.",
            "fuentes": ["FORMATOS PRISA MEDIA – Grupo Nutresa (especificaciones de radio y digital)",
                        "AUDIENCIAS PRISA MEDIA (mapa de afinidad TGI R2 2024; perfiles TGI Choices 2026)"],
        },
        "bandasEdad": BANDAS_EDAD, "nse": NSE, "atajos": ATAJOS,
        "perfiles": perfiles_meta, "emisoras": emisoras, "metricas": [
            {"id": i, "linea1": a, "linea2": b} for i, a, b in METRICAS_ROTULOS],
        "franjas": FRANJAS, "radio": formatos, "digital": digital,
    }


if __name__ == "__main__":
    out = RAIZ / "data" / "formatos.json"
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(construir(), ensure_ascii=False, indent=1), encoding="utf-8")
    print("OK", out)
