"""
Unificador de Expedientes BanEcuador - V5.0

Cambios principales respecto a la V4.1:
  * Deteccion de orientacion en dos etapas:
      1. Tesseract OSD (--psm 0) propone el angulo.
      2. Se verifica con OCR real usando una puntuacion basada en
         palabras validas de alta confianza (el texto girado 90/270
         produce "basura" con pocas palabras validas).
    Si OSD no esta instalado o no es confiable, se evaluan los 4 angulos.
  * La imagen se escala por su lado mayor (antes solo por el ancho), asi
    las fotos verticales/horizontales se evaluan con la misma resolucion.
  * Fotos grandes de celular se reducen antes del OCR (mas rapido y
    menos memoria).
  * OCR con parada temprana: deja de leer cuando el Coactivado y el
    numero de proceso ya fueron confirmados al menos dos veces.
  * Procesamiento de varias imagenes en paralelo (Tesseract corre como
    proceso externo, por lo que los hilos si aprovechan varios nucleos).
  * Busqueda automatica de Tesseract en Windows.
"""

import io
import os
import re
import csv
import zipfile
import tempfile
import threading
import unicodedata

from pathlib import Path, PurePosixPath
from difflib import SequenceMatcher
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

# Tesseract usa OpenMP internamente; con varios hilos propios es mas
# rapido limitar cada proceso a un nucleo. Debe fijarse antes de llamarlo.
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

import numpy as np
import pypdf
import pytesseract

from PIL import (
    Image, ImageOps, ImageSequence, ImageEnhance, ImageFilter, ImageChops
)
from pypdf import PdfReader, PdfWriter, Transformation


# =============================================
# CONFIGURACION GENERAL
# =============================================

VERSION = "5.1"

EXTENSIONES_IMAGEN = {
    ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"
}

ANGULOS = (0, 90, 180, 270)

# Lado mayor (px) de la imagen usada para detectar orientacion.
LADO_ORIENTACION = 2000

# Lado mayor (px) de la imagen usada para el OCR completo.
LADO_OCR = 2600

# Lado mayor (px) maximo de la imagen guardada para el PDF final.
LADO_PDF = 3000

# Confianza minima de Tesseract OSD para confiar en su propuesta.
OSD_CONFIANZA_MINIMA = 1.5

# Si la mejor orientacion candidata puntua menos que esto,
# se evaluan los cuatro angulos.
CALIDAD_MINIMA = 60

# Hilos de trabajo (imagenes procesadas a la vez).
HILOS = max(1, min(6, (os.cpu_count() or 2) - 1))

PALABRAS_CLAVE = (
    "BANECUADOR", "COACTIVADO", "COACTIVADA", "COACTIVA", "PROCESO",
    "DOMICILIO", "OPERACION", "RIOBAMBA", "SUCURSAL", "AGENCIA",
    "CUANTIA", "SECRETARIO", "JUZGADO", "EJECUCION",
)


# =============================================
# CONFIGURACION DE TESSERACT (WINDOWS)
# =============================================

def configurar_tesseract():

    if os.name != "nt":
        return

    posibles = [
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]

    local = os.environ.get("LOCALAPPDATA")

    if local:
        posibles.append(
            Path(local) / "Programs" / "Tesseract-OCR" / "tesseract.exe"
        )

    for ruta in posibles:

        if ruta.exists():

            pytesseract.pytesseract.tesseract_cmd = str(ruta)

            tessdata = ruta.parent / "tessdata"

            if tessdata.exists() and "TESSDATA_PREFIX" not in os.environ:
                os.environ["TESSDATA_PREFIX"] = str(tessdata)

            return


configurar_tesseract()


# =============================================
# NORMALIZAR TEXTO
# =============================================

def quitar_tildes(texto):

    texto = unicodedata.normalize("NFKD", str(texto).upper())

    return "".join(c for c in texto if not unicodedata.combining(c))


def normalizar(texto):

    texto = quitar_tildes(texto)
    texto = re.sub(r"[^A-Z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


# =============================================
# NUMEROS DE PROCESOS
# =============================================

PATRON_PROCESO = re.compile(
    r"(?<![A-Z0-9])"
    r"([0-9OQILSB]{3})"
    r"\s*[-–—.]\s*"
    r"([0-9OQILSB]{4})"
    r"\s*[-–—.]\s*"
    r"([0-9OQILSB]{5})"
    r"(?![A-Z0-9])",
    re.I
)

CORRECCION_NUMEROS = str.maketrans({
    "O": "0", "Q": "0", "I": "1", "L": "1", "S": "5", "B": "8"
})


def extraer_procesos(texto):

    encontrados = set()

    for coincidencia in PATRON_PROCESO.finditer(texto.upper()):

        proceso = "-".join(
            parte.translate(CORRECCION_NUMEROS)
            for parte in coincidencia.groups()
        )

        if re.fullmatch(r"\d{3}-\d{4}-\d{5}", proceso):
            encontrados.add(proceso)

    return encontrados


# =============================================
# EXTRAER NOMBRE DEL ARCHIVO PDF
# =============================================

def nombre_archivo(etiqueta):

    nombre = PurePosixPath(etiqueta).stem
    nombre = PATRON_PROCESO.sub(" ", nombre)
    nombre = re.sub(r"^\s*\d+[\s_.-]+", "", nombre)
    nombre = normalizar(nombre)

    return " ".join(p for p in nombre.split() if p.isalpha())


# =============================================
# EXTRAER EXCLUSIVAMENTE EL COACTIVADO
# =============================================
#
# Las etiquetas se comparan de forma aproximada porque el OCR las lee
# con errores ("Coa: tivado", "Cosctivada", "Cod::udor Coactivado").
# Formatos de caratula soportados:
#   "Coactivado: NOMBRE"          (formato nuevo)
#   "Deudor Principal: NOMBRE"    (formato antiguo)
# Se descartan: Codeudor (Coactivado), Co-deudor, Garante, Juez, etc.

ETIQUETAS_EXCLUIDAS = (
    "CODEUDORCOACTIVADO", "CODEUDOR", "COGARANTE", "GARANTE",
    "JUEZDELEGADODECOACTIVA", "SECRETARIO", "SECRETARIOA", "ABOGADO",
    "FIADOR", "DOMICILIO", "CUANTIA", "FECHADEINICIO",
    "NUMERODEOPERACION",
)

PATRON_FIN_CAMPO = re.compile(
    r"\b(?:DOMICILIO|CODEUDOR|CO DEUDOR|GARANTE|FECHA|CUANTIA|"
    r"SECRETARIO|JUEZ)\b"
)

# Sin dos puntos: "COACTIVADO NOMBRE..." al inicio de la linea.
PATRON_COACTIVADO_SIN_SEPARADOR = re.compile(
    r"^\W{0,3}C[O0]\s?A\s?C?\s?T\s?I\s?V\s?A\s?D\s?[OA0]S?\b\s*[:;.\-]?\s*"
)


def solo_letras(texto):
    return re.sub(r"[^A-Z]", "", quitar_tildes(texto))


def clasificar_etiqueta(etiqueta):
    """'titular', 'excluida' o None segun la etiqueta (texto antes de ':')."""

    letras = solo_letras(etiqueta)

    if len(letras) < 5 or len(letras) > 26:
        return None

    # 1. Formato antiguo: "Deudor Principal".
    if distancia_libre("DEUDORPRINCIPAL", letras) <= 3:
        return "titular"

    # 2. Codeudor en cualquier variante ("Cod::udor Coactivado",
    #    "Cogeudor", "Co-deudor"). Contiene "COACTIVADO", por eso va antes.
    if "UDOR" in letras or "DEUD" in letras:
        return "excluida"

    # 3. Otras etiquetas que no son del titular.
    if min(distancia(e, letras) for e in ETIQUETAS_EXCLUIDAS) <= 2:
        return "excluida"

    for palabra in ("GARANT", "JUEZ", "SECRETAR", "ABOGAD", "FIADOR"):
        if palabra in letras:
            return "excluida"

    # 4. Coactivado / Coactivada, con hasta 2 letras de basura antes
    #    ("| e Coactivado") y hasta 2 errores ("Coa: tivado", "Cosctivada").
    if len(letras) <= 12 and min(
        distancia_libre(e, letras) for e in ("COACTIVADO", "COACTIVADA")
    ) <= 2:
        return "titular"

    return None


def separar_etiqueta(linea):
    """
    Devuelve (tipo, valor) de una linea de la caratula.
    Prueba los primeros separadores ':' o ';' (la etiqueta puede venir
    partida, como en "Coa: tivado: NOMBRE").
    """

    posiciones = [m.start() for m in re.finditer(r"[:;]", linea)][:3]

    for posicion in posiciones:

        tipo = clasificar_etiqueta(linea[:posicion])

        if tipo:
            return tipo, linea[posicion + 1:]

    coincidencia = PATRON_COACTIVADO_SIN_SEPARADOR.search(linea)

    if coincidencia:
        return "titular", linea[coincidencia.end():]

    return None, linea


def limpiar_nombre(texto):

    palabras = [p for p in normalizar(texto).split() if p.isalpha()]

    # Basura al final que deja el subrayado: "A", "SS", "AAA", "O".
    while palabras and (
        len(palabras[-1]) <= 2 or len(set(palabras[-1])) == 1
    ):
        palabras.pop()

    while palabras and len(palabras[0]) == 1:
        palabras.pop(0)

    return palabras


def extraer_coactivados(texto):

    resultados = []

    lineas = [l.strip() for l in quitar_tildes(texto).splitlines()]

    for indice, linea in enumerate(lineas):

        if not linea:
            continue

        tipo, valor = separar_etiqueta(linea)

        if tipo != "titular":
            continue

        nombre = PATRON_FIN_CAMPO.split(valor)[0]

        # Si el OCR coloco el nombre en la siguiente linea.
        if len(limpiar_nombre(nombre)) < 2 and indice + 1 < len(lineas):

            siguiente = lineas[indice + 1]

            if separar_etiqueta(siguiente)[0] is None:
                nombre += " " + siguiente

        # Separar posibles titulares escritos en el mismo campo.
        for parte in re.split(r"\s+-\s+|\s*;\s*", nombre):

            palabras = limpiar_nombre(parte)

            if 2 <= len(palabras) <= 10:

                resultado = " ".join(palabras)

                if resultado not in resultados:
                    resultados.append(resultado)

    return resultados


def lineas_para_busqueda(textos):
    """
    Lineas unicas (solo letras) donde buscar los nombres de los PDF,
    sin las lineas de codeudores, garantes, jueces o secretarios.
    """

    vistas = set()

    for texto in textos:

        for linea in quitar_tildes(texto).splitlines():

            if separar_etiqueta(linea)[0] == "excluida":
                continue

            letras = solo_letras(linea)

            if len(letras) >= 6:
                vistas.add(letras)

    return list(vistas)


# =============================================
# SIMILITUD DE NOMBRES
# =============================================

def distancia(a, b):
    """Distancia de edicion (Levenshtein)."""

    fila = list(range(len(b) + 1))

    for i, x in enumerate(a, 1):

        nueva = [i]

        for j, y in enumerate(b, 1):
            nueva.append(min(
                nueva[-1] + 1,
                fila[j] + 1,
                fila[j - 1] + (x != y)
            ))

        fila = nueva

    return fila[-1]


def distancia_libre(patron, texto, inicio_libre=True, fin_libre=True):
    """
    Distancia de edicion de `patron` contra el mejor fragmento de `texto`.
    inicio_libre y fin_libre: el fragmento puede empezar / terminar en
    cualquier posicion (busqueda aproximada de subcadena).
    """

    if not patron:
        return 0

    fila = [0 if inicio_libre else j for j in range(len(texto) + 1)]

    for i, x in enumerate(patron, 1):

        nueva = [i]

        for j, y in enumerate(texto, 1):
            nueva.append(min(
                nueva[-1] + 1,
                fila[j] + 1,
                fila[j - 1] + (x != y)
            ))

        fila = nueva

    return min(fila) if fin_libre else fila[-1]


def similitud_nombres(nombre_ocr, nombre_pdf):
    """
    Similitud 0-100 entre el nombre leido y el nombre del PDF.
    Tolera:
      * palabras pegadas o separadas ("PRISCILADAYANA")
      * letras mal leidas ("JHONNATAN" / "JONATHAN")
      * nombre del PDF contenido en una lectura mas larga
        (dos titulares juntos o basura al final)
      * lectura truncada al final ("... JUAN" / "... JUAN WILSON")
    """

    a = solo_letras(nombre_ocr)
    b = solo_letras(nombre_pdf)

    if not a or not b:
        return 0.0

    if a == b:
        return 100.0

    # 1. Comparacion completa sin espacios.
    completa = SequenceMatcher(None, a, b).ratio()

    # 2. El nombre del PDF aparece dentro de la lectura.
    contenido = 0.0

    if len(b) >= 8:
        contenido = 1 - distancia_libre(b, a) / len(b)

    # 3. La lectura es el comienzo del nombre del PDF (truncada).
    truncado = 0.0
    cobertura = len(a) / len(b)

    if len(a) >= 10 and cobertura >= 0.6:
        prefijo = 1 - distancia_libre(a, b, inicio_libre=False) / len(a)
        truncado = prefijo * (0.8 + 0.2 * min(cobertura, 1))

    return max(completa, contenido, truncado) * 100


def buscar_en_texto(nombre_pdf, lineas):
    """
    Busca el nombre del PDF en todas las lineas leidas (no solo en el
    campo Coactivado). Devuelve similitud 0-100.
    """

    b = solo_letras(nombre_pdf)

    if len(b) < 8:
        return 0.0

    gramas = {b[i:i + 4] for i in range(len(b) - 3)}
    minimo = max(2, len(gramas) // 4)

    mejor = 0.0

    for linea in lineas:

        if sum(g in linea for g in gramas) < minimo:
            continue

        mejor = max(mejor, 1 - distancia_libre(b, linea) / len(b))

    return mejor * 100


# =============================================
# NUMEROS DE PROCESO APROXIMADOS
# =============================================

def diferencia_procesos(p1, p2):
    """Cantidad de digitos distintos entre dos numeros de proceso."""

    return sum(x != y for x, y in zip(p1, p2))


# =============================================
# LEER ARCHIVOS ZIP
# =============================================

def entradas(z, extensiones):

    return sorted(
        (
            archivo for archivo in z.infolist()
            if not archivo.is_dir()
            and not archivo.filename.startswith("__MACOSX/")
            and not PurePosixPath(archivo.filename).name.startswith("._")
            and PurePosixPath(archivo.filename).suffix.lower() in extensiones
        ),
        key=lambda archivo: archivo.filename.casefold()
    )


def documentos(z, prefijo="", profundidad=0):

    if profundidad > 5:
        raise ValueError("Demasiados archivos ZIP anidados.")

    for archivo in entradas(z, {".pdf", ".zip"}):

        ruta = prefijo + archivo.filename
        data = z.read(archivo)

        if archivo.filename.lower().endswith(".zip"):

            with zipfile.ZipFile(io.BytesIO(data)) as interno:
                yield from documentos(interno, ruta + "/", profundidad + 1)

            continue

        partes = PurePosixPath(ruta).parts

        etiqueta = next(
            (parte for parte in partes if extraer_procesos(parte)),
            partes[-1]
        )

        if etiqueta in partes[:-1]:
            clave = "/".join(partes[:partes.index(etiqueta) + 1])
        else:
            clave = ruta

        yield clave, etiqueta, ruta, data


# =============================================
# PREPARAR IMAGEN
# =============================================

def escalar_lado_mayor(imagen, lado, solo_reducir=False):
    """Escala la imagen para que su lado mayor mida `lado` px."""

    actual = max(imagen.size)

    if actual == lado or (solo_reducir and actual <= lado):
        return imagen

    factor = lado / actual

    return imagen.resize(
        (
            max(1, round(imagen.width * factor)),
            max(1, round(imagen.height * factor))
        ),
        Image.Resampling.LANCZOS
    )


def preparar_ocr(imagen, lado=LADO_OCR):

    gris = ImageOps.autocontrast(ImageOps.grayscale(imagen), cutoff=1)

    return escalar_lado_mayor(gris, lado)


def mejorar_contraste(gris):

    gris = ImageEnhance.Contrast(gris).enhance(1.6)

    return gris.filter(ImageFilter.SHARPEN)


def normalizar_fondo(gris):
    """
    Divide la imagen por su fondo estimado (desenfoque grande).
    Elimina sombras, papel gris y marcas de agua claras: es lo que mas
    mejora la lectura de fotos de celular.
    """

    radio = max(10, max(gris.size) // 60)

    fondo = np.asarray(
        gris.filter(ImageFilter.GaussianBlur(radio)), dtype=np.float32
    ) + 1.0

    imagen = np.asarray(gris, dtype=np.float32) / fondo * 255.0

    return ImageOps.autocontrast(
        Image.fromarray(np.clip(imagen, 0, 255).astype(np.uint8)), cutoff=1
    )


def restar_fondo(gris):
    """Variante por resta, mejor para fotos desenfocadas."""

    radio = max(10, max(gris.size) // 80)

    fondo = gris.filter(ImageFilter.GaussianBlur(radio))

    return ImageOps.autocontrast(
        ImageOps.invert(ImageChops.subtract(fondo, gris)), cutoff=1
    )


# =============================================
# CALIFICAR UNA LECTURA OCR
# =============================================

def calidad_ocr(datos):
    """
    Puntua una lectura OCR. La clave es contar palabras "validas"
    (alfabeticas, de 3+ letras y con confianza alta): un texto leido de
    lado o de cabeza casi no produce palabras asi, aunque Tesseract a
    veces le asigne confianzas medias a la basura.
    """

    confianzas = []
    validas = 0
    lineas = {}

    for i, texto in enumerate(datos["text"]):

        texto = texto.strip()

        try:
            confianza = float(datos["conf"][i])
        except (TypeError, ValueError):
            continue

        if not texto or confianza < 0:
            continue

        confianzas.append(confianza)

        limpio = texto.strip(".,:;()\"'-")

        if confianza >= 60 and len(limpio) >= 3 and limpio.isalpha():
            validas += 1

        clave = (
            datos["block_num"][i],
            datos["par_num"][i],
            datos["line_num"][i]
        )

        lineas.setdefault(clave, []).append(texto)

    texto_final = "\n".join(" ".join(linea) for linea in lineas.values())

    if not confianzas:
        return 0.0, texto_final

    promedio = sum(confianzas) / len(confianzas)

    puntuacion = min(validas, 150) * 2 + promedio * 0.5

    texto_normalizado = normalizar(texto_final)

    puntuacion += 10 * sum(
        palabra in texto_normalizado for palabra in PALABRAS_CLAVE
    )

    if extraer_coactivados(texto_final):
        puntuacion += 30

    if extraer_procesos(texto_final):
        puntuacion += 20

    return puntuacion, texto_final


# =============================================
# DETECTAR ROTACION AUTOMATICA
# =============================================

def orientacion_osd(gris):
    """
    Devuelve (angulo_antihorario, confianza) segun Tesseract OSD,
    o (None, 0) si no esta disponible o no pudo decidir.
    """

    try:

        osd = pytesseract.image_to_osd(
            gris,
            config="--psm 0 -c min_characters_to_try=10",
            output_type=pytesseract.Output.DICT
        )

    except Exception:

        return None, 0.0

    # "rotate" = grados en sentido HORARIO para corregir.
    # PIL.rotate gira en sentido ANTIHORARIO.
    rotar = int(osd.get("rotate", 0)) % 360

    return (360 - rotar) % 360, float(osd.get("orientation_conf", 0))


def evaluar_orientacion(gris, angulo):

    rotada = gris.rotate(angulo, expand=True, fillcolor=255)

    datos = pytesseract.image_to_data(
        rotada,
        lang="spa",
        config="--psm 6",
        output_type=pytesseract.Output.DICT
    )

    calidad, texto = calidad_ocr(datos)

    # Las caratulas son verticales: pequeno desempate a favor del
    # resultado en formato retrato.
    if rotada.height > rotada.width * 1.1:
        calidad += 5

    return calidad, texto


def detectar_orientacion(imagen, log):

    gris = preparar_ocr(imagen, LADO_ORIENTACION)

    angulo_osd, confianza_osd = orientacion_osd(gris)

    osd_confiable = (
        angulo_osd is not None
        and confianza_osd >= OSD_CONFIANZA_MINIMA
    )

    if angulo_osd is None:
        log("OSD: no disponible / sin texto suficiente.")
    else:
        log(
            f"OSD sugiere {angulo_osd} grados "
            f"(confianza {confianza_osd:.2f})"
        )

    # Si OSD es confiable, solo se verifica su angulo y el opuesto
    # (la confusion mas comun de OSD es 0 vs 180).
    if osd_confiable:
        orden = [angulo_osd, (angulo_osd + 180) % 360]
    else:
        orden = list(ANGULOS)

    resultados = {}

    def evaluar(angulos):

        for angulo in angulos:

            if angulo in resultados:
                continue

            calidad, texto = evaluar_orientacion(gris, angulo)

            if osd_confiable and angulo == angulo_osd:
                calidad += min(confianza_osd, 20) * 1.5

            resultados[angulo] = (calidad, texto)

            log(f"Orientacion {angulo} grados: calidad {calidad:.2f}")

    evaluar(orden)

    mejor = max(resultados, key=lambda a: resultados[a][0])

    if resultados[mejor][0] < CALIDAD_MINIMA and len(resultados) < 4:

        log("Calidad baja: evaluando los cuatro angulos.")

        evaluar(ANGULOS)

        mejor = max(resultados, key=lambda a: resultados[a][0])

    log(f"ROTACION SELECCIONADA: {mejor} grados")

    orientada = imagen.rotate(mejor, expand=True, fillcolor="white")

    return orientada, mejor, resultados[mejor][1]


# =============================================
# OCR COMPLETO CON PARADA TEMPRANA
# =============================================

def ocr_completo(imagen, textos_previos=()):

    gris = preparar_ocr(imagen)
    normal = normalizar_fondo(gris)

    ancho, alto = gris.size

    # Cabecera (numero de proceso en el recuadro), ampliada.
    cabecera = normal.crop((0, 0, ancho, max(1, int(alto * 0.22))))
    cabecera = cabecera.resize(
        (cabecera.width * 3 // 2, cabecera.height * 3 // 2),
        Image.Resampling.LANCZOS
    )

    # Zona de datos (Coactivado / Deudor principal).
    zona = normal.crop((0, int(alto * 0.30), ancho, int(alto * 0.80)))

    # Lecturas ordenadas de mayor a menor utilidad. Se generan bajo
    # demanda: si las primeras bastan no se calculan las demas.
    lecturas = (
        (lambda: normal, 6),
        (lambda: cabecera, 6),
        (lambda: normal, 4),
        (lambda: zona, 6),
        (lambda: restar_fondo(gris), 6),
        (lambda: gris, 6),
        (lambda: cabecera, 11),
        (lambda: restar_fondo(gris), 4),
        (lambda: mejorar_contraste(gris), 6),
        (lambda: gris, 4),
        (lambda: zona, 11),
    )

    textos = []

    nombres = Counter()
    procesos = Counter()

    for texto in textos_previos:
        nombres.update(extraer_coactivados(texto))
        procesos.update(extraer_procesos(texto))

    for obtener, modo in lecturas:

        texto = pytesseract.image_to_string(
            obtener(),
            lang="spa",
            config=f"--psm {modo}"
        )

        textos.append(texto)

        nombres.update(extraer_coactivados(texto))
        procesos.update(extraer_procesos(texto))

        # Coactivado y proceso confirmados dos veces: suficiente.
        if (
            len(textos) >= 3
            and nombres and nombres.most_common(1)[0][1] >= 2
            and procesos and procesos.most_common(1)[0][1] >= 2
        ):
            break

    return textos


# =============================================
# ROTAR Y LEER IMAGEN
# =============================================

def leer_imagen(data, log):

    textos = []
    paginas = []
    rotaciones = []

    with Image.open(io.BytesIO(data)) as imagen:

        for frame in ImageSequence.Iterator(imagen):

            rgb = ImageOps.exif_transpose(frame).convert("RGB")

            orientada, angulo, texto = detectar_orientacion(rgb, log)

            rotaciones.append(angulo)
            textos.append(texto)
            textos.extend(ocr_completo(orientada, [texto]))

            # Guardar una copia reducida: ahorra mucha memoria y es
            # mas que suficiente para una pagina a 200 dpi.
            paginas.append(
                escalar_lado_mayor(orientada, LADO_PDF, solo_reducir=True)
            )

    if not paginas:
        raise ValueError("Imagen sin paginas validas.")

    return textos, paginas, rotaciones


def analizar_ocr(textos):

    nombres = Counter()
    procesos = Counter()

    for texto in textos:
        nombres.update(extraer_coactivados(texto))
        procesos.update(extraer_procesos(texto))

    return nombres, procesos


# =============================================
# BUSCAR COINCIDENCIA DEL COACTIVADO
# =============================================
#
# Cada expediente recibe un puntaje con toda la evidencia disponible:
#   * Proceso: exacto (+2), con un digito mal leido (+1, solo si el
#     nombre tambien coincide), distinto (-0.5).
#   * Nombre: el mejor entre el campo Coactivado y la busqueda del
#     nombre del PDF en todo el texto (esta ultima vale un poco menos).
# Niveles:
#   3 = proceso exacto + nombre >= 90   -> PDF_unidos
#   2 = proceso exacto, o proceso casi exacto + nombre >= 85
#   1 = solo nombre >= 85               -> PDF_revisar (niveles 1 y 2)

SIMILITUD_MINIMA = 85


def evaluar_expediente(nombres, procesos, lineas, nombre_pdf, procesos_pdf):

    sim_campo = max(
        (similitud_nombres(n, nombre_pdf) for n in nombres), default=0.0
    )

    mejor_ocr = max(
        nombres, key=lambda n: similitud_nombres(n, nombre_pdf), default=""
    )

    sim_texto = buscar_en_texto(nombre_pdf, lineas) * 0.95

    similitud = max(sim_campo, sim_texto)

    proceso = "ninguno"

    if procesos and procesos_pdf:

        diferencia = min(
            diferencia_procesos(a, b) for a in procesos for b in procesos_pdf
        )

        proceso = (
            "exacto" if diferencia == 0
            else "cercano" if diferencia == 1
            else "distinto"
        )

    nombre_ok = similitud >= SIMILITUD_MINIMA

    if proceso == "exacto":
        puntaje = 2 + similitud / 100
        nivel = 3 if similitud >= 90 else 2
    elif proceso == "cercano" and nombre_ok:
        puntaje = 1 + similitud / 100
        nivel = 2
    elif proceso == "distinto" and similitud >= 95:
        puntaje = similitud / 100 - 0.5
        nivel = 1
    elif proceso in ("ninguno", "cercano") and nombre_ok:
        puntaje = similitud / 100
        nivel = 1
    else:
        return None

    return {
        "puntaje": puntaje,
        "nivel": nivel,
        "similitud": similitud,
        "proceso": proceso,
        "nombre_ocr": mejor_ocr if sim_campo >= sim_texto else "(texto) " + nombre_pdf,
    }


def buscar_coincidencia(nombres, procesos, pdfs, log, textos=()):

    if nombres:
        for nombre, frecuencia in nombres.most_common(5):
            log(f"COACTIVADO OCR: {nombre} (lecturas: {frecuencia})")
    else:
        log("Campo Coactivado no identificado: se usara proceso y texto.")

    if procesos:
        for proceso, frecuencia in procesos.most_common(3):
            log(f"PROCESO OCR: {proceso} (lecturas: {frecuencia})")
    else:
        log("Numero de proceso no identificado.")

    lineas = lineas_para_busqueda(textos)

    candidatos = []

    for indice, (clave, (etiqueta, docs)) in enumerate(pdfs):

        nombre_pdf = nombre_archivo(etiqueta)

        if not nombre_pdf:
            continue

        evaluacion = evaluar_expediente(
            nombres, set(procesos), lineas,
            nombre_pdf, extraer_procesos(etiqueta)
        )

        if evaluacion:
            evaluacion.update(
                indice=indice, etiqueta=etiqueta, nombre_pdf=nombre_pdf
            )
            candidatos.append(evaluacion)

    if not candidatos:
        estado = (
            "COACTIVADO NO IDENTIFICADO" if not nombres and not procesos
            else "SIN COINCIDENCIA"
        )
        return None, estado

    candidatos.sort(key=lambda c: c["puntaje"], reverse=True)

    for c in candidatos[:3]:
        log(
            f"CANDIDATO: {c['etiqueta']} | nombre {c['similitud']:.0f}% | "
            f"proceso {c['proceso']} | nivel {c['nivel']}"
        )

    mejor = candidatos[0]

    if len(candidatos) > 1:

        segundo = candidatos[1]

        # Dos expedientes con evidencia casi igual.
        if segundo["puntaje"] >= mejor["puntaje"] - 0.05:
            return None, "COINCIDENCIA AMBIGUA"

        # Gana solo por proceso, pero otro expediente tiene el nombre:
        # el proceso pudo leerse mal.
        if mejor["similitud"] < SIMILITUD_MINIMA and any(
            c["similitud"] >= 90 for c in candidatos[1:]
        ):
            return None, "COINCIDENCIA AMBIGUA"

    log(f"COINCIDENCIA ENCONTRADA: {mejor['etiqueta']} (nivel {mejor['nivel']})")

    return mejor, "COINCIDENCIA"


# =============================================
# PDF: DIMENSIONES Y AJUSTE DE PAGINAS
# =============================================

def dimensiones_pagina(pagina):

    ancho = float(pagina.cropbox.width) * float(pagina.user_unit)
    alto = float(pagina.cropbox.height) * float(pagina.user_unit)

    if pagina.rotation % 360 in (90, 270):
        ancho, alto = alto, ancho

    return ancho, alto


def dimensiones_pdf(data):

    lector = PdfReader(io.BytesIO(data))

    if not lector.pages:
        raise ValueError("PDF sin paginas.")

    ancho, alto = dimensiones_pagina(lector.pages[0])

    if ancho <= 0 or alto <= 0:
        raise ValueError("Dimensiones invalidas.")

    return ancho, alto


def imagen_a_pdf(imagen, ancho_pdf, alto_pdf, dpi=200):

    ancho_px = max(1, round(ancho_pdf / 72 * dpi))
    alto_px = max(1, round(alto_pdf / 72 * dpi))

    lienzo = Image.new("RGB", (ancho_px, alto_px), "white")

    ajustada = ImageOps.contain(
        imagen.convert("RGB"),
        (ancho_px, alto_px),
        Image.Resampling.LANCZOS
    )

    lienzo.paste(
        ajustada,
        (
            (ancho_px - ajustada.width) // 2,
            (alto_px - ajustada.height) // 2
        )
    )

    buffer = io.BytesIO()

    lienzo.save(buffer, format="PDF", resolution=dpi, quality=90)

    return buffer.getvalue()


def agregar_pagina_ajustada(escritor, pagina, ancho, alto):

    # La pagina debe pertenecer a un PdfWriter antes de modificarla.
    temporal = PdfWriter()
    pagina = temporal.add_page(pagina)

    if pagina.rotation != 0:
        pagina.transfer_rotation_to_content()

    caja = pagina.cropbox

    ancho_original = float(caja.width)
    alto_original = float(caja.height)

    if ancho_original <= 0 or alto_original <= 0:
        raise ValueError("Pagina con dimensiones invalidas.")

    escala = min(ancho / ancho_original, alto / alto_original)

    transformacion = (
        Transformation()
        .translate(tx=-float(caja.left), ty=-float(caja.bottom))
        .scale(sx=escala, sy=escala)
        .translate(
            tx=(ancho - ancho_original * escala) / 2,
            ty=(alto - alto_original * escala) / 2
        )
    )

    nueva = escritor.add_blank_page(width=ancho, height=alto)

    nueva.merge_transformed_page(pagina, transformacion, expand=False)


def pagina_compatible(pagina, ancho, alto):

    caja = pagina.cropbox
    media = pagina.mediabox

    return (
        pagina.rotation % 360 == 0
        and float(pagina.user_unit) == 1
        and abs(float(caja.width) - ancho) < 0.01
        and abs(float(caja.height) - alto) < 0.01
        and abs(float(caja.left)) < 0.01
        and abs(float(caja.bottom)) < 0.01
        and abs(float(media.width) - ancho) < 0.01
        and abs(float(media.height) - alto) < 0.01
        and abs(float(media.left)) < 0.01
        and abs(float(media.bottom)) < 0.01
    )


def unir_documentos(grupo, docs, log):

    if not docs:
        raise ValueError("No hay PDF originales.")

    escritor = PdfWriter()

    ancho, alto = dimensiones_pdf(docs[0][1])

    log(f"DIMENSIONES PDF ORIGINAL: {ancho:.2f} x {alto:.2f}")

    total_esperado = 0

    # PRIMERO: caratulas (imagenes ya orientadas).
    for registro in grupo:

        for imagen in registro["paginas"]:

            lector = PdfReader(io.BytesIO(imagen_a_pdf(imagen, ancho, alto)))

            for pagina in lector.pages:
                agregar_pagina_ajustada(escritor, pagina, ancho, alto)
                total_esperado += 1

    # SEGUNDO: PDF originales.
    for ruta, data in docs:

        for pagina in PdfReader(io.BytesIO(data)).pages:

            if pagina_compatible(pagina, ancho, alto):
                escritor.add_page(pagina)
            else:
                agregar_pagina_ajustada(escritor, pagina, ancho, alto)

            total_esperado += 1

    escritor.compress_identical_objects()

    resultado = io.BytesIO()
    escritor.write(resultado)
    datos_finales = resultado.getvalue()

    # VALIDAR RESULTADO
    verificacion = PdfReader(io.BytesIO(datos_finales))

    if len(verificacion.pages) != total_esperado:
        raise ValueError("Cantidad incorrecta de paginas.")

    for numero, pagina in enumerate(verificacion.pages, 1):

        ancho_pagina, alto_pagina = dimensiones_pagina(pagina)

        if abs(ancho_pagina - ancho) > 0.1 or abs(alto_pagina - alto) > 0.1:
            raise ValueError(f"Dimensiones incorrectas en pagina {numero}.")

    log(f"PDF VALIDADO: {len(verificacion.pages)} paginas.")

    return datos_finales


# =============================================
# PROCESAR UNA IMAGEN (SE EJECUTA EN UN HILO)
# =============================================

def procesar_imagen(data, pdfs):
    """
    Lee y asigna una imagen. Devuelve el registro con sus propias
    lineas de log para mostrarlas en orden, sin mezclar hilos.
    """

    lineas = []
    log = lineas.append

    registro = {
        "log": lineas,
        "angulos": [],
        "nombres": Counter(),
        "procesos": Counter(),
        "candidato": None,
        "paginas": [],
        "estado": "",
        "detalle": ""
    }

    try:

        textos, paginas, angulos = leer_imagen(data, log)

        nombres, procesos = analizar_ocr(textos)

        candidato, estado = buscar_coincidencia(
            nombres, procesos, pdfs, log, textos
        )

        registro.update({
            "angulos": angulos,
            "nombres": nombres,
            "procesos": procesos,
            "candidato": candidato,
            "estado": estado,
        })

        if candidato is not None:
            registro["paginas"] = paginas
        else:
            registro["detalle"] = "\n".join(textos)[:3000]

    except Exception as error:

        registro["estado"] = "ERROR OCR/IMAGEN"
        registro["detalle"] = str(error)

    return registro


def en_orden_paralelo(funcion, elementos, hilos):
    """
    Como executor.map, pero con una ventana limitada de tareas en curso
    para no cargar todas las imagenes en memoria a la vez.
    """

    with ThreadPoolExecutor(max_workers=hilos) as ejecutor:

        pendientes = []
        iterador = iter(elementos)

        for elemento in iterador:

            pendientes.append(ejecutor.submit(funcion, elemento))

            if len(pendientes) >= hilos * 2:
                yield pendientes.pop(0).result()

        for futuro in pendientes:
            yield futuro.result()


# =============================================
# PROCESAMIENTO GENERAL
# =============================================

def procesar(zimg, zpdf, salida, log, progreso, hilos=HILOS):

    filas = []
    rotadas = pendientes = generados = revisar = errores = 0

    with zipfile.ZipFile(zimg) as zi, \
         zipfile.ZipFile(zpdf) as zp, \
         zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zo:

        expedientes = {}

        for clave, etiqueta, ruta, data in documentos(zp):
            expedientes.setdefault(clave, [etiqueta, []])[1].append(
                (ruta, data)
            )

        pdfs = list(expedientes.items())
        imagenes = entradas(zi, EXTENSIONES_IMAGEN)

        if not pdfs:
            raise ValueError("No se encontraron archivos PDF.")

        if not imagenes:
            raise ValueError("No se encontraron imagenes.")

        log(f"TOTAL EXPEDIENTES: {len(pdfs)}")
        log(f"TOTAL IMAGENES: {len(imagenes)}")
        log(f"HILOS DE TRABAJO: {hilos}")

        grupos = {i: [] for i in range(len(pdfs))}

        # =====================================
        # PROCESAR FOTOGRAFIAS
        # =====================================

        # La lectura del ZIP se hace en el hilo principal (generador),
        # el OCR en los hilos de trabajo.
        def tareas():
            for imagen in imagenes:
                yield imagen, zi.read(imagen)

        def trabajo(tarea):
            imagen, data = tarea
            return imagen, data, procesar_imagen(data, pdfs)

        progreso(0, len(imagenes))

        for j, (imagen, data, registro) in enumerate(
            en_orden_paralelo(trabajo, tareas(), hilos), 1
        ):

            log("")
            log("=" * 55)
            log(f"IMAGEN {j}/{len(imagenes)}")
            log(imagen.filename)

            for linea in registro["log"]:
                log(linea)

            progreso(j, len(imagenes))

            angulos = registro["angulos"]

            if any(angulo != 0 for angulo in angulos):
                rotadas += 1

            candidato = registro["candidato"]

            if candidato is not None:

                grupos[candidato["indice"]].append({
                    "archivo": imagen.filename,
                    "original": data,
                    "paginas": registro["paginas"],
                    "rotaciones": angulos,
                    "nivel": candidato["nivel"],
                    "nombre_ocr": candidato["nombre_ocr"],
                    "procesos": list(registro["procesos"])
                })

                log("IMAGEN ASIGNADA CORRECTAMENTE.")

                continue

            pendientes += 1

            log(f"ESTADO: {registro['estado']}")

            zo.writestr(
                f"Pendientes_imagenes/{j:06d}_"
                f"{PurePosixPath(imagen.filename).name}",
                data
            )

            filas.append([
                imagen.filename,
                "",
                " | ".join(registro["nombres"]),
                " | ".join(registro["procesos"]),
                str(angulos),
                registro["estado"],
                registro["detalle"]
            ])

        # =====================================
        # GENERAR PDF FINALES
        # =====================================

        log("")
        log("=" * 55)
        log("GENERANDO PDF CONSOLIDADOS")
        log("=" * 55)

        for i, (clave, (etiqueta, docs)) in enumerate(pdfs):

            grupo = grupos[i]

            base = f"{i + 1:06d}_{PurePosixPath(etiqueta).stem}.pdf"

            if not grupo:

                for k, (ruta, data) in enumerate(docs):
                    zo.writestr(
                        f"PDF_sin_imagen/{i:06d}_{k:03d}_"
                        f"{PurePosixPath(ruta).name}",
                        data
                    )

                filas.append(["", clave, "", "", "", "PDF SIN IMAGEN", ""])

                continue

            try:

                log("")
                log(f"UNIENDO: {etiqueta}")

                resultado = unir_documentos(grupo, docs, log)

                requiere_revision = any(r["nivel"] != 3 for r in grupo)

                carpeta = "PDF_revisar/" if requiere_revision else "PDF_unidos/"

                zo.writestr(carpeta + base, resultado)

                generados += 1

                if requiere_revision:
                    revisar += 1

                log(f"PDF GENERADO: {carpeta}{base}")

                for registro in grupo:
                    filas.append([
                        registro["archivo"],
                        clave,
                        registro["nombre_ocr"],
                        " | ".join(registro["procesos"]),
                        str(registro["rotaciones"]),
                        "REVISAR OCR" if requiere_revision else "UNIDO",
                        carpeta + base
                    ])

            except Exception as error:

                errores += 1

                log(f"ERROR UNION: {error}")

                for k, (ruta, data) in enumerate(docs):
                    zo.writestr(
                        f"PDF_con_error/{i:06d}_{k:03d}_"
                        f"{PurePosixPath(ruta).name}",
                        data
                    )

                for k, registro in enumerate(grupo):

                    zo.writestr(
                        f"Pendientes_imagenes/error_{i}_{k}_"
                        f"{PurePosixPath(registro['archivo']).name}",
                        registro["original"]
                    )

                    filas.append([
                        registro["archivo"],
                        clave,
                        registro["nombre_ocr"],
                        " | ".join(registro["procesos"]),
                        str(registro["rotaciones"]),
                        "ERROR UNION",
                        str(error)
                    ])

            # Liberar memoria de las imagenes ya usadas.
            grupos[i] = []

        # =====================================
        # REPORTE CSV Y RESUMEN
        # =====================================

        reporte = io.StringIO()

        escritor_csv = csv.writer(reporte, delimiter=";")

        escritor_csv.writerow([
            "Imagen", "Expediente", "Coactivado OCR", "Proceso OCR",
            "Rotacion", "Estado", "Detalle"
        ])

        escritor_csv.writerows(filas)

        zo.writestr("reporte.csv", reporte.getvalue().encode("utf-8-sig"))

        resumen = (
            f"UNIFICADOR V{VERSION}\n"
            f"TOTAL IMAGENES: {len(imagenes)}\n"
            f"TOTAL EXPEDIENTES: {len(pdfs)}\n"
            f"IMAGENES ROTADAS: {rotadas}\n"
            f"PDF GENERADOS: {generados}\n"
            f"COINCIDENCIAS EXACTAS: {generados - revisar}\n"
            f"PDF PARA REVISION: {revisar}\n"
            f"IMAGENES PENDIENTES OCR: {pendientes}\n"
            f"ERRORES DE UNION: {errores}\n"
        )

        zo.writestr("RESUMEN.txt", resumen.encode("utf-8-sig"))

        log("")
        log("=" * 55)
        log(resumen)
        log("PROCESAMIENTO FINALIZADO.")

    return resumen


def verificar_tesseract(log):

    idiomas = pytesseract.get_languages(config="")

    if "spa" not in idiomas:
        raise RuntimeError("Instala el idioma español (spa) de Tesseract.")

    if "osd" not in idiomas:
        log(
            "AVISO: falta 'osd.traineddata'. La rotacion se detectara "
            "evaluando los 4 angulos (mas lento)."
        )


# =============================================
# INTERFAZ GRAFICA
# =============================================

def interfaz():

    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(f"Unificador de Expedientes V{VERSION}")
    root.geometry("1000x750")
    root.minsize(750, 550)
    root.configure(bg="#F0F4F8")

    tk.Label(
        root,
        text="UNIFICADOR DE EXPEDIENTES BANECUADOR",
        font=("Segoe UI", 19, "bold"),
        bg="#154360", fg="white", pady=18
    ).pack(fill="x")

    tk.Label(
        root,
        text=(
            f"Version {VERSION} - Coactivado + rotacion automatica "
            f"(OSD + verificacion OCR)"
        ),
        font=("Segoe UI", 11, "bold"),
        bg="#F0F4F8", fg="#148F77"
    ).pack(pady=10)

    variables = [tk.StringVar(), tk.StringVar()]

    titulos = [
        "1. Seleccionar ZIP de imagenes",
        "2. Seleccionar ZIP de documentos PDF"
    ]

    for i, titulo in enumerate(titulos):

        tk.Label(
            root, text=titulo,
            font=("Segoe UI", 11, "bold"), bg="#F0F4F8"
        ).pack(anchor="w", padx=20, pady=(10, 4))

        marco = tk.Frame(root, bg="#F0F4F8")
        marco.pack(fill="x", padx=20)

        tk.Entry(
            marco, textvariable=variables[i], font=("Segoe UI", 10)
        ).pack(side="left", fill="x", expand=True, ipady=6)

        def seleccionar(indice=i):

            archivo = filedialog.askopenfilename(
                filetypes=[("Archivos ZIP", "*.zip")]
            )

            if archivo:
                variables[indice].set(archivo)

        tk.Button(
            marco, text="Seleccionar ZIP",
            bg="#2874A6", fg="white", command=seleccionar
        ).pack(side="left", padx=8)

    # BARRA DE PROGRESO

    estado_progreso = tk.StringVar(value="Esperando archivos...")

    tk.Label(
        root, textvariable=estado_progreso,
        bg="#F0F4F8", font=("Segoe UI", 10, "bold")
    ).pack(anchor="w", padx=20, pady=(15, 4))

    barra = ttk.Progressbar(root, orient="horizontal", mode="determinate")
    barra.pack(fill="x", padx=20, pady=5)

    def progreso(actual, total):

        def actualizar():
            barra["maximum"] = max(total, 1)
            barra["value"] = actual
            estado_progreso.set(f"Imagenes procesadas: {actual} de {total}")

        root.after(0, actualizar)

    # CONSOLA DE RESULTADOS

    tk.Label(
        root, text="Registro detallado del procesamiento",
        font=("Segoe UI", 11, "bold"), bg="#F0F4F8"
    ).pack(anchor="w", padx=20, pady=(10, 5))

    marco_log = tk.Frame(root)
    marco_log.pack(fill="both", expand=True, padx=20, pady=5)

    scroll = tk.Scrollbar(marco_log)
    scroll.pack(side="right", fill="y")

    consola = tk.Text(
        marco_log, bg="#17202A", fg="#82E0AA",
        font=("Consolas", 10), yscrollcommand=scroll.set, wrap="word"
    )
    consola.pack(fill="both", expand=True)

    scroll.config(command=consola.yview)

    # El log se acumula en una cola y se vuelca cada 100 ms:
    # evita miles de llamadas a root.after que congelan la ventana.
    cola_log = []
    candado_log = threading.Lock()

    def log(texto):
        with candado_log:
            cola_log.append(str(texto))

    def volcar_log():

        with candado_log:
            lineas = cola_log[:]
            cola_log.clear()

        if lineas:
            consola.insert("end", "\n".join(lineas) + "\n")
            consola.see("end")

        root.after(100, volcar_log)

    volcar_log()

    # INICIAR PROGRAMA

    def iniciar():

        zimg = variables[0].get()
        zpdf = variables[1].get()

        if not all(Path(x).is_file() for x in (zimg, zpdf)):
            messagebox.showerror("Error", "Selecciona los dos archivos ZIP.")
            return

        if not all(zipfile.is_zipfile(x) for x in (zimg, zpdf)):
            messagebox.showerror("Error", "Uno de los ZIP no es valido.")
            return

        destino = filedialog.asksaveasfilename(
            defaultextension=".zip",
            initialfile=f"expedientes_unidos_V{VERSION}.zip",
            filetypes=[("Archivos ZIP", "*.zip")]
        )

        if not destino:
            return

        if Path(destino).resolve() in (
            Path(zimg).resolve(), Path(zpdf).resolve()
        ):
            messagebox.showerror(
                "Error", "No puedes sobrescribir los archivos originales."
            )
            return

        boton.config(state="disabled")
        consola.delete("1.0", "end")
        barra["value"] = 0

        def trabajo():

            temporal = None

            try:

                verificar_tesseract(log)

                log(f"UNIFICADOR V{VERSION}")
                log(f"Version pypdf: {pypdf.__version__}")
                log(f"Version Tesseract: {pytesseract.get_tesseract_version()}")
                log("OCR ESPAÑOL ACTIVADO.")
                log("ROTACION AUTOMATICA: OSD + VERIFICACION OCR.")
                log("RECONOCIMIENTO EXCLUSIVO DEL COACTIVADO.")

                with tempfile.NamedTemporaryFile(
                    suffix=".zip",
                    dir=str(Path(destino).parent),
                    delete=False
                ) as archivo_temporal:
                    temporal = archivo_temporal.name

                resumen = procesar(zimg, zpdf, temporal, log, progreso)

                os.replace(temporal, destino)
                temporal = None

                mensaje = resumen + "\nArchivo guardado en:\n" + destino

                root.after(0, lambda m=mensaje: messagebox.showinfo(
                    "Proceso finalizado", m
                ))

                root.after(0, lambda: estado_progreso.set(
                    "Procesamiento finalizado."
                ))

            except Exception as error:

                mensaje = str(error)

                log("ERROR: " + mensaje)

                root.after(0, lambda m=mensaje: messagebox.showerror(
                    "Error", m
                ))

            finally:

                if temporal and os.path.exists(temporal):
                    try:
                        os.remove(temporal)
                    except OSError:
                        pass

                root.after(0, lambda: boton.config(state="normal"))

        threading.Thread(target=trabajo, daemon=True).start()

    boton = tk.Button(
        root, text="PROCESAR Y UNIR PDF",
        font=("Segoe UI", 12, "bold"),
        bg="#148F77", fg="white", padx=35, pady=12,
        command=iniciar
    )

    boton.pack(pady=15)

    root.mainloop()


# =============================================
# EJECUTAR
# =============================================

def main():

    import sys

    # Modo consola (util para pruebas o tareas programadas):
    #   python unificador.py imagenes.zip pdfs.zip salida.zip
    if len(sys.argv) == 4:

        verificar_tesseract(print)

        procesar(
            sys.argv[1], sys.argv[2], sys.argv[3],
            print, lambda actual, total: None
        )

        return

    interfaz()


if __name__ == "__main__":
    main()
