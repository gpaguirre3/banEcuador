"""
Completa el Excel "Acta de entrega de expedientes" con los datos que faltan.

Entrada:
  * Excel (.xlsx) con una fila por expediente: Numero de juicio, Deudor
    principal, Cedula deudor principal y columnas vacias para Deudor 2,
    Garante 1..3 (cedula, edad, correo).
  * ZIP con los PDF de los expedientes (nombre del archivo con el numero
    de juicio, por ejemplo "ATI BARBERO DOLORES 080-2026-00106.pdf").

Proceso:
  1. En cada PDF se busca la hoja de firmas del pagare (primero la hoja 4,
     luego las demas): "FIRMA DE LOS DEUDORES" ... "En los mismos terminos
     ..." ... "FIRMA DE LOS GARANTES SOLIDARIOS". Se leen por OCR las
     cedulas de deudores y garantes (validadas con el digito verificador).
  2. Cada cedula se consulta en Mina: nombre (Registro Civil), edad y
     correos.
  3. Se llenan SOLO las celdas vacias del Excel y se agrega una hoja
     "Reporte" con el detalle de cada fila.

Uso por consola:
    python completar_excel.py acta.xlsx pdfs.zip salida.xlsx
"""

import io
import os
import re
import sys
import zipfile
import threading

from copy import copy
from pathlib import Path, PurePosixPath
from concurrent.futures import ThreadPoolExecutor

import pytesseract

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from pypdf import PdfReader

from unificador import (
    HILOS,
    documentos,
    extraer_procesos,
    nombre_archivo,
    normalizar,
    normalizar_fondo,
    preparar_ocr,
    quitar_tildes,
    similitud_nombres,
)


BASE = Path(__file__).resolve().parent

HOJA_PREFERIDA = 4          # hoja donde normalmente estan las firmas
MAX_HOJAS_BUSQUEDA = 15     # si no esta en la 4, se revisan las primeras N
MAX_DEUDORES = 3
MAX_GARANTES = 3

# Celdas completadas automaticamente: fondo verde claro para revisarlas.
RELLENO_COMPLETADO = PatternFill("solid", fgColor="E2F0D9")


# =============================================
# .ENV
# =============================================

def cargar_env(ruta=BASE / ".env"):
    """Lee un archivo .env sencillo (CLAVE=valor) sin dependencias."""

    ruta = Path(ruta)

    if not ruta.exists():
        return

    for linea in ruta.read_text(encoding="utf-8").splitlines():

        linea = linea.strip()

        if not linea or linea.startswith("#") or "=" not in linea:
            continue

        clave, valor = linea.split("=", 1)
        os.environ.setdefault(clave.strip(), valor.strip().strip('"').strip("'"))


# =============================================
# CEDULAS
# =============================================

CORRECCION_DIGITOS = str.maketrans({
    "O": "0", "o": "0", "Q": "0", "D": "0",
    "I": "1", "l": "1", "|": "1", "L": "1",
    "S": "5", "s": "5", "B": "8", "Z": "2",
})

PATRON_NUMERO = re.compile(r"(?<![0-9A-Za-z])[0-9OoQDIlL|SsBZ]{10,13}(?![0-9A-Za-z])")


def cedula_valida(cedula):
    """Cedula ecuatoriana de persona natural (modulo 10)."""

    if len(cedula) != 10 or not cedula.isdigit():
        return False

    provincia = int(cedula[:2])

    if not (1 <= provincia <= 24 or provincia == 30):
        return False

    if int(cedula[2]) >= 6:
        return False

    suma = 0

    for digito, coeficiente in zip(cedula[:9], (2, 1, 2, 1, 2, 1, 2, 1, 2)):
        valor = int(digito) * coeficiente
        suma += valor - 9 if valor > 9 else valor

    return (10 - suma % 10) % 10 == int(cedula[9])


def normalizar_cedula(valor):
    """Cedula desde una celda de Excel (numero sin 0 inicial, '0603...0', etc.)."""

    texto = str(valor or "").strip()

    if texto.endswith(".0"):
        texto = texto[:-2]

    texto = re.sub(r"\D", "", texto)

    return texto.zfill(10) if 9 <= len(texto) <= 10 else texto


def extraer_cedulas(texto):
    """Cedulas validas en orden de aparicion (acepta RUC 13 digitos ...001)."""

    encontradas = []

    for coincidencia in PATRON_NUMERO.finditer(texto):

        bruto = coincidencia.group(0)

        # Debe ser mayoritariamente numerico para no confundir palabras.
        if sum(c.isdigit() for c in bruto) < 7:
            continue

        numero = bruto.translate(CORRECCION_DIGITOS)

        if len(numero) == 13 and numero.endswith("001"):
            numero = numero[:10]

        if cedula_valida(numero) and numero not in encontradas:
            encontradas.append(numero)

    return encontradas


# =============================================
# HOJA DE FIRMAS DEL PAGARE
# =============================================

PATRON_PARRAFO = re.compile(r"MISMOS\s+T\w{0,2}RMINOS")
PATRON_FIN_PARRAFO = re.compile(r"EXCUS\w*|UT\s+SUPRA")
PATRON_SENORES = re.compile(
    r"SE\w{0,2}OR\s*\(?\s*E?S?\s*\)?\s*(.+?)\s+A\s+FAVOR", re.S
)
PATRON_NOMBRE = re.compile(r"NOMBRE\s*[:;.]?\s*(.+)")


def limpiar_persona(texto):
    """Palabras de un nombre: sin basura de firmas ni letras sueltas."""

    palabras = [p for p in normalizar(texto).split() if p.isalpha() and len(p) > 1]

    return " ".join(palabras[:6])


def analizar_hoja_firmas(texto):
    """
    Separa la hoja en deudores / parrafo / garantes.
    Devuelve None si no es la hoja de firmas del pagare.
    """

    t = quitar_tildes(texto)

    parrafo = PATRON_PARRAFO.search(t)

    if not parrafo:
        return None

    fin = PATRON_FIN_PARRAFO.search(t, parrafo.end())
    fin_parrafo = fin.end() if fin else parrafo.end()

    deudores_txt = t[:parrafo.start()]

    # Solo lo que esta despues de "FIRMA DE LOS DEUDORES" (si se leyo).
    firma = list(re.finditer(r"FIRMA\s+DE\s+LOS", deudores_txt))

    if firma:
        deudores_txt = deudores_txt[firma[-1].start():]

    garantes_txt = t[fin_parrafo:]

    # Nombres de los deudores: el parrafo los trae limpios
    # ("por el (los) senor (es) NOMBRE1, NOMBRE2 a favor de BanEcuador").
    nombres_deudores = []
    senores = PATRON_SENORES.search(t[parrafo.start():fin_parrafo + 200])

    if senores:
        for parte in senores.group(1).replace("\n", " ").split(","):
            nombre = limpiar_persona(parte)
            if len(nombre.split()) >= 2:
                nombres_deudores.append(nombre)

    nombres_garantes = [
        limpiar_persona(m.group(1))
        for m in PATRON_NOMBRE.finditer(garantes_txt)
    ]

    return {
        "cedulas_deudores": extraer_cedulas(deudores_txt)[:MAX_DEUDORES],
        "cedulas_garantes": extraer_cedulas(garantes_txt)[:MAX_GARANTES],
        "nombres_deudores": nombres_deudores,
        "nombres_garantes": [n for n in nombres_garantes if len(n.split()) >= 2],
    }


# ---------------------------------------------
# Fuente 2: "... con cedula N, en calidad de deudor / garante"
# (Declaratoria de plazo vencido, Orden de pago, autos del juicio)
# ---------------------------------------------

# Tolera basura del escaneo entre palabras: "EN CALIDAD | DE GARANTE".
PATRON_CALIDAD = re.compile(
    r"CALIDAD[\s|.,:;_\-]{1,6}DE[\s|.,:;_\-]{1,6}(CO\s*-?\s*)?(DEUDOR|GARANT)\w*", re.I
)

PATRON_NOMBRE_CEDULA = re.compile(
    r"([A-Z][A-Z ]{8,80}?)\s*,?\s*CON\s+C\w{0,2}DULA[^0-9]{0,45}?(\d{10})"
)


def analizar_calidad(texto):
    """
    Roles segun "en calidad de deudor(es) principal(es)" / "garante(s)
    solidario(s)" / "codeudor". Cada cedula toma el rol del primer
    "en calidad de" que la sigue.
    """

    t = quitar_tildes(texto)
    t = re.sub(r"[ \t]+", " ", t.replace("\n", " "))

    deudores, garantes, nombres = [], [], {}
    anterior = 0
    marcas = 0

    for marca in PATRON_CALIDAD.finditer(t):

        marcas += 1

        # Solo el texto cercano: evita cedulas de otros parrafos.
        inicio = max(anterior, marca.start() - 600)
        segmento = t[inicio:marca.start()]
        anterior = marca.end()

        destino = garantes if marca.group(2).upper().startswith("GARANT") else deudores

        for cedula in extraer_cedulas(segmento):
            if cedula not in deudores and cedula not in garantes:
                destino.append(cedula)

        for m in PATRON_NOMBRE_CEDULA.finditer(segmento):
            nombre = limpiar_persona(re.split(r"\b(?:Y|DE|A|LOS|SENOR\w*)\b\s*$", m.group(1))[0])
            nombre = limpiar_persona(nombre.split(",")[-1])
            if len(nombre.split()) >= 2:
                nombres.setdefault(m.group(2), nombre)

    if not deudores:
        return None

    return {
        "cedulas_deudores": deudores[:MAX_DEUDORES],
        "cedulas_garantes": garantes[:MAX_GARANTES],
        "nombres_deudores": [],
        "nombres_garantes": [],
        "nombres": nombres,
        "marcas": marcas,
    }


def analizar_solo_deudores(texto):
    """Pagare sin garantes: solo "FIRMA DE LOS DEUDORES" + cedulas."""

    t = quitar_tildes(texto)
    firma = list(re.finditer(r"FIRMA\s+DE\s+LOS\s+DEUDOR", t))

    if not firma:
        return None

    cedulas = extraer_cedulas(t[firma[-1].start():])

    if not cedulas:
        return None

    return {
        "cedulas_deudores": cedulas[:MAX_DEUDORES],
        "cedulas_garantes": [],
        "nombres_deudores": [],
        "nombres_garantes": [],
    }


def imagen_pagina(pagina):
    """Imagen mas grande de la pagina (los expedientes son escaneados)."""

    imagenes = list(pagina.images)

    if not imagenes:
        return None

    mayor = max(imagenes, key=lambda i: i.image.width * i.image.height)

    return mayor.image.convert("RGB")


def texto_pagina(pagina, modo=6, variante="normal", lado=None):

    texto = (pagina.extract_text() or "").strip()

    if len(texto) > 200:   # PDF con texto digital: no hace falta OCR
        return texto

    imagen = imagen_pagina(pagina)

    if imagen is None:
        return texto

    gris = preparar_ocr(imagen, lado) if lado else preparar_ocr(imagen)
    imagen_ocr = normalizar_fondo(gris) if variante == "normal" else gris

    return pytesseract.image_to_string(imagen_ocr, lang="spa", config=f"--psm {modo}")


def unir_listas(*listas, maximo):

    resultado = []

    for lista in listas:
        for elemento in lista:
            if elemento not in resultado:
                resultado.append(elemento)

    return resultado[:maximo]


def releer(pagina, analizar, datos):
    """Segunda lectura OCR de la pagina elegida para recuperar cedulas."""

    extra = analizar(texto_pagina(pagina, modo=4, variante="gris"))

    if not extra:
        return datos

    # La primera lectura manda: la segunda solo agrega cedulas nuevas,
    # nunca cambia el rol de una ya leida.
    conocidas = set(datos["cedulas_deudores"]) | set(datos["cedulas_garantes"])

    for clave, maximo in (("cedulas_deudores", MAX_DEUDORES), ("cedulas_garantes", MAX_GARANTES)):
        nuevas = [c for c in extra[clave] if c not in conocidas]
        datos[clave] = unir_listas(datos[clave], nuevas, maximo=maximo)
        conocidas |= set(nuevas)

    for clave in ("nombres_deudores", "nombres_garantes"):
        if len(extra.get(clave, [])) > len(datos.get(clave, [])):
            datos[clave] = extra[clave]

    datos.setdefault("nombres", {})
    for cedula, nombre in extra.get("nombres", {}).items():
        datos["nombres"].setdefault(cedula, nombre)

    return datos


def leer_firmas_pdf(data):
    """
    Busca deudores y garantes en el expediente. Devuelve
    (numero_de_hoja, datos) o (None, None).

    Orden de busqueda (se detiene en la primera fuente completa):
      1. Pagare con deudores y garantes ("En los mismos terminos..."),
         empezando por la hoja 4.
      2. Hojas con "... con cedula N, en calidad de deudor / garante"
         (declaratoria de plazo vencido, orden de pago).
      3. Pagare solo con deudores (sin garantes).
    """

    lector = PdfReader(io.BytesIO(data))
    total = len(lector.pages)

    orden = [HOJA_PREFERIDA - 1] + [i for i in range(total) if i != HOJA_PREFERIDA - 1]
    solo_deudores = None

    for indice in orden:

        if indice >= total:
            continue

        pagina = lector.pages[indice]
        texto = texto_pagina(pagina, lado=1800)

        # 1. Pagare completo
        if indice < MAX_HOJAS_BUSQUEDA:
            datos = analizar_hoja_firmas(texto)
            if datos and (datos["cedulas_deudores"] or datos["cedulas_garantes"]):
                datos = releer(pagina, analizar_hoja_firmas, analizar_hoja_firmas(texto_pagina(pagina)) or datos)
                datos["fuente"] = "pagare"
                break

        # 2. "en calidad de deudor / garante"
        datos = analizar_calidad(texto)
        if datos:
            # De las dos lecturas manda la que encontro mas marcas de rol
            # (si el OCR pierde un "en calidad de garante", sus cedulas
            # quedarian pegadas al rol siguiente).
            completa = analizar_calidad(texto_pagina(pagina))
            if completa and completa["marcas"] > datos["marcas"]:
                datos = completa
            datos = releer(pagina, analizar_calidad, datos)
            datos["fuente"] = "en calidad de"
            break

        # 3. Pagare sin garantes: se guarda por si no aparece nada mejor.
        if solo_deudores is None:
            encontrado = analizar_solo_deudores(texto)
            if encontrado:
                solo_deudores = (indice, encontrado)

    else:
        if not solo_deudores:
            return None, None
        indice, datos = solo_deudores
        datos["fuente"] = "pagare (sin garantes)"

    # Una cedula no puede ser deudor y garante a la vez.
    datos["cedulas_garantes"] = [
        c for c in datos["cedulas_garantes"] if c not in datos["cedulas_deudores"]
    ]
    datos.setdefault("nombres", {})

    return indice + 1, datos


# =============================================
# MINA
# =============================================

LEER_PERSONA = r"""(cedula) => {
  const norm = t => (t || "").normalize("NFD").replace(/[̀-ͯ]/g, "")
                    .trim().replace(/:$/, "").trim().toLowerCase();
  const d1 = document.querySelector(".field[data-name='datos']");
  const d2 = document.querySelector(".field[data-name='datos2']");
  const valorEn = (campo, rotulo) => {
    if (!campo) return "";
    for (const c of campo.querySelectorAll("th, td"))
      if (norm(c.innerText) === rotulo && c.nextElementSibling?.tagName === "TD") {
        const v = c.nextElementSibling.innerText.trim();
        if (v) return v;
      }
    return "";
  };
  const esCorreo = t => /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(t.trim());
  const correos = [];
  const agregar = t => {
    for (const parte of (t || "").split(/[\s;,]+/))
      if (esCorreo(parte) && !correos.includes(parte.toLowerCase())) correos.push(parte.toLowerCase());
  };

  const nombre = valorEn(d1, "nombre completo") || valorEn(d2, "nombres");
  let edad = valorEn(d1, "edad") || valorEn(d2, "edad");

  for (const tabla of d1 ? d1.querySelectorAll("table") : []) {
    if (!tabla.rows.length) continue;
    const cab = [...tabla.rows[0].cells].map(c => norm(c.innerText));
    // "Telefonos y Correos": columna EMAIL
    const iEmail = cab.findIndex(c => c === "email" || c.startsWith("correo"));
    const iCedula = cab.findIndex(c => c === "cedula");
    const iCorreosFam = cab.findIndex(c => c.startsWith("correos"));
    const iEdad = cab.findIndex(c => c === "edad");
    for (const fila of [...tabla.rows].slice(1)) {
      const celdas = [...fila.cells].map(c => c.innerText.trim());
      if (iCedula >= 0) {
        // Arbol genealogico: solo la fila de la persona consultada
        if (celdas[iCedula] !== cedula) continue;
        if (iCorreosFam >= 0) agregar(celdas[iCorreosFam]);
        if (!edad && iEdad >= 0) edad = celdas[iEdad];
      } else if (iEmail >= 0) {
        agregar(celdas[iEmail]);
      }
    }
  }

  // "Datos 2": filas "correo | NOMBRE" a nombre de la persona
  const clave = norm(nombre).split(/\s+/).slice(0, 2).join(" ");
  for (const tabla of d2 ? d2.querySelectorAll("table") : []) {
    for (const fila of tabla.rows) {
      const celdas = [...fila.cells].map(c => c.innerText.trim());
      if (celdas.length >= 2 && esCorreo(celdas[0]) && (!clave || norm(celdas[1]).startsWith(clave)))
        agregar(celdas[0]);
    }
  }

  return { nombre, edad: /^\d+$/.test(edad || "") ? parseInt(edad) : null, correos,
           listo: !!(d1 && d1.innerText.trim().length > 20) };
}"""


class Mina:
    """Sesion de Mina con Playwright (mismo flujo que el proyecto sacardatos)."""

    def __init__(self, log, espera=60):

        self.log = log
        self.espera = espera
        self.base = os.environ.get("MINA_URL") or "http://186.4.213.167/Mina/"

        if not self.base.endswith("/"):
            self.base += "/"

        self.usuario = os.environ.get("MINA_USUARIO")
        self.clave = os.environ.get("MINA_CLAVE")

        if not self.usuario or not self.clave:
            raise RuntimeError("Faltan MINA_USUARIO y MINA_CLAVE en el archivo .env.")

        if os.environ.get("MINA_AUTORIZADO", "").lower() not in ("si", "sí", "true", "1"):
            raise RuntimeError("Falta MINA_AUTORIZADO=si en el archivo .env.")

    def __enter__(self):

        from playwright.sync_api import sync_playwright

        self._pw = sync_playwright().start()
        self._navegador = self._pw.chromium.launch(headless=True)
        self.page = self._navegador.new_page(viewport={"width": 1280, "height": 900})
        self.iniciar_sesion()

        return self

    def __exit__(self, *_):

        try:
            self._navegador.close()
        finally:
            self._pw.stop()

    def iniciar_sesion(self):

        page = self.page
        page.goto(self.base, wait_until="networkidle", timeout=60000)

        if page.locator("#btn-login").count():
            page.fill("#field-userName", self.usuario)
            page.fill("#field-password", self.clave)
            page.click("#btn-login")

        try:
            page.locator("#btn-login").wait_for(state="detached", timeout=30000)
        except Exception:
            raise RuntimeError("No se pudo iniciar sesion en Mina: revisa usuario y clave.")

    def consultar(self, cedula):
        """Devuelve {'nombre', 'edad', 'correos'} o lanza un error."""

        page = self.page
        crear = self.base + "#Consulta/create"

        page.goto("about:blank")
        page.goto(crear, wait_until="networkidle", timeout=60000)

        if page.locator("#btn-login").count():   # la sesion expiro
            self.iniciar_sesion()
            page.goto(crear, wait_until="networkidle", timeout=60000)

        campo = page.locator("input[data-name='cedula']").first
        campo.wait_for(state="visible", timeout=60000)
        campo.fill(cedula)
        page.locator("button[data-action='save']").first.click()

        try:
            page.wait_for_url(re.compile(r"#Consulta/view/"), timeout=self.espera * 1000)
        except Exception:
            raise RuntimeError("Mina no guardo la consulta")

        datos = None

        for intento in range(max(2, self.espera // 5)):

            page.locator(".field[data-name='datos']").first.wait_for(
                state="attached", timeout=60000
            )
            page.wait_for_timeout(1500)

            datos = page.evaluate(LEER_PERSONA, cedula)

            # Mina llena las secciones poco a poco: esperar nombre y edad.
            if datos["nombre"] and datos["edad"] is not None:
                break

            if intento >= 3 and datos["listo"] and datos["nombre"]:
                break

            page.wait_for_timeout(5000)
            page.reload(wait_until="networkidle")

        page.wait_for_timeout(1000)   # pausa para no saturar el servidor

        if not datos or not datos["nombre"]:
            raise RuntimeError("la consulta no devolvio datos")

        return {
            "nombre": normalizar(datos["nombre"]),
            "edad": datos["edad"],
            "correos": datos["correos"],
        }


# =============================================
# EXCEL
# =============================================

def texto_celda(valor):
    return normalizar(str(valor or ""))


def localizar_columnas(hoja):
    """
    Busca la fila de encabezados y arma los grupos de columnas:
      principal / deudor2 / deudor3 / garante1..3 -> {nombre, cedula, edad, correo}
    Los encabezados repetidos ("EDAD 2" aparece en Deudor 2 y en Garante 2)
    se resuelven por posicion: cada columna pertenece al grupo anterior.
    """

    for fila in range(1, min(hoja.max_row, 20) + 1):

        encabezados = {
            col: texto_celda(hoja.cell(fila, col).value)
            for col in range(1, hoja.max_column + 1)
        }

        if not any("JUICIO" in h for h in encabezados.values()):
            continue

        columnas = {"fila": fila, "grupos": {}}
        grupo = None

        for col, h in encabezados.items():

            if not h:
                continue

            if "JUICIO" in h:
                columnas["juicio"] = col
                continue

            if h == "DEUDOR PRINCIPAL":
                grupo = "principal"
            elif re.fullmatch(r"(?:CO ?)?DEUDOR \d", h):
                grupo = "deudor" + h[-1]
            elif re.fullmatch(r"GARANTE \d", h):
                grupo = "garante" + h[-1]
            elif grupo is None:
                continue
            else:
                campo = (
                    "cedula" if "CEDULA" in h
                    else "edad" if "EDAD" in h
                    else "correo" if "CORREO" in h or "EMAIL" in h
                    else None
                )
                if campo:
                    columnas["grupos"].setdefault(grupo, {})[campo] = col
                continue

            columnas["grupos"].setdefault(grupo, {})["nombre"] = col

        if "juicio" in columnas and "principal" in columnas["grupos"]:
            return columnas

    raise ValueError(
        "No se encontro la fila de encabezados (NUMERO DE JUICIO / DEUDOR PRINCIPAL)."
    )


def leer_filas(hoja, columnas):

    filas = []
    principal = columnas["grupos"]["principal"]

    for fila in range(columnas["fila"] + 1, hoja.max_row + 1):

        juicio = hoja.cell(fila, columnas["juicio"]).value
        nombre = hoja.cell(fila, principal["nombre"]).value

        if not juicio and not nombre:
            continue

        procesos = extraer_procesos(str(juicio or ""))

        filas.append({
            "fila": fila,
            "juicio": str(juicio or "").strip(),
            "proceso": next(iter(procesos), ""),
            "deudor": texto_celda(nombre),
            "cedula": normalizar_cedula(
                hoja.cell(fila, principal["cedula"]).value if "cedula" in principal else ""
            ),
        })

    return filas


def escribir(hoja, fila, columna, valor, como_texto=False):
    """Escribe solo si la celda esta vacia. Devuelve True si escribio."""

    if columna is None or valor in (None, "", []):
        return False

    celda = hoja.cell(fila, columna)

    if str(celda.value or "").strip():
        return False

    celda.value = str(valor) if como_texto else valor

    if como_texto:
        celda.number_format = "@"

    celda.fill = RELLENO_COMPLETADO

    return True


def columnas_estandar(n_deudores=2, n_garantes=3):
    """Encabezados en el mismo orden que la hoja modelo del acta."""

    encabezados = ["CORREO DEUDOR PRINCIPAL"]

    for n in range(2, n_deudores + 1):
        encabezados += [f"DEUDOR {n}", f"CEDULA DEUDOR {n}", f"EDAD {n}", f"CORREO DEUDOR {n}"]

    for n in range(1, n_garantes + 1):
        encabezados += [f"GARANTE {n}", f"CEDULA GARANTE {n}", f"EDAD {n}", f"CORREO G{n}"]

    return encabezados


def asegurar_columnas(hoja, columnas, log):
    """
    Si la hoja no tiene las columnas de Deudor 2 / Garantes, las agrega al
    final copiando el estilo del ultimo encabezado. Devuelve las columnas
    actualizadas.
    """

    grupos = columnas["grupos"]
    faltan = []

    if "correo" not in grupos["principal"]:
        faltan.append("CORREO DEUDOR PRINCIPAL")

    estandar = columnas_estandar()

    for encabezado in estandar[1:]:
        grupo = (
            "deudor" + encabezado[-1] if encabezado.startswith("DEUDOR ")
            else "garante" + encabezado[-1] if encabezado.startswith("GARANTE ")
            else None
        )
        if grupo and grupo in grupos:
            continue
        if grupo:
            indice = estandar.index(encabezado)
            faltan += estandar[indice:indice + 4]

    if not faltan:
        return columnas

    fila = columnas["fila"]
    ultima = max(
        (c for c in range(1, hoja.max_column + 1) if hoja.cell(fila, c).value not in (None, "")),
        default=1,
    )
    modelo = hoja.cell(fila, ultima)

    for desplazamiento, encabezado in enumerate(faltan, 1):

        celda = hoja.cell(fila, ultima + desplazamiento, encabezado)
        celda.font = copy(modelo.font)
        celda.fill = copy(modelo.fill)
        celda.border = copy(modelo.border)
        celda.alignment = copy(modelo.alignment)

        letra = celda.column_letter
        hoja.column_dimensions[letra].width = 38 if not re.search(r"CEDULA|EDAD", encabezado) else 14

    log(f"HOJA '{hoja.title}': se agregaron {len(faltan)} columnas ({faltan[0]} ... {faltan[-1]})")

    return localizar_columnas(hoja)


# =============================================
# PROCESO COMPLETO
# =============================================

def asignar_pdfs(filas, pdfs, log):
    """Relaciona cada PDF con su fila del Excel (por numero de juicio o nombre)."""

    por_proceso = {}

    for fila in filas:
        if fila["proceso"]:
            por_proceso.setdefault(fila["proceso"], []).append(fila)

    asignados = {}
    sin_fila = []

    for ruta, data in pdfs:

        nombre_pdf = PurePosixPath(ruta).name
        procesos = extraer_procesos(nombre_pdf)
        destino = [f for p in procesos for f in por_proceso.get(p, [])]

        if not destino:
            nombre = nombre_archivo(nombre_pdf)
            destino = [f for f in filas if similitud_nombres(nombre, f["deudor"]) >= 95]

        if len(destino) > 1:
            # El mismo juicio en varias hojas: se prefiere la fila cuyo
            # deudor coincide con el nombre del PDF.
            nombre = nombre_archivo(nombre_pdf)
            mejores = [f for f in destino if similitud_nombres(nombre, f["deudor"]) >= 90]
            if len(mejores) == 1:
                destino = mejores

        if len(destino) != 1:
            motivo = "no esta en el Excel" if not destino else (
                "aparece en varias filas: " + ", ".join(
                    f"{f['hoja'].title} fila {f['fila']}" for f in destino
                )
            )
            log(f"PDF SIN FILA EN EL EXCEL: {nombre_pdf} ({motivo})")
            sin_fila.append(nombre_pdf)
            continue

        fila = destino[0]
        aviso = ""

        if procesos and fila["proceso"] not in procesos:
            aviso = (
                f"asignado por nombre: el juicio del PDF ({', '.join(sorted(procesos))}) "
                f"no coincide con el del Excel ({fila['juicio']})"
            )
            fila.setdefault("avisos", []).append(aviso)

        log(
            f"{nombre_pdf} -> hoja '{fila['hoja'].title}' fila {fila['fila']}"
            + (f"  [AVISO: {aviso}]" if aviso else "")
        )
        asignados.setdefault(fila["clave"], []).append((ruta, data))

    return asignados, sin_fila


def procesar(ruta_excel, ruta_zip, salida, log=print, progreso=lambda *a, **k: None,
             usar_mina=True):

    libro = load_workbook(ruta_excel)

    # Todas las hojas con encabezados del acta (una por abogado, etc.).
    filas = []

    for hoja in libro.worksheets:

        if hoja.title == "Reporte":
            continue

        try:
            columnas = localizar_columnas(hoja)
        except ValueError:
            continue

        filas_hoja = leer_filas(hoja, columnas)

        for fila in filas_hoja:
            fila.update(hoja=hoja, columnas=columnas, clave=(hoja.title, fila["fila"]))

        filas += filas_hoja

        log(
            f"HOJA '{hoja.title}': encabezados en fila {columnas['fila']}, "
            f"{len(filas_hoja)} expedientes, columnas: "
            + ", ".join(columnas["grupos"])
        )

    if not filas:
        raise ValueError(
            "El Excel no tiene la fila de encabezados (NUMERO DE JUICIO / DEUDOR PRINCIPAL)."
        )

    por_clave = {fila["clave"]: fila for fila in filas}

    log(f"EXPEDIENTES EN EL EXCEL: {len(filas)}")

    with zipfile.ZipFile(ruta_zip) as zp:
        pdfs = [(ruta, data) for _, _, ruta, data in documentos(zp)]

    log(f"PDF EN EL ZIP: {len(pdfs)}")
    log("")

    asignados, sin_fila = asignar_pdfs(filas, pdfs, log)

    # Hojas que necesitan columnas nuevas (solo las que tienen PDF).
    for hoja in {por_clave[clave]["hoja"] for clave in asignados}:

        columnas = asegurar_columnas(hoja, next(
            f["columnas"] for f in filas if f["hoja"] is hoja
        ), log)

        for fila in filas:
            if fila["hoja"] is hoja:
                fila["columnas"] = columnas

    # ---------------- 1. OCR de la hoja de firmas ----------------

    trabajos = [(clave, ruta, data) for clave, lista in asignados.items() for ruta, data in lista]
    lecturas = {}
    candado = threading.Lock()
    hechos = [0]

    def leer(trabajo):

        clave, ruta, data = trabajo

        try:
            hoja_firmas, datos = leer_firmas_pdf(data)
        except Exception as error:
            hoja_firmas, datos = None, {"error": str(error)}

        with candado:
            hechos[0] += 1
            progreso(hechos[0], len(trabajos), "Leyendo hojas de firmas (OCR)")

        return clave, ruta, hoja_firmas, datos

    log("")
    progreso(0, len(trabajos), "Leyendo hojas de firmas (OCR)")

    with ThreadPoolExecutor(max_workers=HILOS) as ejecutor:

        for clave, ruta, hoja_firmas, datos in ejecutor.map(leer, trabajos):

            nombre = PurePosixPath(ruta).name

            if hoja_firmas is None:
                detalle = (datos or {}).get("error", "no se encontro la hoja de firmas")
                log(f"{nombre}: {detalle}")
                continue

            log(
                f"{nombre}: hoja {hoja_firmas} ({datos['fuente']}) | deudores "
                f"{datos['cedulas_deudores']} | garantes {datos['cedulas_garantes']}"
            )

            # Si una fila tiene varios PDF, se queda con el que trae mas datos.
            anterior = lecturas.get(clave)
            total = len(datos["cedulas_deudores"]) + len(datos["cedulas_garantes"])

            if not anterior or total > anterior["total"]:
                lecturas[clave] = {**datos, "hoja": hoja_firmas, "pdf": nombre, "total": total}

    # ---------------- 2. Personas por fila ----------------

    def nombre_ocr(lectura, lista_nombres, lista_cedulas, cedula):
        if cedula in lectura.get("nombres", {}):
            return lectura["nombres"][cedula]
        if len(lista_nombres) == len(lista_cedulas) and cedula in lista_cedulas:
            return lista_nombres[lista_cedulas.index(cedula)]
        return ""

    personas_por_fila = {}

    for clave, lectura in lecturas.items():

        fila = por_clave[clave]
        deudores = list(lectura["cedulas_deudores"])
        nombres_deudores = lectura["nombres_deudores"]

        # El deudor principal es el del Excel; los demas son Deudor 2, 3...
        principal = fila["cedula"] if cedula_valida(fila["cedula"]) else None

        if not principal and deudores:
            principal = deudores[0]
            for i, nombre in enumerate(nombres_deudores[:len(deudores)]):
                if similitud_nombres(nombre, fila["deudor"]) >= 90:
                    principal = deudores[i]

        personas = {"principal": {"cedula": principal, "nombre_ocr": fila["deudor"]}}

        for i, cedula in enumerate([c for c in deudores if c != principal], 2):
            personas[f"deudor{i}"] = {
                "cedula": cedula,
                "nombre_ocr": nombre_ocr(lectura, nombres_deudores, deudores, cedula),
            }

        for i, cedula in enumerate([c for c in lectura["cedulas_garantes"] if c != principal], 1):
            personas[f"garante{i}"] = {
                "cedula": cedula,
                "nombre_ocr": nombre_ocr(
                    lectura, lectura["nombres_garantes"], lectura["cedulas_garantes"], cedula
                ),
            }

        personas_por_fila[clave] = personas

    # ---------------- 3. Mina ----------------

    a_consultar = []

    for clave, personas in personas_por_fila.items():
        grupos = por_clave[clave]["columnas"]["grupos"]
        for grupo, persona in personas.items():
            if persona["cedula"] and grupo in grupos and persona["cedula"] not in a_consultar:
                a_consultar.append(persona["cedula"])

    mina = {}
    errores_mina = {}

    if usar_mina and a_consultar:

        log("")
        log(f"CONSULTANDO MINA: {len(a_consultar)} cedulas (unos 20 s cada una)")
        progreso(0, len(a_consultar), "Consultando Mina")

        try:

            with Mina(log) as sesion:

                for i, cedula in enumerate(a_consultar, 1):

                    for _ in range(2):
                        try:
                            mina[cedula] = sesion.consultar(cedula)
                            break
                        except Exception as error:
                            errores_mina[cedula] = str(error).splitlines()[0]

                    if cedula in mina:
                        errores_mina.pop(cedula, None)
                        d = mina[cedula]
                        log(
                            f"MINA {cedula}: {d['nombre']} | edad {d['edad']} | "
                            f"correos {len(d['correos'])}"
                        )
                    else:
                        log(f"MINA {cedula}: ERROR {errores_mina[cedula]}")

                    progreso(i, len(a_consultar), "Consultando Mina")

        except Exception as error:
            log(f"MINA NO DISPONIBLE: {error}")
            log("Se completaran solo los datos leidos de los PDF.")

    # ---------------- 4. Escribir el Excel ----------------

    reporte = []
    completadas = 0
    sin_firmas = 0

    for clave in asignados:

        fila = por_clave[clave]
        hoja = fila["hoja"]
        grupos = fila["columnas"]["grupos"]
        numero = fila["fila"]
        personas = personas_por_fila.get(clave)
        lectura = lecturas.get(clave)

        if not lectura:
            sin_firmas += 1
            reporte.append([
                hoja.title, numero, fila["juicio"], fila["deudor"],
                "SIN HOJA DE FIRMAS", ", ".join(PurePosixPath(r).name for r, _ in asignados[clave]),
                "", "", "",
            ])
            continue

        escritas = 0

        for grupo, persona in personas.items():

            columnas_grupo = grupos.get(grupo)
            cedula = persona["cedula"]

            if not columnas_grupo or not cedula:
                continue

            datos_mina = mina.get(cedula, {})
            nombre = datos_mina.get("nombre") or persona["nombre_ocr"]

            if grupo != "principal":
                escritas += escribir(hoja, numero, columnas_grupo.get("nombre"), nombre)

            escritas += escribir(hoja, numero, columnas_grupo.get("cedula"), cedula, como_texto=True)
            escritas += escribir(hoja, numero, columnas_grupo.get("edad"), datos_mina.get("edad"))
            escritas += escribir(
                hoja, numero, columnas_grupo.get("correo"),
                "; ".join(datos_mina.get("correos", []))
            )

        if escritas:
            completadas += 1

        sin_columna = [g for g in personas if g not in grupos]
        detalle_mina = "; ".join(
            f"{p['cedula']}: {errores_mina[p['cedula']]}"
            for p in personas.values() if p["cedula"] in errores_mina
        )

        reporte.append([
            hoja.title, numero, fila["juicio"], fila["deudor"],
            ("REVISAR: " if fila.get("avisos") else "")
            + ("COMPLETADO" if escritas else "SIN CAMBIOS"),
            lectura["pdf"], f"{lectura['hoja']} ({lectura['fuente']})",
            " | ".join(f"{g}: {p['cedula']}" for g, p in personas.items() if p["cedula"]),
            "; ".join(filter(None, [
                *fila.get("avisos", []),
                detalle_mina,
                f"sin columnas para {', '.join(sin_columna)}" if sin_columna else "",
            ])),
        ])

    for nombre_pdf in sin_fila:
        reporte.append(["", "", "", "", "PDF SIN FILA EN EL EXCEL", nombre_pdf, "", "", ""])

    # Hoja "Reporte"
    if "Reporte" in libro.sheetnames:
        del libro["Reporte"]

    hoja_reporte = libro.create_sheet("Reporte")
    hoja_reporte.append([
        "Hoja", "Fila", "Numero de juicio", "Deudor principal", "Estado", "PDF",
        "Hoja usada", "Cedulas encontradas", "Observaciones",
    ])

    for celda in hoja_reporte[1]:
        celda.font = Font(bold=True)

    for linea in reporte:
        hoja_reporte.append(linea)

    for letra, ancho in zip("ABCDEFGHI", (22, 7, 18, 38, 24, 50, 14, 70, 60)):
        hoja_reporte.column_dimensions[letra].width = ancho

    hoja_reporte.freeze_panes = "A2"

    libro.save(salida)

    resumen = (
        f"PDF RECIBIDOS: {len(pdfs)}\n"
        f"PDF SIN FILA EN EL EXCEL: {len(sin_fila)}\n"
        f"EXPEDIENTES COMPLETADOS: {completadas}\n"
        f"SIN HOJA DE FIRMAS: {sin_firmas}\n"
        f"CEDULAS CONSULTADAS EN MINA: {len(mina)}\n"
        f"ERRORES DE MINA: {len(errores_mina)}\n"
    )

    log("")
    log(resumen)

    return resumen


def main():

    if len(sys.argv) != 4:
        sys.exit("Uso: python completar_excel.py acta.xlsx pdfs.zip salida.xlsx")

    cargar_env()

    procesar(sys.argv[1], sys.argv[2], sys.argv[3])


if __name__ == "__main__":
    main()
