"""
Servidor web del Unificador de Expedientes.

Uso:
    python servidor.py            (puerto 5000)
    python servidor.py 9000       (otro puerto)

Abrir en el navegador: http://localhost:5000
Solo usa la libreria estandar de Python (ademas de las dependencias
del unificador).
"""

import json
import shutil
import sys
import threading
import traceback
import uuid
import zipfile

from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import unificador
import completar_excel


BASE = Path(__file__).resolve().parent
CARPETA_TRABAJOS = BASE / "trabajos"
PAGINA = BASE / "index.html"

TAMANO_MAXIMO = 4 * 1024 ** 3      # 4 GB por archivo
BLOQUE = 1024 * 1024

# Un solo procesamiento a la vez: el OCR ya usa todos los nucleos.
CANDADO_PROCESO = threading.Lock()

TRABAJOS = {}
CANDADO_TRABAJOS = threading.Lock()


TIPOS = {
    # tipo: (archivos requeridos, nombre de salida, nombre de descarga, content-type)
    "unir": (
        ("imagenes", "pdfs"), "expedientes_unidos.zip",
        f"expedientes_unidos_V{unificador.VERSION}.zip", "application/zip",
    ),
    "excel": (
        ("excel", "pdfs"), "acta_completada.xlsx", "acta_completada.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ),
}


class Trabajo:

    def __init__(self, tipo="unir"):

        self.tipo = tipo
        self.etapa = ""

        self.id = uuid.uuid4().hex[:12]
        self.carpeta = CARPETA_TRABAJOS / self.id
        self.carpeta.mkdir(parents=True, exist_ok=True)

        self.log = []
        self.actual = 0
        self.total = 0
        self.estado = "esperando"     # esperando|en_cola|procesando|listo|error
        self.resumen = ""
        self.error = ""
        self.candado = threading.Lock()

    def agregar_log(self, texto):
        with self.candado:
            self.log.extend(str(texto).split("\n"))

    def progreso(self, actual, total, etapa=None):
        self.actual, self.total = actual, total
        if etapa:
            self.etapa = etapa

    def archivo(self, nombre):
        return self.carpeta / ("acta.xlsx" if nombre == "excel" else f"{nombre}.zip")

    @property
    def imagenes(self):
        return self.carpeta / "imagenes.zip"

    @property
    def pdfs(self):
        return self.carpeta / "pdfs.zip"

    @property
    def salida(self):
        return self.carpeta / TIPOS[self.tipo][1]


def ejecutar(trabajo):

    trabajo.estado = "en_cola"
    trabajo.agregar_log("En cola, esperando turno...")

    with CANDADO_PROCESO:

        trabajo.estado = "procesando"

        try:

            unificador.verificar_tesseract(trabajo.agregar_log)

            temporal = trabajo.carpeta / ("salida.tmp" + trabajo.salida.suffix)

            if trabajo.tipo == "excel":

                trabajo.agregar_log("COMPLETAR EXCEL: hoja de firmas + Mina")

                trabajo.resumen = completar_excel.procesar(
                    str(trabajo.archivo("excel")),
                    str(trabajo.pdfs),
                    str(temporal),
                    trabajo.agregar_log,
                    trabajo.progreso
                )

            else:

                trabajo.agregar_log(f"UNIFICADOR V{unificador.VERSION}")

                trabajo.resumen = unificador.procesar(
                    str(trabajo.imagenes),
                    str(trabajo.pdfs),
                    str(temporal),
                    trabajo.agregar_log,
                    trabajo.progreso
                )

            temporal.replace(trabajo.salida)

            # Los archivos subidos ya no hacen falta.
            for nombre in TIPOS[trabajo.tipo][0]:
                trabajo.archivo(nombre).unlink(missing_ok=True)

            trabajo.estado = "listo"

        except Exception as error:

            trabajo.error = str(error) or error.__class__.__name__
            trabajo.agregar_log("ERROR: " + trabajo.error)
            trabajo.agregar_log(traceback.format_exc())
            trabajo.estado = "error"


class Manejador(BaseHTTPRequestHandler):

    server_version = "Unificador"

    def log_message(self, formato, *args):
        # Silenciar el sondeo de estado en la consola.
        if "/estado" not in (args[0] if args else ""):
            super().log_message(formato, *args)

    # ---------- utilidades ----------

    def responder_json(self, datos, codigo=200):

        cuerpo = json.dumps(datos, ensure_ascii=False).encode("utf-8")

        self.send_response(codigo)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(cuerpo)

    def error_json(self, codigo, mensaje):
        self.responder_json({"error": mensaje}, codigo)

    def obtener_trabajo(self, id_trabajo):
        with CANDADO_TRABAJOS:
            return TRABAJOS.get(id_trabajo)

    # ---------- GET ----------

    def do_GET(self):

        url = urlparse(self.path)
        partes = url.path.strip("/").split("/")

        if url.path in ("/", "/index.html"):

            cuerpo = PAGINA.read_bytes()

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(cuerpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(cuerpo)
            return

        if len(partes) == 4 and partes[:2] == ["api", "trabajo"]:

            trabajo = self.obtener_trabajo(partes[2])

            if trabajo is None:
                return self.error_json(404, "Trabajo no encontrado.")

            if partes[3] == "estado":

                desde = int(parse_qs(url.query).get("desde", ["0"])[0])

                with trabajo.candado:
                    lineas = trabajo.log[desde:]
                    total_lineas = len(trabajo.log)

                return self.responder_json({
                    "estado": trabajo.estado,
                    "actual": trabajo.actual,
                    "total": trabajo.total,
                    "etapa": trabajo.etapa,
                    "log": lineas,
                    "siguiente": total_lineas,
                    "resumen": trabajo.resumen,
                    "error": trabajo.error,
                })

            if partes[3] == "descargar":

                if trabajo.estado != "listo" or not trabajo.salida.exists():
                    return self.error_json(409, "El resultado aun no esta listo.")

                tamano = trabajo.salida.stat().st_size
                _, _, descarga, tipo_contenido = TIPOS[trabajo.tipo]

                self.send_response(200)
                self.send_header("Content-Type", tipo_contenido)
                self.send_header("Content-Length", str(tamano))
                self.send_header(
                    "Content-Disposition", f'attachment; filename="{descarga}"'
                )
                self.end_headers()

                with trabajo.salida.open("rb") as archivo:
                    shutil.copyfileobj(archivo, self.wfile, BLOQUE)

                return

        self.error_json(404, "Ruta no encontrada.")

    # ---------- POST ----------

    def do_POST(self):

        url = urlparse(self.path)
        partes = url.path.strip("/").split("/")

        # Crear trabajo (?tipo=unir | excel)
        if partes == ["api", "trabajo"]:

            tipo = parse_qs(url.query).get("tipo", ["unir"])[0]

            if tipo not in TIPOS:
                return self.error_json(400, "Tipo de trabajo desconocido.")

            trabajo = Trabajo(tipo)

            with CANDADO_TRABAJOS:
                TRABAJOS[trabajo.id] = trabajo

            return self.responder_json({"id": trabajo.id})

        if len(partes) == 4 and partes[:2] == ["api", "trabajo"]:

            trabajo = self.obtener_trabajo(partes[2])

            if trabajo is None:
                return self.error_json(404, "Trabajo no encontrado.")

            if trabajo.estado != "esperando":
                return self.error_json(409, "El trabajo ya fue iniciado.")

            # Subir un archivo (cuerpo binario crudo, sin multipart)
            if partes[3] in TIPOS[trabajo.tipo][0]:
                return self.recibir_zip(trabajo, partes[3])

            if partes[3] == "iniciar":

                for nombre in TIPOS[trabajo.tipo][0]:
                    if not trabajo.archivo(nombre).exists():
                        return self.error_json(400, f"Falta el archivo de {nombre}.")

                threading.Thread(
                    target=ejecutar, args=(trabajo,), daemon=True
                ).start()

                return self.responder_json({"ok": True})

        self.error_json(404, "Ruta no encontrada.")

    def recibir_zip(self, trabajo, tipo):

        try:
            largo = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            largo = 0

        if largo <= 0:
            return self.error_json(400, "Archivo vacio.")

        if largo > TAMANO_MAXIMO:
            return self.error_json(413, "Archivo demasiado grande.")

        destino = trabajo.archivo(tipo)
        temporal = destino.with_suffix(".subiendo")

        restante = largo

        with temporal.open("wb") as archivo:

            while restante > 0:

                bloque = self.rfile.read(min(BLOQUE, restante))

                if not bloque:
                    break

                archivo.write(bloque)
                restante -= len(bloque)

        if restante:
            temporal.unlink(missing_ok=True)
            return self.error_json(400, "La subida se interrumpio.")

        # Un .xlsx tambien es un ZIP por dentro.
        if not zipfile.is_zipfile(temporal):
            temporal.unlink(missing_ok=True)
            descripcion = "un Excel .xlsx" if tipo == "excel" else "un ZIP"
            return self.error_json(400, f"El archivo de {tipo} no es {descripcion} valido.")

        temporal.replace(destino)

        self.responder_json({"ok": True, "bytes": largo})


def main():

    puerto = int(sys.argv[1]) if len(sys.argv) > 1 else 5000

    CARPETA_TRABAJOS.mkdir(exist_ok=True)

    # Credenciales de Mina (MINA_USUARIO, MINA_CLAVE, MINA_AUTORIZADO).
    completar_excel.cargar_env()

    servidor = ThreadingHTTPServer(("0.0.0.0", puerto), Manejador)

    print(f"Unificador V{unificador.VERSION} listo en http://localhost:{puerto}")
    print("Ctrl+C para detener.")

    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        servidor.server_close()


if __name__ == "__main__":
    main()
