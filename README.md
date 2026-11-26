# Unificador de Expedientes BanEcuador (V5.0)

Une las fotos de las carátulas (ZIP de imágenes) con los PDF de cada expediente
(ZIP de PDF), identificando el **Coactivado** y el **número de proceso** por OCR.

## Requisitos

1. Python 3.10+
2. `pip install -r requirements.txt`
3. Tesseract OCR 5 (Windows: https://github.com/UB-Mannheim/tesseract/wiki)
   con los idiomas **spa** y **osd** (marcarlos en el instalador).

## Uso

- Interfaz gráfica: `python unificador.py`
- Consola: `python unificador.py imagenes.zip pdfs.zip salida.zip`

## Rotación automática

1. Se aplica la orientación EXIF de la cámara.
2. Tesseract OSD propone un ángulo.
3. Se verifica con OCR real: un texto leído de lado o de cabeza casi no produce
   palabras válidas, así que gana la orientación con más palabras legibles
   (+ palabras clave, Coactivado y número de proceso).
4. Si OSD no es confiable, se evalúan los 4 ángulos (0/90/180/270).

## Completar Excel (pestaña "Completar Excel")

Llena las columnas vacías del acta de entrega (Deudor 2, Garantes 1–3: nombre,
cédula, edad, correo; y correo del deudor principal).

1. Las cédulas salen de la hoja de firmas del pagaré de cada PDF (hoja 4, o la
   que tenga "FIRMA DE LOS DEUDORES / GARANTES SOLIDARIOS"), validadas con el
   dígito verificador.
2. Nombre exacto, edad y correos se consultan en Mina con Playwright.
3. Solo se escriben celdas vacías (quedan en verde) y se agrega la hoja `Reporte`.

Requiere un archivo `.env` (no se sube a git) con:

```
MINA_USUARIO=...
MINA_CLAVE=...
MINA_AUTORIZADO=si
```

y el navegador de Playwright: `python -m playwright install chromium`.

Consola: `python completar_excel.py acta.xlsx pdfs.zip salida.xlsx`
