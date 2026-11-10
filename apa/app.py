"""
Convertidor web a formato APA 7 (Flask) con contraseña de acceso.

Ejecutar:
  pip install -r requirements.txt
  python app.py
Abre http://127.0.0.1:5000 y escribe la contraseña.

Variables de entorno opcionales (recomendado al publicarlo en internet):
  APA_PASSWORD  -> contraseña (por defecto: consola)
  SECRET_KEY    -> clave para firmar la sesión (por defecto: aleatoria)
  PORT          -> puerto (por defecto: 5000)
"""

import hmac
import io
import os
import time
from datetime import date
from pathlib import Path

from docx import Document
from flask import Flask, redirect, render_template_string, request, send_file, session

from convertir_apa import formatear_documento

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # máximo 20 MB
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(32)
CONTRASENA = os.environ.get("APA_PASSWORD", "consola")

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

ESTILO = """
<style>
  *{box-sizing:border-box}
  body{font-family:Segoe UI,Arial,sans-serif;background:linear-gradient(135deg,#1d2b64,#3a5bd9);
       margin:0;padding:24px;color:#222;min-height:100vh}
  .caja{max-width:560px;margin:30px auto;background:#fff;padding:30px;border-radius:14px;
        box-shadow:0 8px 30px rgba(0,0,0,.25)}
  h1{font-size:1.45rem;margin:0 0 6px}
  .sub{color:#667;margin:0 0 18px;font-size:.95rem}
  label{display:block;font-size:.85rem;font-weight:600;margin:12px 0 4px}
  input[type=text],input[type=password]{width:100%;padding:10px;border:1px solid #c5cbe0;
        border-radius:8px;font-size:1rem}
  input[type=file]{width:100%;padding:12px;border:2px dashed #9aa5c0;border-radius:8px;background:#f7f8fc}
  fieldset{border:1px solid #dde1ef;border-radius:10px;margin:16px 0;padding:6px 14px 14px}
  legend{font-weight:700;font-size:.9rem;padding:0 6px;color:#2f4bbd}
  .check{display:flex;gap:8px;align-items:center;font-weight:400;margin-top:12px}
  button{width:100%;padding:13px;font-size:1.05rem;background:#2f5bea;color:#fff;border:0;
        border-radius:8px;cursor:pointer;margin-top:10px}
  button:hover{background:#2448c0}
  .error{background:#fde8e8;color:#a11;padding:10px;border-radius:8px;margin-bottom:12px}
  .salir{display:block;text-align:right;font-size:.85rem;color:#2f5bea;text-decoration:none}
  small{color:#778}
</style>"""

LOGIN = """<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Acceso - Convertidor APA 7</title>""" + ESTILO + """</head><body>
<div class="caja">
  <h1>🔒 Convertidor APA 7</h1>
  <p class="sub">Escribe la contraseña para entrar.</p>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post" action="/login">
    <label for="clave">Contraseña</label>
    <input type="password" id="clave" name="clave" autofocus required>
    <button type="submit">Entrar</button>
  </form>
</div></body></html>"""

PAGINA = """<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Convertidor APA 7</title>""" + ESTILO + """</head><body>
<div class="caja">
  <a class="salir" href="/salir">Cerrar sesión</a>
  <h1>Convertidor a formato APA 7</h1>
  <p class="sub">Times New Roman 12 · interlineado 1,5 · justificado · márgenes 2,5 cm ·
     número de página · títulos, tablas y referencias APA 7.</p>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post" enctype="multipart/form-data">
    <label>Documento de Word (.docx)</label>
    <input type="file" name="archivo" accept=".docx" required>

    <fieldset>
      <legend>Portada APA 7 (opcional)</legend>
      <small>Si dejas el título vacío, no se agrega portada.</small>
      <label>Título del trabajo</label><input type="text" name="titulo">
      <label>Autor</label><input type="text" name="autor">
      <label>Institución</label><input type="text" name="institucion">
      <label>Carrera o programa</label><input type="text" name="programa">
      <label>Curso</label><input type="text" name="curso">
      <label>Asesor o docente</label><input type="text" name="asesor">
      <label>Fecha</label><input type="text" name="fecha" value="{{ fecha }}">
    </fieldset>

    <label class="check"><input type="checkbox" name="ordenar_refs" checked>
      Ordenar las referencias en orden alfabético</label>
    <label class="check"><input type="checkbox" name="conservar" checked>
      Conservar la estructura (para formularios e informes con formato propio: no borra espacios ni cambia títulos o tablas)</label>
    <button type="submit">Convertir y descargar</button>
  </form>
</div></body></html>"""


def fecha_hoy():
    h = date.today()
    return f"{h.day} de {MESES[h.month - 1]} de {h.year}"


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        clave = request.form.get("clave", "")
        if hmac.compare_digest(clave.encode(), CONTRASENA.encode()):
            session["ok"] = True
            return redirect("/")
        time.sleep(1)  # frena intentos repetidos
        return render_template_string(LOGIN, error="Contraseña incorrecta."), 401
    return render_template_string(LOGIN, error=None)


@app.route("/salir")
def salir():
    session.clear()
    return redirect("/login")


@app.route("/", methods=["GET", "POST"])
def inicio():
    if not session.get("ok"):
        return redirect("/login")

    if request.method == "GET":
        return render_template_string(PAGINA, error=None, fecha=fecha_hoy())

    def error(msg):
        return render_template_string(PAGINA, error=msg, fecha=fecha_hoy()), 400

    archivo = request.files.get("archivo")
    if not archivo or not archivo.filename:
        return error("Elige un archivo.")
    if not archivo.filename.lower().endswith(".docx"):
        return error("El archivo debe ser .docx (si es .doc, guárdalo como .docx).")

    portada = {k: request.form.get(k, "") for k in
               ("titulo", "autor", "institucion", "programa", "curso", "asesor", "fecha")}
    try:
        doc = Document(io.BytesIO(archivo.read()))
        formatear_documento(doc, portada=portada,
                            ordenar_refs=bool(request.form.get("ordenar_refs")),
                            conservar=bool(request.form.get("conservar")))
        salida = io.BytesIO()
        doc.save(salida)
        salida.seek(0)
    except Exception:
        return error("No se pudo procesar el archivo. Verifica que sea un .docx válido.")

    return send_file(
        salida,
        as_attachment=True,
        download_name=Path(archivo.filename).stem + "_APA.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
