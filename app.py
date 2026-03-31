from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import re
from datetime import datetime
 
# PDF (reportlab)
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
 
# ======================================================
# APP
# ======================================================
app = Flask(__name__)
 
# ======================================================
# RUTAS
# ======================================================
BASE_DIR = r"C:\Users\s1868070\OneDrive - The Bank of Nova Scotia\CL-OPERACIONES-BOINT - Manuales y procedimientos\LBTR"
RUTA_CLIENTES = os.path.join(BASE_DIR, "cliente.xlsx")
RUTA_TABLA = os.path.join(BASE_DIR, "TABLA_PROCEDIMIENTO.xlsx")
RUTA_PDF = BASE_DIR
 
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)
 
# ======================================================
# FUNCIONES AUXILIARES
# ======================================================
def normalizar_rut(rut: str) -> str:
    return (
        str(rut)
        .replace(".", "")
        .replace("-", "")
        .replace(" ", "")
        .upper()
        .strip()
    )
 
def detectar_banca(segmento: str):
    if not segmento:
        return None
    segmento = str(segmento).upper()
    reglas = {
        "RETAIL": "Retail",
        "PERSONA": "Retail",
        "SUCURSAL": "Retail",
        "WHOLE": "Wholesale",
        "EMPRESA": "Wholesale",
        "PYME": "Wholesale",
        "WEALTH": "Wealth",
        "PRIVADA": "Wealth",
        "WM": "Wealth",
    }
    for palabra, hoja in reglas.items():
        if palabra in segmento:
            return hoja
    return None
 
def extraer_nombre_pdf(texto):
    if not isinstance(texto, str):
        return ""
    texto = re.sub(r"[\[\]\(\)\\/:]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
 
    m = re.search(r"([A-Za-z0-9_.-]+)\s*\((\d+)\)", texto)
    if m:
        return f"{m.group(1)} ({m.group(2)}).pdf"
 
    m = re.search(r"([A-Za-z0-9_.-]+)\.pdf", texto, re.IGNORECASE)
    if m:
        return f"{m.group(1)}.pdf"
 
    return ""
 
def parse_checklist(data: dict):
    """
    ✅ NUEVO FORMATO:
      value = "TEMA||TIP\\nSI|NO|NA"
 
    Compatibilidad:
      - "TEMA\\nSI" => TIP vacío
      - "SI" => Item queda el índice
    """
    rows = []
    for key, value in (data or {}).items():
        if str(key).startswith("check_"):
            item = key.replace("check_", "").strip()
            tip = ""
            respuesta = ""
 
            if isinstance(value, str) and "\n" in value:
                izquierda, resp_val = value.split("\n", 1)
                respuesta = (resp_val or "").strip()
 
                if "||" in izquierda:
                    item_val, tip_val = izquierda.split("||", 1)
                    item = (item_val or item).strip()
                    tip = (tip_val or "").strip()
                else:
                    item = (izquierda or item).strip()
                    tip = ""
            else:
                respuesta = str(value).strip()
 
            rows.append({"Item": item, "Subtitulo": tip, "Respuesta": respuesta})
    return rows
 
def wrap_text(c: canvas.Canvas, text: str, x: int, y: int, max_width: int, line_height: int,
              font_name: str = "Helvetica", font_size: int = 10):
    """
    Dibuja texto con salto automático por ancho.
    Retorna el nuevo Y.
    """
    if text is None:
        text = ""
    c.setFont(font_name, font_size)
    words = str(text).split()
    line = ""
    for w in words:
        test = (line + " " + w).strip()
        if c.stringWidth(test, font_name, font_size) <= max_width:
            line = test
        else:
            if line:
                c.drawString(x, y, line)
                y -= line_height
            line = w
    if line:
        c.drawString(x, y, line)
        y -= line_height
    return y
 
def safe_folder_name(name: str) -> str:
    """
    Sanitiza el nombre para usarlo como carpeta en Windows.
    Reemplaza caracteres inválidos: < > : " / \\ | ? *
    """
    if not name:
        return "SIN_ANALISTA"
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "SIN_ANALISTA"
 
def get_analyst_output_dir(analista: str) -> str:
    folder = safe_folder_name(analista)
    out_dir = os.path.join(OUTPUT_DIR, folder)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir
 
# ======================================================
# CARGA DE EXCEL (al iniciar la app)
# ======================================================
CLIENTES = pd.read_excel(RUTA_CLIENTES, engine="openpyxl")
CLIENTES.columns = CLIENTES.columns.str.upper().str.strip()
 
if "RUT_FINAL" not in CLIENTES.columns:
    raise ValueError("No existe columna 'RUT_FINAL' en cliente.xlsx. Revisa el nombre real de la columna.")
 
CLIENTES["RUT_LIMPIO"] = CLIENTES["RUT_FINAL"].apply(normalizar_rut)
 
TABLAS = pd.read_excel(RUTA_TABLA, sheet_name=None, header=1, engine="openpyxl")
for hoja in TABLAS:
    TABLAS[hoja].columns = TABLAS[hoja].columns.str.upper().str.strip()
 
# ======================================================
# HOME
# ======================================================
@app.route("/")
def home():
    return render_template("form.html")
 
# ======================================================
# BUSCAR CLIENTE POR RUT
# ======================================================
@app.route("/buscar_cliente", methods=["POST"])
def buscar_cliente():
    data = request.get_json(silent=True) or {}
    rut = data.get("rut")
    if not rut:
        return jsonify({"error": "RUT no recibido"}), 400
 
    rut_limpio = normalizar_rut(rut)
    match = CLIENTES[CLIENTES["RUT_LIMPIO"] == rut_limpio]
    if match.empty:
        return jsonify({"error": "Cliente no encontrado"}), 404
 
    fila = match.iloc[0]
    nombre = fila.get("NOMBRE CLIENTE", fila.get("NOMBRE", ""))
    segmento = fila.get("SEGMENTO_BANCA", fila.get("SEGMENTO", ""))
    return jsonify({"nombre": nombre, "segmento": segmento})
 
# ======================================================
# CHECKLIST (SEGMENTO + PRODUCTO)
# ======================================================
@app.route("/checklist", methods=["POST"])
def checklist_endpoint():
    data = request.get_json(silent=True) or {}
    segmento = data.get("segmento")
    producto = data.get("producto")
 
    if not segmento or not producto:
        return jsonify([])
 
    banca = detectar_banca(segmento)
    if not banca or banca not in TABLAS:
        return jsonify([])
 
    df = TABLAS[banca].fillna("")
    items = []
    for _, row in df.iterrows():
        pdf = extraer_nombre_pdf(row.get("FUENTE", ""))
        link = os.path.join(RUTA_PDF, pdf) if pdf else ""
        items.append({
            "tema": row.get("TEMA", ""),
            "tip": row.get("TIP", ""),
            "link": link
        })
    return jsonify(items)
 
# ======================================================
# GUARDAR EXCEL (Datos + Checklist) y descargar
# ======================================================
@app.route("/guardar", methods=["POST"])
def guardar():
    data = request.get_json(silent=True) or {}
 
    analista = data.get("analista", "")
    out_dir = get_analyst_output_dir(analista)
 
    datos_principales = {
        "RUT": data.get("rut", ""),
        "Nombre": data.get("nombre", ""),
        "Segmento": data.get("segmento", ""),
        "Producto": data.get("producto", ""),
        "Analista": analista,
        "Fecha_Generacion": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    }
    df_main = pd.DataFrame([datos_principales])
 
    checklist_rows = parse_checklist(data)
    df_check = (
        pd.DataFrame(checklist_rows)
        if checklist_rows
        else pd.DataFrame(columns=["Item", "Subtitulo", "Respuesta"])
    )
 
    rut_limpio = normalizar_rut(data.get("rut", "SINRUT"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"Checklist_{rut_limpio}_{timestamp}.xlsx"
    ruta_salida = os.path.join(out_dir, nombre_archivo)
 
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df_main.to_excel(writer, sheet_name="Datos", index=False)
        df_check.to_excel(writer, sheet_name="Checklist", index=False)
 
    return send_file(ruta_salida, as_attachment=True, download_name=nombre_archivo)
 
# ======================================================
# GUARDAR PDF y descargar
# ======================================================
@app.route("/guardar_pdf", methods=["POST"])
def guardar_pdf():
    data = request.get_json(silent=True) or {}
 
    analista = data.get("analista", "")
    out_dir = get_analyst_output_dir(analista)
 
    rut_limpio = normalizar_rut(data.get("rut", "SINRUT"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"Checklist_{rut_limpio}_{timestamp}.pdf"
    ruta_salida = os.path.join(out_dir, nombre_archivo)
 
    checklist_rows = parse_checklist(data)
 
    c = canvas.Canvas(ruta_salida, pagesize=letter)
    width, height = letter
    margin_x = 50
    y = height - 50
 
    # Encabezado
    c.setFont("Helvetica-Bold", 14)
    c.drawString(margin_x, y, "Checklist LBTR")
    y -= 22
 
    y = wrap_text(c, f"RUT: {data.get('rut','')}", margin_x, y, width - 2 * margin_x, 14, "Helvetica", 10)
    y = wrap_text(c, f"Nombre: {data.get('nombre','')}", margin_x, y, width - 2 * margin_x, 14, "Helvetica", 10)
    y = wrap_text(c, f"Segmento: {data.get('segmento','')}", margin_x, y, width - 2 * margin_x, 14, "Helvetica", 10)
    y = wrap_text(c, f"Producto: {data.get('producto','')}", margin_x, y, width - 2 * margin_x, 14, "Helvetica", 10)
    y = wrap_text(c, f"Analista: {analista}", margin_x, y, width - 2 * margin_x, 14, "Helvetica", 10)
    y = wrap_text(c, f"Fecha: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}", margin_x, y,
                  width - 2 * margin_x, 14, "Helvetica", 10)
 
    y -= 8
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin_x, y, "Checklist:")
    y -= 18
 
    if not checklist_rows:
        c.setFont("Helvetica", 10)
        c.drawString(margin_x, y, "- (Sin ítems seleccionados)")
        y -= 14
    else:
        for row in checklist_rows:
            item = row.get("Item", "")
            tip = row.get("Subtitulo", "")
            resp = row.get("Respuesta", "")
 
            # salto de página
            if y < 90:
                c.showPage()
                y = height - 50
 
            # Item + respuesta (negrita)
            y = wrap_text(c, f"- {item}  [{resp}]", margin_x, y, width - 2 * margin_x, 14, "Helvetica-Bold", 10)
 
            # Subtítulo (debajo, más pequeño)
            if tip:
                y = wrap_text(c, f"  {tip}", margin_x, y, width - 2 * margin_x, 12, "Helvetica", 9)
 
            y -= 4
 
    c.save()
    return send_file(ruta_salida, as_attachment=True, download_name=nombre_archivo)
 
# ======================================================
# MAIN
# ======================================================
 
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)