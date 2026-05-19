import os
import re
import uuid
from datetime import datetime, date
from io import BytesIO

import pandas as pd
import streamlit as st

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="LexMonitor AI",
    page_icon="⚖️",
    layout="wide",
)

APP_TITLE = "LexMonitor AI"
LOCAL_CSV = "expedientes_demo.csv"
SHEET_NAME = "expedientes"
LIMITE_ABREVIADO = 50000

COLUMNAS = [
    "codigo",
    "fecha_creacion",
    "estado",
    "rol_ultima_accion",
    "demandante",
    "email_demandante",
    "demandado",
    "email_demandado",
    "cuantia",
    "concepto_deuda",
    "hechos",
    "documentacion_valida",
    "categorias_art_812",
    "documentos_subidos",
    "respuesta_deudor",
    "motivo_oposicion",
    "fecha_respuesta_deudor",
    "accion_recomendada",
    "borrador",
]

ARTICULOS = {
    "art_812": {
        "referencia": "Art. 812 LEC",
        "resumen": (
            "Permite acudir al proceso monitorio cuando se reclama una deuda dineraria, líquida, "
            "determinada, vencida y exigible, acreditada mediante documentos firmados o aceptados "
            "por el deudor, facturas, albaranes, certificaciones u otros documentos habituales, "
            "documentos comerciales de relación duradera o certificaciones de impago de comunidades."
        ),
    },
    "admision": {
        "referencia": "Art. 815 LEC",
        "resumen": "Si los documentos aportados constituyen principio de prueba del derecho del peticionario, se requerirá al deudor para pagar o comparecer y alegar oposición.",
    },
    "pago": {
        "referencia": "Art. 817 LEC",
        "resumen": "Si el deudor atiende el requerimiento de pago, se archivan las actuaciones.",
    },
    "oposicion": {
        "referencia": "Art. 818 LEC",
        "resumen": "Si el deudor formula oposición, el asunto se resuelve por el juicio que corresponda según la cuantía.",
    },
    "incomparecencia": {
        "referencia": "Art. 816 LEC",
        "resumen": "Si el deudor no paga ni comparece, se dicta decreto dando por terminado el monitorio y se despacha ejecución.",
    },
}

# ============================================================
# GOOGLE SHEETS / ALMACENAMIENTO LOCAL
# ============================================================

def get_storage_mode():
    """Devuelve 'sheets' si hay configuración válida; si no, 'local'."""
    if gspread is None or Credentials is None:
        return "local"

    has_secrets = "gcp_service_account" in st.secrets and "SPREADSHEET_ID" in st.secrets
    return "sheets" if has_secrets else "local"


def get_worksheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scopes
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])

    try:
        ws = spreadsheet.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(COLUMNAS))
        ws.append_row(COLUMNAS)

    headers = ws.row_values(1)
    if headers != COLUMNAS:
        ws.clear()
        ws.append_row(COLUMNAS)

    return ws


def empty_df():
    return pd.DataFrame(columns=COLUMNAS)


def load_data():
    mode = get_storage_mode()

    if mode == "sheets":
        ws = get_worksheet()
        records = ws.get_all_records()
        if not records:
            return empty_df()
        df = pd.DataFrame(records)
        for col in COLUMNAS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUMNAS]

    if os.path.exists(LOCAL_CSV):
        df = pd.read_csv(LOCAL_CSV, dtype=str).fillna("")
        for col in COLUMNAS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUMNAS]

    return empty_df()


def save_record(record):
    record = {col: str(record.get(col, "")) for col in COLUMNAS}
    mode = get_storage_mode()

    if mode == "sheets":
        ws = get_worksheet()
        ws.append_row([record[col] for col in COLUMNAS])
        return

    df = load_data()
    df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
    df.to_csv(LOCAL_CSV, index=False)


def update_record(codigo, updates):
    mode = get_storage_mode()

    if mode == "sheets":
        ws = get_worksheet()
        values = ws.get_all_values()
        if not values:
            return False
        headers = values[0]
        try:
            codigo_idx = headers.index("codigo") + 1
        except ValueError:
            return False

        codigos = ws.col_values(codigo_idx)
        if codigo not in codigos:
            return False
        row_number = codigos.index(codigo) + 1

        for key, value in updates.items():
            if key in headers:
                col_number = headers.index(key) + 1
                ws.update_cell(row_number, col_number, str(value))
        return True

    df = load_data()
    if df.empty or codigo not in df["codigo"].values:
        return False

    for key, value in updates.items():
        if key in df.columns:
            df.loc[df["codigo"] == codigo, key] = str(value)

    df.to_csv(LOCAL_CSV, index=False)
    return True


def get_record(codigo):
    df = load_data()
    if df.empty:
        return None
    coincidencias = df[df["codigo"].astype(str).str.upper() == codigo.upper()]
    if coincidencias.empty:
        return None
    return coincidencias.iloc[0].to_dict()

# ============================================================
# LECTURA DE ARCHIVOS
# ============================================================

def leer_txt(archivo):
    return archivo.read().decode("utf-8", errors="ignore")


def leer_docx(archivo):
    if Document is None:
        return "ERROR: instala python-docx con: pip install python-docx"
    doc = Document(archivo)
    textos = []
    for p in doc.paragraphs:
        if p.text.strip():
            textos.append(p.text.strip())
    for tabla in doc.tables:
        for fila in tabla.rows:
            textos.append(" | ".join(c.text.strip() for c in fila.cells))
    return "\n".join(textos)


def leer_pdf(archivo):
    if PyPDF2 is None:
        return "ERROR: instala PyPDF2 con: pip install PyPDF2"
    reader = PyPDF2.PdfReader(archivo)
    textos = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            textos.append(text)
    return "\n".join(textos)


def leer_documento(archivo):
    nombre = archivo.name.lower()
    if nombre.endswith(".txt"):
        return leer_txt(archivo)
    if nombre.endswith(".docx"):
        return leer_docx(archivo)
    if nombre.endswith(".pdf"):
        return leer_pdf(archivo)
    return ""

# ============================================================
# EXTRACCIÓN Y VALIDACIÓN ART. 812 LEC
# ============================================================

def limpiar_numero(texto_numero):
    if not texto_numero:
        return 0.0
    texto_numero = texto_numero.lower().replace("euros", "").replace("eur", "")
    texto_numero = texto_numero.replace("€", "").replace("lempiras", "").replace("l.", "")
    texto_numero = texto_numero.strip()

    if "," in texto_numero and "." in texto_numero:
        if texto_numero.find(".") < texto_numero.find(","):
            texto_numero = texto_numero.replace(".", "").replace(",", ".")
        else:
            texto_numero = texto_numero.replace(",", "")
    elif "," in texto_numero:
        partes = texto_numero.split(",")
        texto_numero = texto_numero.replace(",", ".") if len(partes[-1]) == 2 else texto_numero.replace(",", "")
    elif "." in texto_numero:
        partes = texto_numero.split(".")
        if len(partes[-1]) != 2:
            texto_numero = texto_numero.replace(".", "")
    try:
        return float(texto_numero)
    except ValueError:
        return 0.0


def buscar_patron(texto, patrones):
    for patron in patrones:
        match = re.search(patron, texto, re.IGNORECASE | re.MULTILINE)
        if match:
            return match.group(1).strip()
    return ""


def detectar_documentos_art_812(texto):
    texto = texto.lower()
    categorias = {
        "Documentos firmados por el deudor o con sello/señal": [
            "firmado por el deudor", "firma del deudor", "sello del deudor", "impronta",
            "marca del deudor", "firma electrónica", "firma electronica", "señal electrónica", "senal electronica",
            "aceptado por el deudor", "conformidad del deudor",
        ],
        "Facturas, albaranes, certificaciones, telegramas, fax u otros documentos habituales": [
            "factura", "facturas", "albarán", "albaran", "albaranes", "certificación", "certificacion",
            "certificaciones", "telegrama", "telegramas", "fax", "burofax", "recibo", "recibos",
            "estado de cuenta", "comprobante", "orden de pedido", "pedido", "presupuesto aceptado",
        ],
        "Documentos comerciales de relación anterior duradera": [
            "relación comercial anterior", "relacion comercial anterior", "relación anterior duradera",
            "relacion anterior duradera", "relación contractual continuada", "relacion contractual continuada",
            "contrato marco", "contrato de suministro", "historial de pedidos", "documentos comerciales",
            "relación comercial continuada", "relacion comercial continuada",
        ],
        "Certificaciones de impago de comunidades de propietarios": [
            "certificación de impago", "certificacion de impago", "gastos comunes", "comunidad de propietarios",
            "inmueble urbano", "cuotas comunitarias", "cuota comunitaria", "acta de la comunidad",
        ],
    }
    detectados = {}
    for categoria, palabras in categorias.items():
        coincidencias = sorted(set(p for p in palabras if p in texto))
        detectados[categoria] = coincidencias
    hay_valido = any(detectados[c] for c in detectados)
    categorias_detectadas = [c for c, v in detectados.items() if v]
    return hay_valido, detectados, categorias_detectadas


def extraer_datos_demanda(texto_demanda, texto_documentos=""):
    texto_total = f"{texto_demanda}\n\n{texto_documentos}"
    texto_unido = re.sub(r"\s+", " ", texto_total)
    texto_demanda_unido = re.sub(r"\s+", " ", texto_demanda)

    demandante = buscar_patron(texto_demanda_unido, [
        r"demandante[:\s]+(.+?)(?:demandado|deudor|contra|frente a|,|\.)",
        r"acreedor[:\s]+(.+?)(?:demandado|deudor|contra|frente a|,|\.)",
        r"solicitante[:\s]+(.+?)(?:demandado|deudor|contra|frente a|,|\.)",
        r"a instancia de[:\s]+(.+?)(?:contra|frente a|,|\.)",
    ])

    demandado = buscar_patron(texto_demanda_unido, [
        r"demandado[:\s]+(.+?)(?:,|\.)",
        r"deudor[:\s]+(.+?)(?:,|\.)",
        r"contra[:\s]+(.+?)(?:,|\.)",
        r"frente a[:\s]+(.+?)(?:,|\.)",
    ])

    cuantia_txt = buscar_patron(texto_unido, [
        r"cuantía[:\s]+(?:de\s*)?(?:€|eur|euros|L\.?\s*)?([\d\.,]+)",
        r"importe[:\s]+(?:de\s*)?(?:€|eur|euros|L\.?\s*)?([\d\.,]+)",
        r"cantidad[:\s]+(?:de\s*)?(?:€|eur|euros|L\.?\s*)?([\d\.,]+)",
        r"suma[:\s]+(?:de\s*)?(?:€|eur|euros|L\.?\s*)?([\d\.,]+)",
        r"reclama(?:\s+la)?\s+cantidad\s+de\s+(?:€|eur|euros|L\.?\s*)?([\d\.,]+)",
        r"por\s+importe\s+de\s+(?:€|eur|euros|L\.?\s*)?([\d\.,]+)",
    ])
    cuantia = limpiar_numero(cuantia_txt)

    concepto = buscar_patron(texto_demanda_unido, [
        r"concepto(?:\s+de\s+la\s+deuda)?[:\s]+(.{10,250})",
        r"la deuda deriva de[:\s]+(.{10,250})",
        r"por los siguientes hechos[:\s]+(.{10,250})",
    ])

    hay_doc_812, documentos_art_812, categorias_detectadas = detectar_documentos_art_812(texto_documentos)

    if not hay_doc_812:
        hay_doc_812, documentos_art_812, categorias_detectadas = detectar_documentos_art_812(texto_total)

    menciona_monitorio = "monitorio" in texto_total.lower()
    deuda_ok = any(p in texto_total.lower() for p in [
        "deuda dineraria", "deuda líquida", "deuda liquida", "deuda determinada",
        "deuda vencida", "deuda exigible", "cantidad determinada", "vencida y exigible",
    ])

    datos_faltantes = []
    if not demandante:
        datos_faltantes.append("Demandante / acreedor")
    if not demandado:
        datos_faltantes.append("Demandado / deudor")
    if cuantia <= 0:
        datos_faltantes.append("Cuantía")
    if not hay_doc_812:
        datos_faltantes.append("Documentación válida del art. 812 LEC")

    cumple_requisitos_auto = bool(demandante) and bool(demandado) and cuantia > 0 and hay_doc_812

    return {
        "demandante": demandante,
        "demandado": demandado,
        "cuantia": cuantia,
        "concepto_deuda": concepto,
        "menciona_monitorio": menciona_monitorio,
        "deuda_ok_texto": deuda_ok,
        "documentacion_valida": hay_doc_812,
        "documentos_art_812": documentos_art_812,
        "categorias_detectadas": categorias_detectadas,
        "hechos_resumidos": texto_demanda_unido[:1000],
        "datos_faltantes": datos_faltantes,
        "cumple_requisitos_auto": cumple_requisitos_auto,
    }

# ============================================================
# GENERACIÓN DE CÓDIGO Y BORRADORES
# ============================================================

def generar_codigo():
    year = datetime.now().year
    corto = uuid.uuid4().hex[:6].upper()
    return f"MON-{year}-{corto}"


def generar_borrador_admision(record):
    return f"""BORRADOR DE ADMISIÓN / REQUERIMIENTO DE PAGO

Código de expediente: {record['codigo']}

Vista la solicitud monitoria presentada por {record['demandante']} frente a {record['demandado']}, por importe de {record['cuantia']} euros, y examinada la documentación aportada al amparo del art. 812 LEC, se aprecia inicialmente la existencia de documentación suficiente para tramitar la petición.

Procede requerir al deudor para que pague la cantidad reclamada o comparezca formulando oposición en el plazo legalmente previsto.

Referencia normativa: {ARTICULOS['art_812']['referencia']} y {ARTICULOS['admision']['referencia']}.

Actuación recomendada: notificar al demandado/deudor.
"""


def generar_borrador_subsanacion(datos):
    faltantes = ", ".join(datos.get("datos_faltantes", [])) or "documentación o datos necesarios"
    return f"""BORRADOR DE REQUERIMIENTO DE SUBSANACIÓN

Examinada la solicitud presentada, se advierte que no constan suficientemente los siguientes extremos: {faltantes}.

Antes de admitir la petición monitoria, procede requerir a la parte solicitante para que complete o aclare la documentación necesaria, especialmente la documentación acreditativa prevista en el art. 812 LEC.

Referencia normativa: {ARTICULOS['art_812']['referencia']}.
"""


def generar_borrador_respuesta(record, respuesta, motivo=""):
    if respuesta == "Pagado":
        return f"""DILIGENCIA DE PAGO Y ARCHIVO

Expediente: {record['codigo']}

Constando que el deudor {record['demandado']} ha manifestado su voluntad de pagar la deuda reclamada por {record['demandante']} por importe de {record['cuantia']} euros, procede dejar constancia del pago y archivar las actuaciones, previa comprobación efectiva del abono.

Referencia normativa: {ARTICULOS['pago']['referencia']}.
"""
    if respuesta == "Oposición presentada":
        return f"""BORRADOR DE ADMISIÓN DE OPOSICIÓN

Expediente: {record['codigo']}

El deudor {record['demandado']} formula oposición frente a la reclamación presentada por {record['demandante']}.

Motivo alegado:
{motivo or 'No especificado'}

Procede tener por formulada oposición y continuar por el procedimiento que corresponda según la cuantía.

Referencia normativa: {ARTICULOS['oposicion']['referencia']}.
"""
    return f"""BORRADOR DE FINALIZACIÓN DEL MONITORIO Y DESPACHO DE EJECUCIÓN

Expediente: {record['codigo']}

Constando que el deudor {record['demandado']} no paga ni comparece, procede dar por terminado el proceso monitorio y abrir la vía de ejecución por la cantidad reclamada de {record['cuantia']} euros, previa validación por el órgano competente.

Referencia normativa: {ARTICULOS['incomparecencia']['referencia']}.
"""

# ============================================================
# AUTENTICACIÓN FICTICIA
# ============================================================

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "perfil" not in st.session_state:
    st.session_state.perfil = None

st.title(f"⚖️ {APP_TITLE}")
st.caption("Prototipo de hackathon para proceso monitorio: registro, documentación art. 812 LEC y respuesta del deudor.")

if not st.session_state.autenticado:
    st.header("🔐 Autenticación con certificado digital")
    st.info("Para la demo, sube un archivo vacío o ficticio con nombre de certificado. No se valida criptográficamente.")
    certificado = st.file_uploader("Sube tu certificado digital", type=["pdf", "txt", "cer", "crt", "pem"], key="certificado")
    if certificado is not None:
        st.success(f"Certificado '{certificado.name}' detectado correctamente.")
        if st.button("Acceder a LexMonitor AI", type="primary"):
            st.session_state.autenticado = True
            st.rerun()
    st.stop()

# ============================================================
# MENÚ PRINCIPAL
# ============================================================

with st.sidebar:
    st.success("Certificado validado")
    modo = get_storage_mode()
    st.caption(f"Almacenamiento: {'Google Sheets' if modo == 'sheets' else 'CSV local de demo'}")
    if st.button("Cerrar sesión"):
        st.session_state.autenticado = False
        st.session_state.perfil = None
        st.rerun()

st.header("Selecciona tu perfil")
col_a, col_b = st.columns(2)
with col_a:
    if st.button("👤 Demandante / Acreedor", use_container_width=True, type="primary"):
        st.session_state.perfil = "demandante"
with col_b:
    if st.button("👤 Demandado / Deudor", use_container_width=True):
        st.session_state.perfil = "demandado"

st.divider()

# ============================================================
# PERFIL DEMANDANTE
# ============================================================

if st.session_state.perfil == "demandante":
    st.header("📄 Zona del Demandante / Acreedor")

    tab_pdf, tab_form = st.tabs(["1. Subir PDF/plantilla de demanda", "2. Rellenar formulario desde la app"])

    datos_demanda = None
    texto_demanda = ""

    with tab_pdf:
        demanda_pdf = st.file_uploader("Sube la plantilla de demanda monitoria rellena", type=["pdf", "docx", "txt"], key="demanda_file")
        if demanda_pdf:
            texto_demanda = leer_documento(demanda_pdf)
            st.text_area("Texto leído de la demanda", texto_demanda, height=220)
            datos_demanda = extraer_datos_demanda(texto_demanda, "")
            st.session_state["texto_demanda"] = texto_demanda
            st.session_state["datos_demanda_base"] = datos_demanda

    with tab_form:
        with st.form("form_demanda_manual"):
            c1, c2 = st.columns(2)
            demandante_manual = c1.text_input("Demandante / acreedor")
            email_demandante_manual = c2.text_input("Email del demandante")
            c1, c2 = st.columns(2)
            demandado_manual = c1.text_input("Demandado / deudor")
            email_demandado_manual = c2.text_input("Email del demandado")
            cuantia_manual = st.number_input("Cuantía reclamada (€)", min_value=0.0, step=100.0)
            concepto_manual = st.text_input("Concepto de la deuda")
            hechos_manual = st.text_area("Hechos principales", height=150)
            usar_form = st.form_submit_button("Usar estos datos")
        if usar_form:
            texto_demanda = f"Demandante: {demandante_manual}. Demandado: {demandado_manual}. Cuantía: {cuantia_manual}. Concepto: {concepto_manual}. Hechos: {hechos_manual}"
            datos_demanda = {
                "demandante": demandante_manual,
                "demandado": demandado_manual,
                "cuantia": cuantia_manual,
                "concepto_deuda": concepto_manual,
                "hechos_resumidos": hechos_manual,
                "menciona_monitorio": True,
                "deuda_ok_texto": True,
                "documentacion_valida": False,
                "documentos_art_812": {},
                "categorias_detectadas": [],
                "datos_faltantes": [],
                "cumple_requisitos_auto": False,
                "email_demandante": email_demandante_manual,
                "email_demandado": email_demandado_manual,
            }
            st.session_state["texto_demanda"] = texto_demanda
            st.session_state["datos_demanda_base"] = datos_demanda
            st.success("Datos de demanda cargados desde formulario.")

    datos_base = st.session_state.get("datos_demanda_base")
    texto_demanda_guardado = st.session_state.get("texto_demanda", "")

    st.subheader("📎 Documentos / pruebas para admisión art. 812 LEC")
    documentos = st.file_uploader(
        "Sube facturas, albaranes, certificaciones, documentos firmados, etc.",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
        key="docs_812",
    )

    texto_documentos = ""
    nombres_docs = []
    if documentos:
        partes = []
        for doc in documentos:
            nombres_docs.append(doc.name)
            partes.append(f"\n--- DOCUMENTO: {doc.name} ---\n{leer_documento(doc)}")
        texto_documentos = "\n".join(partes)
        st.text_area("Texto leído de documentos", texto_documentos, height=220)

    if datos_base:
        datos_finales = extraer_datos_demanda(texto_demanda_guardado, texto_documentos)
        datos_finales["email_demandante"] = datos_base.get("email_demandante", "")
        datos_finales["email_demandado"] = datos_base.get("email_demandado", "")

        # Si el PDF no detectó bien, mantenemos lo que venga del formulario o permitimos editar.
        st.subheader("✅ Validación de datos antes de registrar")
        with st.form("validar_registro"):
            c1, c2 = st.columns(2)
            demandante = c1.text_input("Demandante / acreedor", value=datos_finales.get("demandante") or datos_base.get("demandante", ""))
            demandado = c2.text_input("Demandado / deudor", value=datos_finales.get("demandado") or datos_base.get("demandado", ""))
            c1, c2 = st.columns(2)
            email_demandante = c1.text_input("Email demandante", value=datos_finales.get("email_demandante", ""))
            email_demandado = c2.text_input("Email demandado", value=datos_finales.get("email_demandado", ""))
            cuantia = st.number_input("Cuantía (€)", min_value=0.0, value=float(datos_finales.get("cuantia") or datos_base.get("cuantia") or 0.0), step=100.0)
            concepto = st.text_input("Concepto de la deuda", value=datos_finales.get("concepto_deuda") or datos_base.get("concepto_deuda", ""))
            hechos = st.text_area("Hechos", value=datos_finales.get("hechos_resumidos") or datos_base.get("hechos_resumidos", ""), height=130)

            documentacion_valida = st.checkbox(
                "Documentación válida según art. 812 LEC",
                value=bool(datos_finales.get("documentacion_valida")),
            )

            st.markdown("**Categorías detectadas art. 812 LEC**")
            docs_812 = datos_finales.get("documentos_art_812", {})
            for categoria, coincidencias in docs_812.items():
                if coincidencias:
                    st.success(f"{categoria}: {', '.join(coincidencias)}")
                else:
                    st.info(f"{categoria}: no detectado")

            registrar = st.form_submit_button("Generar código de demanda y registrar", type="primary")

        if registrar:
            faltantes = []
            if not demandante:
                faltantes.append("demandante")
            if not demandado:
                faltantes.append("demandado")
            if cuantia <= 0:
                faltantes.append("cuantía")
            if not documentacion_valida:
                faltantes.append("documentación art. 812 LEC")

            codigo = generar_codigo()
            categorias = datos_finales.get("categorias_detectadas", [])

            estado = "Pendiente de respuesta del deudor" if not faltantes else "Pendiente de subsanación"
            accion = "Notificar al demandado para pagar u oponerse" if not faltantes else "Requerir subsanación al demandante"

            temp_record = {
                "codigo": codigo,
                "demandante": demandante,
                "demandado": demandado,
                "cuantia": cuantia,
            }
            borrador = generar_borrador_admision(temp_record) if not faltantes else generar_borrador_subsanacion({"datos_faltantes": faltantes})

            record = {
                "codigo": codigo,
                "fecha_creacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "estado": estado,
                "rol_ultima_accion": "Demandante / Acreedor",
                "demandante": demandante,
                "email_demandante": email_demandante,
                "demandado": demandado,
                "email_demandado": email_demandado,
                "cuantia": cuantia,
                "concepto_deuda": concepto,
                "hechos": hechos,
                "documentacion_valida": "Sí" if documentacion_valida else "No",
                "categorias_art_812": "; ".join(categorias),
                "documentos_subidos": "; ".join(nombres_docs),
                "respuesta_deudor": "",
                "motivo_oposicion": "",
                "fecha_respuesta_deudor": "",
                "accion_recomendada": accion,
                "borrador": borrador,
            }
            save_record(record)

            if faltantes:
                st.warning("Expediente registrado, pero pendiente de subsanación: " + ", ".join(faltantes))
            else:
                st.success("Demanda registrada correctamente.")

            st.code(codigo, language="text")
            st.text_area("Borrador generado", borrador, height=300)
            st.info("Entrega este código al demandado/deudor para que pueda consultar la reclamación y responder.")
    else:
        st.info("Sube una demanda o rellena el formulario para continuar.")

# ============================================================
# PERFIL DEMANDADO
# ============================================================

elif st.session_state.perfil == "demandado":
    st.header("📬 Zona del Demandado / Deudor")
    codigo_busqueda = st.text_input("Introduce el código de demanda", placeholder="Ejemplo: MON-2026-ABC123")

    if st.button("Buscar demanda", type="primary") and codigo_busqueda:
        st.session_state["codigo_consultado"] = codigo_busqueda.strip().upper()

    codigo_actual = st.session_state.get("codigo_consultado")
    if codigo_actual:
        record = get_record(codigo_actual)
        if not record:
            st.error("No se ha encontrado ninguna demanda con ese código.")
        else:
            st.success("Demanda encontrada.")
            c1, c2, c3 = st.columns(3)
            c1.metric("Código", record.get("codigo", ""))
            c2.metric("Estado", record.get("estado", ""))
            c3.metric("Cuantía", f"{record.get('cuantia', '')} €")

            st.subheader("Resumen de la reclamación")
            st.write(f"**Demandante / acreedor:** {record.get('demandante', '')}")
            st.write(f"**Demandado / deudor:** {record.get('demandado', '')}")
            st.write(f"**Concepto de deuda:** {record.get('concepto_deuda', '')}")
            st.write(f"**Hechos:** {record.get('hechos', '')}")
            st.write(f"**Documentación aportada:** {record.get('documentos_subidos', '') or 'No consta'}")
            st.write(f"**Categorías art. 812 LEC:** {record.get('categorias_art_812', '') or 'No consta'}")
            st.write(f"**Actuación recomendada actual:** {record.get('accion_recomendada', '')}")

            st.subheader("Respuesta del demandado")
            c1, c2, c3 = st.columns(3)

            with c1:
                if st.button("💳 Pagar deuda", use_container_width=True):
                    borrador = generar_borrador_respuesta(record, "Pagado")
                    update_record(record["codigo"], {
                        "estado": "Pagado. Pendiente de comprobación y archivo",
                        "rol_ultima_accion": "Demandado / Deudor",
                        "respuesta_deudor": "Pagado",
                        "fecha_respuesta_deudor": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "accion_recomendada": "Comprobar pago y archivar actuaciones",
                        "borrador": borrador,
                    })
                    st.success("Respuesta registrada: pago de deuda.")
                    st.text_area("Borrador generado", borrador, height=260)

            with c2:
                abrir_oposicion = st.button("✍️ Oponerse al pago", use_container_width=True)
                if abrir_oposicion:
                    st.session_state["mostrar_oposicion"] = True

            with c3:
                if st.button("⏳ No pagar ni comparecer", use_container_width=True):
                    borrador = generar_borrador_respuesta(record, "No comparece")
                    update_record(record["codigo"], {
                        "estado": "Sin pago ni comparecencia. Procede ejecución",
                        "rol_ultima_accion": "Demandado / Deudor",
                        "respuesta_deudor": "No paga ni comparece",
                        "fecha_respuesta_deudor": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "accion_recomendada": "Dar por terminado el monitorio y despachar ejecución",
                        "borrador": borrador,
                    })
                    st.warning("Respuesta registrada: no pago ni comparecencia.")
                    st.text_area("Borrador generado", borrador, height=260)

            if st.session_state.get("mostrar_oposicion"):
                with st.form("form_oposicion"):
                    motivo = st.text_area("Motivo de oposición", height=160)
                    doc_oposicion = st.file_uploader("Documento de oposición, si existe", type=["pdf", "docx", "txt"], key="doc_oposicion")
                    enviar_oposicion = st.form_submit_button("Registrar oposición")
                if enviar_oposicion:
                    borrador = generar_borrador_respuesta(record, "Oposición presentada", motivo)
                    update_record(record["codigo"], {
                        "estado": "Oposición presentada",
                        "rol_ultima_accion": "Demandado / Deudor",
                        "respuesta_deudor": "Oposición presentada",
                        "motivo_oposicion": motivo,
                        "fecha_respuesta_deudor": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "accion_recomendada": "Dar traslado al demandante y determinar procedimiento por cuantía",
                        "borrador": borrador,
                    })
                    st.success("Oposición registrada correctamente.")
                    st.text_area("Borrador generado", borrador, height=300)
                    st.session_state["mostrar_oposicion"] = False

else:
    st.info("Elige si accedes como Demandante/Acreedor o como Demandado/Deudor.")

# ============================================================
# DASHBOARD INTERNO DE DEMO
# ============================================================

st.divider()
with st.expander("📊 Ver registros guardados - demo"):
    df = load_data()
    if df.empty:
        st.info("Todavía no hay expedientes registrados.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)
