import streamlit as st
import pandas as pd
import re
from datetime import date, timedelta

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="LexMonitor AI",
    page_icon="⚖️",
    layout="wide"
)

LIMITE_ABREVIADO = 50000

ARTICULOS = {
    "inicio_demanda": {
        "referencia": "Art. 679 CPC",
        "resumen": "El proceso monitorio se inicia mediante demanda o solicitud por escrito."
    },
    "competencia": {
        "referencia": "Art. 678.1 CPC",
        "resumen": "El juez examina de oficio la jurisdicción y competencia del tribunal."
    },
    "subsanacion": {
        "referencia": "Arts. 586.3 y 587.1 CPC",
        "resumen": "Si la demanda presenta defectos formales, se requiere subsanación."
    },
    "archivo_subsanacion": {
        "referencia": "Art. 587.2 CPC",
        "resumen": "Si no se subsana en plazo, procede el archivo definitivo."
    },
    "admision_requerimiento": {
        "referencia": "Arts. 139.1 y 680.1 CPC",
        "resumen": "Si la demanda es admisible, se requiere al deudor para pagar u oponerse en veinte días."
    },
    "pago": {
        "referencia": "Art. 683 CPC",
        "resumen": "Si el demandado paga, se entrega comprobante y se archiva el expediente."
    },
    "silencio": {
        "referencia": "Arts. 681 y 682 CPC",
        "resumen": "Si el deudor no paga ni comparece, procede iniciar ejecución por vía de apremio."
    },
    "oposicion": {
        "referencia": "Art. 684 CPC",
        "resumen": "Si el deudor se opone, el asunto se resuelve en el juicio que corresponda por cuantía."
    },
    "plus_peticion": {
        "referencia": "Art. 684.3 CPC",
        "resumen": "Si hay plus petición, se continúa respecto de la cantidad reconocida como debida."
    },
    "abreviado": {
        "referencia": "Arts. 684.1 y 400.2 CPC",
        "resumen": "Si la cuantía no excede de 50.000 lempiras, continúa por procedimiento abreviado."
    },
    "ordinario": {
        "referencia": "Arts. 424, 685 y 399.2 CPC",
        "resumen": "Si la cuantía excede de 50.000 lempiras, continúa por procedimiento ordinario."
    },
    "sobreseimiento": {
        "referencia": "Art. 685 CPC",
        "resumen": "Si no se interpone la demanda correspondiente tras la oposición, se sobreseen las actuaciones."
    }
}


# ============================================================
# LECTURA DE DOCUMENTOS
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
            celdas = [celda.text.strip() for celda in fila.cells]
            textos.append(" | ".join(celdas))

    return "\n".join(textos)


def leer_pdf(archivo):
    if PyPDF2 is None:
        return "ERROR: instala PyPDF2 con: pip install PyPDF2"

    lector = PyPDF2.PdfReader(archivo)
    textos = []

    for pagina in lector.pages:
        texto = pagina.extract_text()
        if texto:
            textos.append(texto)

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
# EXTRACCIÓN AUTOMÁTICA
# ============================================================

def limpiar_numero(texto_numero):
    if not texto_numero:
        return 0.0

    texto_numero = texto_numero.lower()
    texto_numero = texto_numero.replace("lempiras", "")
    texto_numero = texto_numero.replace("l.", "")
    texto_numero = texto_numero.replace("l ", "")
    texto_numero = texto_numero.strip()

    if "," in texto_numero and "." in texto_numero:
        if texto_numero.find(",") < texto_numero.find("."):
            texto_numero = texto_numero.replace(",", "")
        else:
            texto_numero = texto_numero.replace(".", "").replace(",", ".")
    elif "," in texto_numero:
        partes = texto_numero.split(",")
        if len(partes[-1]) == 2:
            texto_numero = texto_numero.replace(",", ".")
        else:
            texto_numero = texto_numero.replace(",", "")
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
        "documento_firmado_deudor": [
            "firmado por el deudor",
            "firma del deudor",
            "sello del deudor",
            "impronta",
            "marca del deudor",
            "firma electrónica",
            "señal electrónica",
        ],
        "documentos_comerciales_habituales": [
            "factura",
            "facturas",
            "albarán",
            "albaranes",
            "certificación",
            "certificaciones",
            "telegrama",
            "fax",
            "telefax",
            "recibo",
            "estado de cuenta",
            "comprobante",
        ],
        "relacion_anterior_duradera": [
            "relación comercial anterior",
            "relación anterior duradera",
            "relación contractual continuada",
            "contrato marco",
            "contrato de suministro",
            "historial de pedidos",
            "documentos comerciales",
        ],
        "comunidad_propietarios": [
            "certificación de impago",
            "gastos comunes",
            "comunidad de propietarios",
            "inmueble urbano",
        ],
    }

    detectados = {}

    for categoria, palabras in categorias.items():
        coincidencias = [p for p in palabras if p in texto]
        detectados[categoria] = coincidencias

    hay_documento_valido = any(len(v) > 0 for v in detectados.values())

    return hay_documento_valido, detectados
    
def extraer_datos_demanda(texto):
    texto_unido = re.sub(r"\s+", " ", texto)

    demandante = buscar_patron(
        texto_unido,
        [
            r"demandante[:\s]+(.+?)(?:demandado|deudor|contra|frente a|,|\.)",
            r"acreedor[:\s]+(.+?)(?:demandado|deudor|contra|frente a|,|\.)",
            r"promovido por[:\s]+(.+?)(?:contra|frente a|,|\.)",
            r"a instancia de[:\s]+(.+?)(?:contra|frente a|,|\.)",
        ]
    )

    demandado = buscar_patron(
        texto_unido,
        [
            r"demandado[:\s]+(.+?)(?:,|\.)",
            r"deudor[:\s]+(.+?)(?:,|\.)",
            r"contra[:\s]+(.+?)(?:,|\.)",
            r"frente a[:\s]+(.+?)(?:,|\.)",
        ]
    )

    cuantia_txt = buscar_patron(
        texto_unido,
        [
            r"cuantía[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"importe[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"cantidad[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"suma[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"reclama(?:\s+la)?\s+cantidad\s+de\s+(?:L\.?\s*)?([\d\.,]+)",
            r"por\s+importe\s+de\s+(?:L\.?\s*)?([\d\.,]+)",
        ]
    )

    cuantia = limpiar_numero(cuantia_txt)

    palabras_deuda = [
    # Documentos firmados o con señal del deudor
    "firmado por el deudor",
    "firma del deudor",
    "sello del deudor",
    "impronta",
    "marca del deudor",
    "firma electrónica",
    "señal electrónica",

    # Documentos habituales en relaciones comerciales
    "factura",
    "facturas",
    "albarán",
    "albaranes",
    "albarán de entrega",
    "certificación",
    "certificaciones",
    "telegrama",
    "telegramas",
    "fax",
    "telefax",
    "recibo",
    "recibos",
    "estado de cuenta",
    "comprobante",

    # Relación anterior duradera
    "relación comercial anterior",
    "relación anterior duradera",
    "relación contractual continuada",
    "documentos comerciales",
    "contrato marco",
    "contrato de suministro",
    "historial de pedidos",

    # Comunidad de propietarios
    "certificación de impago",
    "gastos comunes",
    "comunidad de propietarios",
    "inmueble urbano",

    # Requisitos de la deuda
    "deuda dineraria",
    "deuda líquida",
    "deuda determinada",
    "deuda vencida",
    "deuda exigible",
    "cantidad determinada",
    "crédito vencido",
    "obligación de pago",
]

    hay_documento_deuda, documentos_art_812 = detectar_documentos_art_812(texto_unido)

    menciona_monitorio = "monitorio" in texto_unido.lower()

    peticion = buscar_patron(
        texto_unido,
        [
            r"solicito[:\s]+(.{20,300})",
            r"suplico[:\s]+(.{20,300})",
            r"pido[:\s]+(.{20,300})",
            r"petición[:\s]+(.{20,300})",
        ]
    )

    datos_faltantes = []

    if not demandante:
        datos_faltantes.append("Demandante / acreedor")
    if not demandado:
        datos_faltantes.append("Demandado / deudor")
    if cuantia <= 0:
        datos_faltantes.append("Cuantía")
    if not hay_documento_deuda:
        datos_faltantes.append("Documento acreditativo de deuda")

    cumple_requisitos_auto = (
        bool(demandante)
        and bool(demandado)
        and cuantia > 0
        and hay_documento_deuda
    )

    return {
        "demandante": demandante,
        "demandado": demandado,
        "cuantia": cuantia,
        "menciona_monitorio": menciona_monitorio,
        "hay_documento_deuda": hay_documento_deuda,
        "documentos_art_812": documentos_art_812,
        "peticion": peticion,
        "hechos_resumidos": texto_unido[:1000],
        "datos_faltantes": datos_faltantes,
        "demanda_presentada_auto": True,
        "tribunal_competente_auto": True,
        "cumple_requisitos_auto": cumple_requisitos_auto,
        "subsanacion_requerida_auto": not cumple_requisitos_auto,
    }


# ============================================================
# PLAZOS
# ============================================================

def fecha_vencimiento(fecha_inicio, plazo_dias):
    if fecha_inicio is None:
        return None
    return fecha_inicio + timedelta(days=plazo_dias)


def dias_restantes(fecha_inicio, plazo_dias):
    if fecha_inicio is None:
        return None
    vencimiento = fecha_vencimiento(fecha_inicio, plazo_dias)
    return (vencimiento - date.today()).days


def calcular_riesgo(dias):
    if dias is None:
        return "Sin plazo"
    if dias < 0:
        return "Vencido"
    if dias <= 3:
        return "Alto"
    if dias <= 7:
        return "Medio"
    return "Bajo"


def pintar_riesgo(riesgo):
    if riesgo == "Bajo":
        return "🟢 Bajo"
    if riesgo == "Medio":
        return "🟡 Medio"
    if riesgo == "Alto":
        return "🔴 Alto"
    if riesgo == "Vencido":
        return "⛔ Vencido"
    return "⚪ Sin plazo"


def completar_plazo(resultado, fecha_inicio):
    plazo = resultado.get("plazo_dias")

    if plazo is not None and fecha_inicio is not None:
        resultado["fecha_vencimiento"] = fecha_vencimiento(fecha_inicio, plazo)
        resultado["dias_restantes"] = dias_restantes(fecha_inicio, plazo)
        resultado["riesgo"] = calcular_riesgo(resultado["dias_restantes"])

    return resultado


# ============================================================
# MOTOR DE REGLAS
# ============================================================

def evaluar_expediente(datos):
    resultado = {
        "estado": "Pendiente de evaluación",
        "articulo": "",
        "resumen_articulo": "",
        "accion": "Revisar expediente",
        "plazo_dias": None,
        "fecha_vencimiento": None,
        "dias_restantes": None,
        "riesgo": "Sin plazo",
        "explicacion": "No se ha podido determinar la fase procesal."
    }

    if not datos["demanda_presentada"]:
        art = ARTICULOS["inicio_demanda"]
        resultado.update({
            "estado": "Sin demanda presentada",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Esperar presentación de demanda o solicitud por escrito",
            "explicacion": "No consta demanda presentada, por lo que el proceso no puede iniciarse."
        })
        return resultado

    if not datos["tribunal_competente"]:
        art = ARTICULOS["competencia"]
        resultado.update({
            "estado": "Rechazo por falta de competencia",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Rechazar de plano y remitir al tribunal competente",
            "explicacion": "El tribunal debe examinar su competencia antes de continuar."
        })
        return resultado

    if not datos["cumple_requisitos"]:
        if datos["subsanacion_requerida"] and not datos["demanda_subsanada"]:
            dias = dias_restantes(datos["fecha_requerimiento_subsanacion"], 5)

            if dias is not None and dias < 0:
                art = ARTICULOS["archivo_subsanacion"]
                resultado.update({
                    "estado": "Archivo definitivo por falta de subsanación",
                    "articulo": art["referencia"],
                    "resumen_articulo": art["resumen"],
                    "accion": "Archivar definitivamente el expediente",
                    "explicacion": "Se requirió subsanación y el plazo ha vencido sin corrección."
                })
                return resultado

            art = ARTICULOS["subsanacion"]
            resultado.update({
                "estado": "Subsanación pendiente",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Esperar o requerir subsanación del demandante",
                "plazo_dias": 5,
                "explicacion": "La demanda tiene defectos o datos faltantes. Procede subsanar."
            })
            return completar_plazo(resultado, datos["fecha_requerimiento_subsanacion"])

        art = ARTICULOS["subsanacion"]
        resultado.update({
            "estado": "Demanda incompleta: requerir subsanación",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Emitir requerimiento de subsanación por cinco días",
            "plazo_dias": 5,
            "explicacion": "La demanda no cumple los requisitos mínimos para su admisión."
        })
        return completar_plazo(resultado, datos["fecha_requerimiento_subsanacion"])

    if datos["cumple_requisitos"] and not datos["demanda_admitida"]:
        art = ARTICULOS["admision_requerimiento"]
        resultado.update({
            "estado": "Demanda lista para admisión",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Admitir la demanda y requerir al deudor para pagar u oponerse",
            "explicacion": "La demanda contiene datos básicos suficientes y documento de deuda. Procede admisión inicial."
        })
        return resultado

    if datos["demanda_admitida"] and not datos["deudor_notificado"]:
        art = ARTICULOS["admision_requerimiento"]
        resultado.update({
            "estado": "Pendiente de notificación al deudor",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Notificar al deudor el requerimiento de pago u oposición",
            "explicacion": "La demanda ya está admitida, pero falta notificar al deudor."
        })
        return resultado

    if datos["deudor_notificado"]:

        if datos["demandado_pago"]:
            art = ARTICULOS["pago"]
            resultado.update({
                "estado": "Pago realizado y archivo",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Entregar comprobante de pago y archivar actuaciones",
                "explicacion": "El demandado ha pagado la cantidad reclamada."
            })
            return resultado

        if datos["demandado_oposicion"]:
            return evaluar_oposicion(datos)

        dias = dias_restantes(datos["fecha_notificacion_deudor"], 20)

        if dias is not None and dias < 0:
            art = ARTICULOS["silencio"]
            resultado.update({
                "estado": "Procede ejecución por vía de apremio",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Dictar auto iniciando ejecución por vía de apremio",
                "explicacion": "El deudor fue notificado y han pasado veinte días sin pago ni oposición."
            })
            return resultado

        art = ARTICULOS["admision_requerimiento"]
        resultado.update({
            "estado": "Esperando pago u oposición del deudor",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Controlar vencimiento del plazo de veinte días",
            "plazo_dias": 20,
            "explicacion": "El deudor está dentro del plazo para pagar o formular oposición."
        })
        return completar_plazo(resultado, datos["fecha_notificacion_deudor"])

    return resultado


def evaluar_oposicion(datos):
    resultado = {
        "estado": "Oposición presentada",
        "articulo": "",
        "resumen_articulo": "",
        "accion": "",
        "plazo_dias": None,
        "fecha_vencimiento": None,
        "dias_restantes": None,
        "riesgo": "Sin plazo",
        "explicacion": ""
    }

    if datos["plus_peticion"]:
        art = ARTICULOS["plus_peticion"]
        resultado.update({
            "estado": "Oposición por plus petición",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Continuar respecto de la cantidad reconocida como debida",
            "explicacion": "La oposición sostiene que se reclama más cantidad de la debida."
        })
        return resultado

    if not datos["traslado_demandante"]:
        art = ARTICULOS["oposicion"]
        resultado.update({
            "estado": "Oposición pendiente de traslado",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Admitir oposición y dar traslado al demandante",
            "explicacion": "El demandado ha presentado oposición y debe darse traslado al demandante."
        })
        return resultado

    if datos["cuantia"] <= LIMITE_ABREVIADO:
        art = ARTICULOS["abreviado"]
        resultado.update({
            "estado": "Procedimiento abreviado",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Convocar audiencia de procedimiento abreviado",
            "explicacion": "La cuantía no excede de 50.000 lempiras."
        })
        return resultado

    if datos["cuantia"] > LIMITE_ABREVIADO:
        if not datos["demanda_ordinaria_presentada"]:
            dias = dias_restantes(datos["fecha_traslado_demandante"], 30)

            if dias is not None and dias < 0:
                art = ARTICULOS["sobreseimiento"]
                resultado.update({
                    "estado": "Sobreseimiento por falta de demanda ordinaria",
                    "articulo": art["referencia"],
                    "resumen_articulo": art["resumen"],
                    "accion": "Sobreseer actuaciones y condenar en costas al acreedor",
                    "explicacion": "Tras la oposición, la cuantía exige ordinario, pero no consta demanda ordinaria en plazo."
                })
                return resultado

            art = ARTICULOS["ordinario"]
            resultado.update({
                "estado": "Pendiente de demanda ordinaria",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Esperar presentación de demanda ordinaria por el acreedor",
                "plazo_dias": 30,
                "explicacion": "Al exceder la cuantía de 50.000 lempiras, debe presentarse demanda ordinaria."
            })
            return completar_plazo(resultado, datos["fecha_traslado_demandante"])

        art = ARTICULOS["ordinario"]
        resultado.update({
            "estado": "Procedimiento ordinario",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Continuar por los trámites del procedimiento ordinario",
            "explicacion": "El acreedor ya presentó demanda ordinaria tras la oposición."
        })
        return resultado

    return resultado


# ============================================================
# BORRADORES
# ============================================================

def generar_borrador(datos, resultado):
    expediente = datos.get("expediente", "Sin número")
    demandante = datos.get("demandante", "parte demandante")
    demandado = datos.get("demandado", "parte demandada")
    cuantia = datos.get("cuantia", 0.0)

    encabezado = f"""
EXPEDIENTE: {expediente}
DEMANDANTE: {demandante}
DEMANDADO: {demandado}
CUANTÍA: {cuantia:,.2f} lempiras

ESTADO DETECTADO: {resultado['estado']}
REFERENCIA LEGAL: {resultado['articulo']}
RESUMEN NORMATIVO: {resultado['resumen_articulo']}

"""

    estado = resultado["estado"].lower()

    if "lista para admisión" in estado:
        cuerpo = f"""
AUTO DE ADMISIÓN DE DEMANDA MONITORIA

Visto el escrito presentado por {demandante} frente a {demandado}, y apreciándose que la solicitud contiene los datos básicos necesarios y documentación acreditativa de la deuda reclamada, procede admitir la demanda monitoria.

En consecuencia, se acuerda requerir al deudor para que, en el plazo legal de veinte días, pague la cantidad reclamada o comparezca formulando oposición.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "subsanación" in estado or "incompleta" in estado:
        cuerpo = f"""
AUTO DE REQUERIMIENTO DE SUBSANACIÓN

Visto el escrito presentado, y apreciándose que la solicitud inicial no reúne todos los requisitos necesarios para su admisión, procede requerir a la parte actora para que subsane los defectos advertidos en el plazo legal de cinco días.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "pendiente de notificación" in estado:
        cuerpo = f"""
DILIGENCIA DE NOTIFICACIÓN AL DEUDOR

Admitida la demanda monitoria, procede notificar al deudor {demandado}, requiriéndole para que pague o formule oposición en el plazo legal.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "esperando pago" in estado:
        cuerpo = f"""
INFORME DE CONTROL DE PLAZO

Consta que el deudor ha sido notificado y se encuentra abierto el plazo para pagar o formular oposición.

DÍAS RESTANTES:
{resultado.get('dias_restantes', '-')}

FECHA DE VENCIMIENTO:
{resultado.get('fecha_vencimiento', '-')}

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "pago realizado" in estado:
        cuerpo = f"""
DILIGENCIA DE PAGO Y ARCHIVO

Constando que el demandado ha procedido al pago de la cantidad reclamada, procede entregar el correspondiente comprobante y acordar el archivo de las actuaciones.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "ejecución" in estado or "apremio" in estado:
        cuerpo = f"""
AUTO DESPACHANDO EJECUCIÓN POR VÍA DE APREMIO

Visto que ha transcurrido el plazo legal conferido al deudor sin que conste pago ni oposición, procede iniciar la ejecución por vía de apremio por la cantidad reclamada.

Se genera este borrador para revisión y validación por el órgano competente.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "oposición pendiente" in estado:
        cuerpo = f"""
AUTO DE ADMISIÓN DE OPOSICIÓN Y TRASLADO

Presentado escrito de oposición por el demandado, procede admitirlo y dar traslado a la parte demandante para que actúe conforme corresponda.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "abreviado" in estado:
        cuerpo = f"""
PROVIDENCIA DE CONTINUACIÓN POR PROCEDIMIENTO ABREVIADO

Formulada oposición y no excediendo la cuantía de {LIMITE_ABREVIADO:,.2f} lempiras, procede continuar la tramitación por el procedimiento abreviado.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "ordinario" in estado:
        cuerpo = f"""
PROVIDENCIA DE CONTINUACIÓN POR PROCEDIMIENTO ORDINARIO

Formulada oposición y excediendo la cuantía de {LIMITE_ABREVIADO:,.2f} lempiras, procede continuar por los trámites del procedimiento ordinario.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    if "sobreseimiento" in estado:
        cuerpo = f"""
AUTO DE SOBRESEIMIENTO

Constando que no se ha presentado la demanda correspondiente dentro del plazo legal tras la oposición, procede acordar el sobreseimiento de las actuaciones.

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
        return encabezado + cuerpo

    cuerpo = f"""
INFORME DE ESTADO PROCESAL

{resultado['explicacion']}

PRÓXIMA ACTUACIÓN:
{resultado['accion']}
"""
    return encabezado + cuerpo


# ============================================================
# EJEMPLOS
# ============================================================

def cargar_ejemplos():
    hoy = date.today()

    return [
        {
            "expediente": "MON-001",
            "demandante": "Banco Atlántico S.A.",
            "demandado": "Carlos Mejía Rodríguez",
            "cuantia": 35000.0,
            "demanda_presentada": True,
            "tribunal_competente": True,
            "cumple_requisitos": True,
            "subsanacion_requerida": False,
            "demanda_subsanada": False,
            "demanda_admitida": False,
            "deudor_notificado": False,
            "demandado_pago": False,
            "demandado_oposicion": False,
            "plus_peticion": False,
            "traslado_demandante": False,
            "demanda_ordinaria_presentada": False,
            "fecha_requerimiento_subsanacion": hoy,
            "fecha_notificacion_deudor": hoy,
            "fecha_traslado_demandante": hoy,
            "hechos_resumidos": "Demanda monitoria de ejemplo con factura y contrato."
        }
    ]


# ============================================================
# SESSION STATE
# ============================================================

if "expedientes" not in st.session_state:
    st.session_state.expedientes = cargar_ejemplos()

if "datos_extraidos" not in st.session_state:
    st.session_state.datos_extraidos = None

if "texto_subido" not in st.session_state:
    st.session_state.texto_subido = ""


# ============================================================
# INTERFAZ
# ============================================================

st.title("⚖️ LexMonitor AI")
st.subheader("Lectura automática y monitorización del proceso monitorio")

st.warning(
    "Prototipo para hackathon. El sistema no dicta sentencia de forma autónoma: "
    "extrae datos, aplica reglas procesales y genera borradores revisables."
)

st.sidebar.title("Menú")
pagina = st.sidebar.radio(
    "Selecciona pantalla",
    [
        "1. Subir demanda",
        "2. Dashboard",
        "3. Detalle de expediente",
        "4. Mapa de reglas"
    ]
)


# ============================================================
# SUBIR DEMANDA
# ============================================================

if pagina == "1. Subir demanda":
    st.header("📄 Subir demanda o solicitud monitoria")

    archivos = st.file_uploader(
    "Sube la demanda y los documentos acreditativos de la deuda",
    type=["txt", "docx", "pdf"],
    accept_multiple_files=True
)

if archivos:
    textos = []
    nombres_archivos = []

    for archivo in archivos:
        texto_archivo = leer_documento(archivo)
        textos.append(f"\n\n--- DOCUMENTO: {archivo.name} ---\n{texto_archivo}")
        nombres_archivos.append(archivo.name)

    texto = "\n".join(textos)
    st.session_state.texto_subido = texto

    datos_extraidos = extraer_datos_demanda(texto)
    st.session_state.datos_extraidos = datos_extraidos

    st.subheader("Documentos subidos")
    for nombre in nombres_archivos:
        st.write(f"📎 {nombre}")

    st.subheader("Texto leído de los documentos")
    st.text_area("Contenido detectado", texto, height=260)

    st.subheader("Datos extraídos automáticamente")
    col1, col2, col3 = st.columns(3)

    col1.metric("Demandante", datos_extraidos["demandante"] or "No detectado")
    col2.metric("Demandado", datos_extraidos["demandado"] or "No detectado")
    col3.metric("Cuantía", f"{datos_extraidos['cuantia']:,.2f} L")

    st.write(f"**Menciona proceso monitorio:** {'Sí' if datos_extraidos['menciona_monitorio'] else 'No'}")
    st.write(f"**Detecta documento válido art. 812 LEC:** {'Sí' if datos_extraidos['hay_documento_deuda'] else 'No'}")
    st.subheader("Documentación detectada según art. 812 LEC")

    docs_812 = datos_extraidos.get("documentos_art_812", {})
    
    for categoria, coincidencias in docs_812.items():
        if coincidencias:
            st.success(f"{categoria}: {', '.join(coincidencias)}")
        else:
            st.info(f"{categoria}: no detectado")

    if datos_extraidos["datos_faltantes"]:
        st.warning("Datos faltantes o dudosos: " + ", ".join(datos_extraidos["datos_faltantes"]))
    else:
        st.success("El sistema detecta demanda y documentación mínima para iniciar el flujo.")
        st.session_state.texto_subido = texto

        datos_extraidos = extraer_datos_demanda(texto)
        st.session_state.datos_extraidos = datos_extraidos

        st.subheader("Texto leído del documento")
        st.text_area("Contenido detectado", texto, height=220)

        st.subheader("Datos extraídos automáticamente")
        col1, col2, col3 = st.columns(3)

        col1.metric("Demandante", datos_extraidos["demandante"] or "No detectado")
        col2.metric("Demandado", datos_extraidos["demandado"] or "No detectado")
        col3.metric("Cuantía", f"{datos_extraidos['cuantia']:,.2f} L")

        st.write(f"**Menciona proceso monitorio:** {'Sí' if datos_extraidos['menciona_monitorio'] else 'No'}")
        st.write(f"**Detecta documento de deuda:** {'Sí' if datos_extraidos['hay_documento_deuda'] else 'No'}")

        if datos_extraidos["datos_faltantes"]:
            st.warning("Datos faltantes o dudosos: " + ", ".join(datos_extraidos["datos_faltantes"]))
        else:
            st.success("El sistema detecta los datos mínimos para iniciar el flujo.")

    st.divider()

    st.header("✅ Validación mínima antes de aplicar el flujo")

    datos = st.session_state.datos_extraidos or {
        "demandante": "",
        "demandado": "",
        "cuantia": 0.0,
        "hay_documento_deuda": False,
        "menciona_monitorio": False,
        "peticion": "",
        "hechos_resumidos": "",
        "datos_faltantes": [],
        "demanda_presentada_auto": True,
        "tribunal_competente_auto": True,
        "cumple_requisitos_auto": False,
        "subsanacion_requerida_auto": True,
    }

    cumple_requisitos_auto = (
        bool(datos.get("demandante"))
        and bool(datos.get("demandado"))
        and datos.get("cuantia", 0) > 0
        and datos.get("hay_documento_deuda", False)
    )

    with st.form("validacion_auto"):
        c1, c2, c3 = st.columns(3)

        expediente = c1.text_input(
            "Número de expediente",
            value=f"MON-{len(st.session_state.expedientes) + 1:03d}"
        )

        demandante = c2.text_input(
            "Demandante / acreedor",
            value=datos.get("demandante", "")
        )

        demandado = c3.text_input(
            "Demandado / deudor",
            value=datos.get("demandado", "")
        )

        cuantia = st.number_input(
            "Cuantía reclamada en lempiras",
            min_value=0.0,
            value=float(datos.get("cuantia", 0.0)),
            step=1000.0
        )

        st.subheader("Resultado automático de la lectura inicial")

        c1, c2, c3, c4 = st.columns(4)

        demanda_presentada = c1.checkbox(
            "Demanda presentada",
            value=True
        )

        tribunal_competente = c2.checkbox(
            "Tribunal competente",
            value=True
        )

        cumple_requisitos = c3.checkbox(
            "Cumple requisitos básicos",
            value=cumple_requisitos_auto
        )

        subsanacion_requerida = c4.checkbox(
            "Requiere subsanación",
            value=not cumple_requisitos_auto
        )

        st.caption(
            "Estos campos se rellenan automáticamente según la lectura del documento, "
            "pero se pueden corregir manualmente."
        )

        st.subheader("Seguimiento procesal posterior")

        c1, c2, c3 = st.columns(3)

        demanda_admitida = c1.checkbox(
            "Demanda ya admitida",
            value=False
        )

        deudor_notificado = c2.checkbox(
            "Deudor ya notificado",
            value=False
        )

        fecha_notificacion_deudor = c3.date_input(
            "Fecha notificación al deudor",
            value=date.today()
        )

        c1, c2, c3 = st.columns(3)

        demandado_pago = c1.checkbox(
            "El demandado pagó",
            value=False
        )

        demandado_oposicion = c2.checkbox(
            "El demandado presentó oposición",
            value=False
        )

        plus_peticion = c3.checkbox(
            "Oposición por plus petición",
            value=False
        )

        c1, c2, c3 = st.columns(3)

        traslado_demandante = c1.checkbox(
            "Traslado al demandante",
            value=False
        )

        fecha_traslado_demandante = c2.date_input(
            "Fecha traslado al demandante",
            value=date.today()
        )

        demanda_ordinaria_presentada = c3.checkbox(
            "Demanda ordinaria presentada",
            value=False
        )

        c1, c2 = st.columns(2)

        demanda_subsanada = c1.checkbox(
            "Demanda subsanada",
            value=False
        )

        fecha_requerimiento_subsanacion = c2.date_input(
            "Fecha requerimiento subsanación",
            value=date.today()
        )

        hechos_resumidos = st.text_area(
            "Resumen de hechos detectado",
            value=datos.get("hechos_resumidos", ""),
            height=140
        )

        guardar = st.form_submit_button("Aplicar flujo y guardar expediente")

    if guardar:
        expediente_datos = {
            "expediente": expediente,
            "demandante": demandante,
            "demandado": demandado,
            "cuantia": cuantia,
            "demanda_presentada": demanda_presentada,
            "tribunal_competente": tribunal_competente,
            "cumple_requisitos": cumple_requisitos,
            "subsanacion_requerida": subsanacion_requerida,
            "demanda_subsanada": demanda_subsanada,
            "demanda_admitida": demanda_admitida,
            "deudor_notificado": deudor_notificado,
            "demandado_pago": demandado_pago,
            "demandado_oposicion": demandado_oposicion,
            "plus_peticion": plus_peticion,
            "traslado_demandante": traslado_demandante,
            "demanda_ordinaria_presentada": demanda_ordinaria_presentada,
            "fecha_requerimiento_subsanacion": fecha_requerimiento_subsanacion,
            "fecha_notificacion_deudor": fecha_notificacion_deudor,
            "fecha_traslado_demandante": fecha_traslado_demandante,
            "hechos_resumidos": hechos_resumidos,
            "texto_original": st.session_state.texto_subido,
        }

        resultado = evaluar_expediente(expediente_datos)
        borrador = generar_borrador(expediente_datos, resultado)

        st.session_state.expedientes.append(expediente_datos)

        st.success("Expediente guardado y evaluado.")

        st.subheader("Resultado automático del flujo")
        c1, c2, c3 = st.columns(3)
        c1.metric("Estado", resultado["estado"])
        c2.metric("Riesgo", pintar_riesgo(resultado["riesgo"]))
        c3.metric("Días restantes", resultado["dias_restantes"] if resultado["dias_restantes"] is not None else "-")

        st.write(f"**Artículo aplicable:** {resultado['articulo']}")
        st.write(f"**Resumen del artículo:** {resultado['resumen_articulo']}")
        st.write(f"**Explicación aplicada al caso:** {resultado['explicacion']}")
        st.write(f"**Próxima actuación:** {resultado['accion']}")

        st.subheader("Borrador generado")
        st.text_area("Borrador revisable", borrador, height=420)

        st.download_button(
            "Descargar borrador TXT",
            data=borrador,
            file_name=f"borrador_{expediente}.txt",
            mime="text/plain"
        )


# ============================================================
# DASHBOARD
# ============================================================

elif pagina == "2. Dashboard":
    st.header("📊 Dashboard de expedientes")

    filas = []

    for exp in st.session_state.expedientes:
        res = evaluar_expediente(exp)

        filas.append({
            "Expediente": exp["expediente"],
            "Demandante": exp["demandante"],
            "Demandado": exp["demandado"],
            "Cuantía": f"{exp['cuantia']:,.2f} L",
            "Estado": res["estado"],
            "Artículo": res["articulo"],
            "Vencimiento": res["fecha_vencimiento"] if res["fecha_vencimiento"] else "-",
            "Días restantes": res["dias_restantes"] if res["dias_restantes"] is not None else "-",
            "Riesgo": pintar_riesgo(res["riesgo"]),
            "Próxima actuación": res["accion"],
        })

    df = pd.DataFrame(filas)
    st.dataframe(df, use_container_width=True, hide_index=True)

    total = len(filas)
    vencidos = sum(1 for f in filas if "Vencido" in f["Riesgo"])
    alto = sum(1 for f in filas if "Alto" in f["Riesgo"])
    oposiciones = sum(1 for exp in st.session_state.expedientes if exp.get("demandado_oposicion"))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total expedientes", total)
    c2.metric("Vencidos", vencidos)
    c3.metric("Riesgo alto", alto)
    c4.metric("Con oposición", oposiciones)


# ============================================================
# DETALLE
# ============================================================

elif pagina == "3. Detalle de expediente":
    st.header("🔎 Detalle de expediente")

    opciones = [exp["expediente"] for exp in st.session_state.expedientes]

    if not opciones:
        st.warning("No hay expedientes.")
    else:
        seleccionado = st.selectbox("Selecciona expediente", opciones)

        exp = next(e for e in st.session_state.expedientes if e["expediente"] == seleccionado)
        res = evaluar_expediente(exp)
        borrador = generar_borrador(exp, res)

        c1, c2, c3 = st.columns(3)
        c1.metric("Estado", res["estado"])
        c2.metric("Riesgo", pintar_riesgo(res["riesgo"]))
        c3.metric("Días restantes", res["dias_restantes"] if res["dias_restantes"] is not None else "-")

        st.subheader("Datos del expediente")
        st.json({
            "expediente": exp["expediente"],
            "demandante": exp["demandante"],
            "demandado": exp["demandado"],
            "cuantia": exp["cuantia"],
            "demanda_presentada": exp["demanda_presentada"],
            "cumple_requisitos": exp["cumple_requisitos"],
            "demanda_admitida": exp["demanda_admitida"],
            "deudor_notificado": exp["deudor_notificado"],
            "demandado_pago": exp["demandado_pago"],
            "demandado_oposicion": exp["demandado_oposicion"],
        })

        st.subheader("Evaluación procesal")
        st.write(f"**Estado:** {res['estado']}")
        st.write(f"**Artículo:** {res['articulo']}")
        st.write(f"**Resumen artículo:** {res['resumen_articulo']}")
        st.write(f"**Explicación:** {res['explicacion']}")
        st.write(f"**Próxima actuación:** {res['accion']}")

        if exp.get("hechos_resumidos"):
            st.subheader("Hechos resumidos")
            st.write(exp["hechos_resumidos"])

        st.subheader("Borrador generado")
        st.text_area("Borrador", borrador, height=450)

        st.download_button(
            "Descargar borrador TXT",
            data=borrador,
            file_name=f"borrador_{exp['expediente']}.txt",
            mime="text/plain"
        )


# ============================================================
# MAPA DE REGLAS
# ============================================================

elif pagina == "4. Mapa de reglas":
    st.header("🧠 Mapa de reglas")

    st.markdown(
        """
        ### Flujo lógico usado por el sistema

        1. Lee la demanda.
        2. Extrae demandante, demandado, cuantía y documentos de deuda.
        3. Decide si la demanda parece completa o requiere subsanación.
        4. Si está completa, propone admisión y requerimiento al deudor.
        5. Si el deudor paga, archiva.
        6. Si no paga ni se opone en 20 días, propone ejecución.
        7. Si se opone, decide abreviado u ordinario según cuantía.
        """
    )

    st.code(
        """
if demanda_presentada and tribunal_competente and cumple_requisitos:
    if not demanda_admitida:
        estado = "Demanda lista para admisión"

elif not cumple_requisitos:
    estado = "Requerir subsanación"

elif deudor_notificado and not pago and not oposicion and plazo_20_dias_vencido:
    estado = "Ejecución por vía de apremio"

elif oposicion and cuantia <= 50000:
    estado = "Procedimiento abreviado"

elif oposicion and cuantia > 50000:
    estado = "Procedimiento ordinario"
        """,
        language="python"
    )

    st.subheader("Artículos usados")
    tabla = []

    for clave, valor in ARTICULOS.items():
        tabla.append({
            "Regla": clave,
            "Referencia": valor["referencia"],
            "Resumen": valor["resumen"]
        })

    st.dataframe(pd.DataFrame(tabla), use_container_width=True, hide_index=True)
