import streamlit as st
import pandas as pd
import re
from datetime import date, timedelta
from io import BytesIO

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None


# ============================================================
# CONFIGURACIÓN GENERAL
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
        "resumen": "Si la demanda presenta defectos formales, se requiere subsanación en plazo legal."
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
        return "ERROR: Falta instalar python-docx. Ejecuta: pip install python-docx"

    doc = Document(archivo)
    textos = []

    for p in doc.paragraphs:
        if p.text.strip():
            textos.append(p.text)

    for tabla in doc.tables:
        for fila in tabla.rows:
            celdas = [celda.text.strip() for celda in fila.cells]
            textos.append(" | ".join(celdas))

    return "\n".join(textos)


def leer_pdf(archivo):
    if PyPDF2 is None:
        return "ERROR: Falta instalar PyPDF2. Ejecuta: pip install PyPDF2"

    lector = PyPDF2.PdfReader(archivo)
    textos = []

    for pagina in lector.pages:
        texto = pagina.extract_text()
        if texto:
            textos.append(texto)

    return "\n".join(textos)


def leer_documento_subido(archivo):
    nombre = archivo.name.lower()

    if nombre.endswith(".txt"):
        return leer_txt(archivo)

    if nombre.endswith(".docx"):
        return leer_docx(archivo)

    if nombre.endswith(".pdf"):
        return leer_pdf(archivo)

    return ""


# ============================================================
# EXTRACCIÓN DE DATOS DESDE LA DEMANDA
# ============================================================

def limpiar_numero(texto_numero):
    """
    Convierte textos como:
    50,000.00
    50.000,00
    50000
    en float.
    """
    if not texto_numero:
        return 0.0

    texto_numero = texto_numero.replace("L", "")
    texto_numero = texto_numero.replace("lempiras", "")
    texto_numero = texto_numero.strip()

    if "," in texto_numero and "." in texto_numero:
        # Caso 50,000.00
        if texto_numero.find(",") < texto_numero.find("."):
            texto_numero = texto_numero.replace(",", "")
        # Caso 50.000,00
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


def extraer_datos_demanda(texto):
    """
    Extractor básico para hackathon.
    La idea es convertir el documento en datos estructurados.
    Después estos datos entran en el motor de reglas.
    """

    texto_limpio = re.sub(r"\s+", " ", texto)

    demandante = buscar_patron(
        texto_limpio,
        [
            r"demandante[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:demandado|contra|,|\.|$)",
            r"acreedor[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:deudor|contra|,|\.|$)",
            r"promovido por[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:contra|frente a|,|\.|$)",
            r"a instancia de[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:contra|frente a|,|\.|$)"
        ]
    )

    demandado = buscar_patron(
        texto_limpio,
        [
            r"demandado[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:,|\.|$)",
            r"deudor[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:,|\.|$)",
            r"contra[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:,|\.|$)",
            r"frente a[:\s]+([A-ZÁÉÍÓÚÑa-záéíóúñ\s\.]+?)(?:,|\.|$)"
        ]
    )

    cuantia_texto = buscar_patron(
        texto_limpio,
        [
            r"cuantía[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"importe[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"cantidad[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"suma[:\s]+(?:de\s*)?(?:L\.?\s*)?([\d\.,]+)",
            r"reclama(?:\s+la)?\s+cantidad\s+de\s+(?:L\.?\s*)?([\d\.,]+)"
        ]
    )

    cuantia = limpiar_numero(cuantia_texto)

    palabras_documento_deuda = [
        "factura",
        "contrato",
        "recibo",
        "pagaré",
        "letra de cambio",
        "certificación",
        "documento de deuda",
        "documento justificativo",
        "estado de cuenta",
        "comprobante",
        "obligación de pago"
    ]

    hay_documento_deuda = any(
        palabra.lower() in texto_limpio.lower()
        for palabra in palabras_documento_deuda
    )

    menciona_monitorio = "monitorio" in texto_limpio.lower()

    peticion = buscar_patron(
        texto_limpio,
        [
            r"solicito[:\s]+(.{20,250})",
            r"pido[:\s]+(.{20,250})",
            r"suplico[:\s]+(.{20,250})",
            r"petición[:\s]+(.{20,250})"
        ]
    )

    hechos_resumidos = texto_limpio[:900]

    datos = {
        "demandante": demandante,
        "demandado": demandado,
        "cuantia": cuantia,
        "menciona_monitorio": menciona_monitorio,
        "hay_documento_de_deuda": hay_documento_deuda,
        "peticion": peticion,
        "hechos_resumidos": hechos_resumidos,
        "datos_faltantes": []
    }

    if not demandante:
        datos["datos_faltantes"].append("Demandante / acreedor")
    if not demandado:
        datos["datos_faltantes"].append("Demandado / deudor")
    if cuantia == 0:
        datos["datos_faltantes"].append("Cuantía")
    if not hay_documento_deuda:
        datos["datos_faltantes"].append("Documento acreditativo de la deuda")

    return datos


# ============================================================
# PLAZOS Y RIESGOS
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


# ============================================================
# MOTOR DE REGLAS DEL PROCESO MONITORIO
# ============================================================

def completar_plazo(resultado, fecha_inicio):
    plazo = resultado.get("plazo_dias")

    if plazo is not None and fecha_inicio is not None:
        resultado["fecha_vencimiento"] = fecha_vencimiento(fecha_inicio, plazo)
        resultado["dias_restantes"] = dias_restantes(fecha_inicio, plazo)
        resultado["riesgo"] = calcular_riesgo(resultado["dias_restantes"])

    return resultado


def evaluar_expediente(datos):
    """
    Esta función es el corazón del sistema.
    Recibe datos estructurados y devuelve la fase procesal.
    """

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

    # 1. No hay demanda
    if not datos["demanda_presentada"]:
        art = ARTICULOS["inicio_demanda"]
        resultado.update({
            "estado": "Sin demanda presentada",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Esperar presentación de demanda o solicitud por escrito",
            "explicacion": "El proceso no puede avanzar porque todavía no consta demanda presentada."
        })
        return resultado

    # 2. Competencia
    if not datos["tribunal_competente"]:
        art = ARTICULOS["competencia"]
        resultado.update({
            "estado": "Rechazo por falta de competencia",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Rechazar de plano y remitir al tribunal competente",
            "explicacion": "El tribunal debe examinar de oficio su competencia. Si no es competente, no debe continuar con el fondo."
        })
        return resultado

    # 3. Requisitos de admisibilidad
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
                    "explicacion": "Se requirió subsanación y el plazo ha vencido sin que conste subsanación."
                })
                return resultado

            art = ARTICULOS["subsanacion"]
            resultado.update({
                "estado": "Subsanación pendiente",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Esperar subsanación del demandante",
                "plazo_dias": 5,
                "explicacion": "La demanda presenta defectos formales y el demandante está dentro del plazo para corregirlos."
            })
            return completar_plazo(resultado, datos["fecha_requerimiento_subsanacion"])

        art = ARTICULOS["subsanacion"]
        resultado.update({
            "estado": "Demanda incompleta: requerir subsanación",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Emitir requerimiento de subsanación por cinco días",
            "plazo_dias": 5,
            "explicacion": "La demanda no cumple todos los requisitos de admisibilidad. Procede requerir subsanación."
        })
        return completar_plazo(resultado, datos["fecha_requerimiento_subsanacion"])

    # 4. Demanda completa, pero no admitida todavía
    if datos["cumple_requisitos"] and not datos["demanda_admitida"]:
        art = ARTICULOS["admision_requerimiento"]
        resultado.update({
            "estado": "Demanda lista para admisión",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Admitir la demanda y requerir al deudor para pagar u oponerse",
            "explicacion": "La demanda cumple los requisitos y el tribunal es competente. Procede admitirla."
        })
        return resultado

    # 5. Admitida, pero sin notificar al deudor
    if datos["demanda_admitida"] and not datos["deudor_notificado"]:
        art = ARTICULOS["admision_requerimiento"]
        resultado.update({
            "estado": "Pendiente de notificación al deudor",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Notificar al deudor el requerimiento de pago u oposición",
            "explicacion": "La demanda ya está admitida, pero todavía debe notificarse formalmente al deudor."
        })
        return resultado

    # 6. Deudor notificado
    if datos["deudor_notificado"]:

        # 6.1 Pago
        if datos["demandado_pago"]:
            art = ARTICULOS["pago"]
            resultado.update({
                "estado": "Pago realizado y archivo",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Entregar comprobante de pago y archivar actuaciones",
                "explicacion": "El demandado ha pagado la deuda reclamada. El proceso monitorio finaliza por pago."
            })
            return resultado

        # 6.2 Oposición
        if datos["demandado_oposicion"]:
            return evaluar_oposicion(datos)

        # 6.3 Silencio del deudor
        dias = dias_restantes(datos["fecha_notificacion_deudor"], 20)

        if dias is not None and dias < 0:
            art = ARTICULOS["silencio"]
            resultado.update({
                "estado": "Procede ejecución por vía de apremio",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Dictar auto iniciando ejecución por vía de apremio",
                "explicacion": "El deudor fue notificado y ha vencido el plazo de veinte días sin pago ni oposición."
            })
            return resultado

        art = ARTICULOS["admision_requerimiento"]
        resultado.update({
            "estado": "Esperando pago u oposición del deudor",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Controlar vencimiento del plazo de veinte días",
            "plazo_dias": 20,
            "explicacion": "El deudor está dentro del plazo legal para pagar o formular oposición."
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

    # Plus petición
    if datos["plus_peticion"]:
        art = ARTICULOS["plus_peticion"]
        resultado.update({
            "estado": "Oposición por plus petición",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Continuar respecto de la cantidad reconocida como debida",
            "explicacion": "La oposición se basa en que se reclama una cantidad superior a la debida."
        })
        return resultado

    # Todavía no se dio traslado al demandante
    if not datos["traslado_demandante"]:
        art = ARTICULOS["oposicion"]
        resultado.update({
            "estado": "Oposición pendiente de traslado",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Admitir oposición y dar traslado al demandante",
            "explicacion": "El demandado se ha opuesto. El siguiente paso es dar traslado al demandante."
        })
        return resultado

    # Según cuantía: abreviado u ordinario
    if datos["cuantia"] <= LIMITE_ABREVIADO:
        art = ARTICULOS["abreviado"]
        resultado.update({
            "estado": "Procedimiento abreviado",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Convocar audiencia de procedimiento abreviado",
            "explicacion": "La cuantía no excede de 50.000 lempiras, por lo que corresponde procedimiento abreviado."
        })
        return resultado

    # Cuantía superior: ordinario
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
                    "explicacion": "La cuantía exige procedimiento ordinario, pero no consta demanda ordinaria dentro del plazo."
                })
                return resultado

            art = ARTICULOS["ordinario"]
            resultado.update({
                "estado": "Pendiente de demanda ordinaria",
                "articulo": art["referencia"],
                "resumen_articulo": art["resumen"],
                "accion": "Esperar presentación de demanda ordinaria por el acreedor",
                "plazo_dias": 30,
                "explicacion": "Al exceder la cuantía de 50.000 lempiras, el acreedor debe presentar demanda ordinaria."
            })
            return completar_plazo(resultado, datos["fecha_traslado_demandante"])

        art = ARTICULOS["ordinario"]
        resultado.update({
            "estado": "Procedimiento ordinario",
            "articulo": art["referencia"],
            "resumen_articulo": art["resumen"],
            "accion": "Continuar por los trámites del procedimiento ordinario",
            "explicacion": "El acreedor ha presentado demanda ordinaria tras la oposición."
        })
        return resultado

    return resultado


# ============================================================
# GENERACIÓN DE BORRADORES
# ============================================================

def generar_borrador(datos, resultado):
    expediente = datos.get("expediente", "Sin número")
    demandante = datos.get("demandante", "parte demandante")
    demandado = datos.get("demandado", "parte demandada")
    cuantia = datos.get("cuantia", 0.0)

    estado = resultado["estado"]
    articulo = resultado["articulo"]
    accion = resultado["accion"]

    encabezado = f"""
EXPEDIENTE: {expediente}
DEMANDANTE: {demandante}
DEMANDADO: {demandado}
CUANTÍA: {cuantia:,.2f} lempiras

REFERENCIA LEGAL: {articulo}
ESTADO DETECTADO: {estado}
"""

    if "subsanación" in estado.lower() or "incompleta" in estado.lower():
        cuerpo = f"""
AUTO DE REQUERIMIENTO DE SUBSANACIÓN

Visto el escrito presentado por la parte demandante, y apreciándose que la solicitud inicial no reúne todos los requisitos necesarios para su admisión, procede requerir a la parte actora para que subsane los defectos advertidos en el plazo legal de cinco días.

En atención a lo anterior, se acuerda requerir a {demandante} para que proceda a la subsanación correspondiente, con apercibimiento de archivo definitivo en caso de no hacerlo.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "lista para admisión" in estado.lower():
        cuerpo = f"""
AUTO DE ADMISIÓN DE DEMANDA MONITORIA

Visto el escrito presentado por {demandante} frente a {demandado}, y apreciándose que concurren los presupuestos iniciales de admisibilidad, procede admitir la solicitud monitoria.

En consecuencia, se acuerda requerir al deudor para que, en el plazo legal de veinte días, pague la cantidad reclamada o comparezca formulando oposición.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "pendiente de notificación" in estado.lower():
        cuerpo = f"""
DILIGENCIA DE NOTIFICACIÓN AL DEUDOR

Admitida la demanda monitoria, procede practicar la notificación al deudor {demandado}, requiriéndole para que en el plazo legal pague la cantidad reclamada o formule oposición.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "esperando pago" in estado.lower():
        cuerpo = f"""
INFORME DE CONTROL DE PLAZO

Consta que el deudor {demandado} ha sido notificado del requerimiento de pago. Actualmente se encuentra abierto el plazo legal para que pueda pagar o formular oposición.

DÍAS RESTANTES:
{resultado.get("dias_restantes", "-")}

FECHA DE VENCIMIENTO:
{resultado.get("fecha_vencimiento", "-")}

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "pago realizado" in estado.lower():
        cuerpo = f"""
DILIGENCIA DE PAGO Y ARCHIVO

Constando que el demandado {demandado} ha procedido al pago de la cantidad reclamada, procede entregar el correspondiente comprobante y acordar el archivo de las actuaciones.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "ejecución" in estado.lower() or "apremio" in estado.lower():
        cuerpo = f"""
AUTO DESPACHANDO EJECUCIÓN POR VÍA DE APREMIO

Visto que ha transcurrido el plazo legal conferido al deudor sin que conste pago ni oposición, procede iniciar la ejecución por vía de apremio por la cantidad reclamada.

En consecuencia, se acuerda despachar ejecución frente a {demandado} por importe de {cuantia:,.2f} lempiras, sin perjuicio de la revisión y validación que corresponda por el órgano competente.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "oposición pendiente" in estado.lower():
        cuerpo = f"""
AUTO DE ADMISIÓN DE OPOSICIÓN Y TRASLADO

Presentado escrito de oposición por {demandado}, procede admitirlo y dar traslado a la parte demandante para que pueda actuar conforme corresponda.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "plus petición" in estado.lower():
        cuerpo = f"""
AUTO SOBRE OPOSICIÓN POR PLUS PETICIÓN

Formulada oposición basada en plus petición, procede continuar las actuaciones respecto de la cantidad reconocida como debida, sin perjuicio de resolver lo controvertido conforme al cauce correspondiente.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "abreviado" in estado.lower():
        cuerpo = f"""
PROVIDENCIA DE CONTINUACIÓN POR PROCEDIMIENTO ABREVIADO

Formulada oposición por el demandado, y no excediendo la cuantía de {LIMITE_ABREVIADO:,.2f} lempiras, procede continuar la tramitación por el procedimiento abreviado.

En consecuencia, se acuerda convocar a las partes a la audiencia correspondiente.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "ordinario" in estado.lower() or "pendiente de demanda ordinaria" in estado.lower():
        cuerpo = f"""
PROVIDENCIA DE CONTINUACIÓN POR PROCEDIMIENTO ORDINARIO

Formulada oposición por el demandado y atendida la cuantía reclamada, que excede de {LIMITE_ABREVIADO:,.2f} lempiras, procede continuar por los trámites del procedimiento ordinario.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    if "sobreseimiento" in estado.lower():
        cuerpo = f"""
AUTO DE SOBRESEIMIENTO

Constando que, tras la oposición formulada por el demandado, no se ha presentado la demanda correspondiente dentro del plazo legal, procede acordar el sobreseimiento de las actuaciones.

En consecuencia, se acuerda sobreseer el procedimiento, con los efectos procesales que correspondan.

PRÓXIMA ACTUACIÓN:
{accion}
"""
        return encabezado + cuerpo

    cuerpo = f"""
INFORME DE ESTADO PROCESAL

El sistema ha identificado el siguiente estado procesal:

{resultado["explicacion"]}

PRÓXIMA ACTUACIÓN:
{accion}
"""
    return encabezado + cuerpo


# ============================================================
# DATOS DE EJEMPLO
# ============================================================

def cargar_ejemplos():
    hoy = date.today()

    return [
        {
            "expediente": "MON-001",
            "demandante": "Banco Atlántico",
            "demandado": "Carlos Mejía",
            "cuantia": 35000.0,
            "demanda_presentada": True,
            "tribunal_competente": True,
            "cumple_requisitos": True,
            "subsanacion_requerida": False,
            "demanda_subsanada": False,
            "demanda_admitida": True,
            "deudor_notificado": True,
            "demandado_pago": False,
            "demandado_oposicion": False,
            "plus_peticion": False,
            "traslado_demandante": False,
            "demanda_ordinaria_presentada": False,
            "fecha_requerimiento_subsanacion": hoy,
            "fecha_notificacion_deudor": hoy - timedelta(days=10),
            "fecha_traslado_demandante": hoy,
            "texto_original": ""
        },
        {
            "expediente": "MON-002",
            "demandante": "Comercial Norte S.A.",
            "demandado": "Inversiones López",
            "cuantia": 80000.0,
            "demanda_presentada": True,
            "tribunal_competente": True,
            "cumple_requisitos": True,
            "subsanacion_requerida": False,
            "demanda_subsanada": False,
            "demanda_admitida": True,
            "deudor_notificado": True,
            "demandado_pago": False,
            "demandado_oposicion": False,
            "plus_peticion": False,
            "traslado_demandante": False,
            "demanda_ordinaria_presentada": False,
            "fecha_requerimiento_subsanacion": hoy,
            "fecha_notificacion_deudor": hoy - timedelta(days=25),
            "fecha_traslado_demandante": hoy,
            "texto_original": ""
        },
        {
            "expediente": "MON-003",
            "demandante": "Servicios Técnicos HN",
            "demandado": "María Rivera",
            "cuantia": 120000.0,
            "demanda_presentada": True,
            "tribunal_competente": True,
            "cumple_requisitos": True,
            "subsanacion_requerida": False,
            "demanda_subsanada": False,
            "demanda_admitida": True,
            "deudor_notificado": True,
            "demandado_pago": False,
            "demandado_oposicion": True,
            "plus_peticion": False,
            "traslado_demandante": True,
            "demanda_ordinaria_presentada": False,
            "fecha_requerimiento_subsanacion": hoy,
            "fecha_notificacion_deudor": hoy - timedelta(days=5),
            "fecha_traslado_demandante": hoy - timedelta(days=20),
            "texto_original": ""
        }
    ]


# ============================================================
# ESTADO DE SESIÓN
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
st.subheader("MVP para lectura, clasificación y monitorización del proceso monitorio")

st.markdown(
    """
    Esta aplicación permite subir una demanda o solicitud monitoria, extraer datos básicos,
    validar esos datos y aplicar un motor de reglas para determinar la fase procesal,
    el plazo, el riesgo y la actuación recomendada.
    """
)

st.warning(
    "Prototipo para hackathon. No dicta sentencias ni sustituye la revisión jurídica humana. "
    "Genera propuestas y borradores revisables."
)

st.sidebar.title("Menú")
pagina = st.sidebar.radio(
    "Selecciona una pantalla",
    [
        "1. Subir demanda",
        "2. Dashboard",
        "3. Detalle de expediente",
        "4. Mapa de reglas"
    ]
)


# ============================================================
# PANTALLA 1: SUBIR DEMANDA
# ============================================================

if pagina == "1. Subir demanda":
    st.header("📄 Subir demanda o solicitud monitoria")

    archivo = st.file_uploader(
        "Sube un documento en formato TXT, DOCX o PDF",
        type=["txt", "docx", "pdf"]
    )

    if archivo is not None:
        texto = leer_documento_subido(archivo)
        st.session_state.texto_subido = texto

        st.subheader("Texto detectado")
        st.text_area("Contenido del documento", texto, height=250)

        datos_extraidos = extraer_datos_demanda(texto)
        st.session_state.datos_extraidos = datos_extraidos

        st.subheader("Datos extraídos automáticamente")
        st.json(datos_extraidos)

        if datos_extraidos["datos_faltantes"]:
            st.warning(
                "Datos que conviene revisar o completar: "
                + ", ".join(datos_extraidos["datos_faltantes"])
            )
        else:
            st.success("El sistema ha detectado los datos principales.")

    st.divider()

    st.header("✅ Validación humana de datos")

    datos = st.session_state.datos_extraidos or {
        "demandante": "",
        "demandado": "",
        "cuantia": 0.0,
        "hay_documento_de_deuda": False,
        "menciona_monitorio": False,
        "peticion": "",
        "hechos_resumidos": "",
        "datos_faltantes": []
    }

    with st.form("form_validacion"):
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

        st.subheader("Revisión inicial")

        c1, c2, c3 = st.columns(3)

        demanda_presentada = c1.checkbox(
            "Demanda presentada",
            value=True
        )

        tribunal_competente = c2.checkbox(
            "Tribunal competente",
            value=True
        )

        cumple_requisitos = c3.checkbox(
            "Cumple requisitos del monitorio",
            value=bool(datos.get("hay_documento_de_deuda", False))
        )

        st.caption(
            "Para el MVP, se considera que cumple requisitos si contiene datos básicos y algún documento acreditativo de deuda. "
            "Este punto debe validarlo una persona."
        )

        st.subheader("Subsanación")

        c1, c2, c3 = st.columns(3)

        subsanacion_requerida = c1.checkbox(
            "Subsanación requerida",
            value=not bool(datos.get("hay_documento_de_deuda", False))
        )

        demanda_subsanada = c2.checkbox(
            "Demanda subsanada",
            value=False
        )

        fecha_requerimiento_subsanacion = c3.date_input(
            "Fecha requerimiento subsanación",
            value=date.today()
        )

        st.subheader("Admisión y notificación")

        c1, c2, c3 = st.columns(3)

        demanda_admitida = c1.checkbox(
            "Demanda admitida",
            value=False
        )

        deudor_notificado = c2.checkbox(
            "Deudor notificado",
            value=False
        )

        fecha_notificacion_deudor = c3.date_input(
            "Fecha notificación al deudor",
            value=date.today()
        )

        st.subheader("Actuación del demandado")

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

        st.subheader("Fase posterior a la oposición")

        c1, c2, c3 = st.columns(3)

        traslado_demandante = c1.checkbox(
            "Se dio traslado al demandante",
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

        peticion = st.text_area(
            "Petición detectada o validada",
            value=datos.get("peticion", ""),
            height=100
        )

        hechos_resumidos = st.text_area(
            "Resumen de hechos",
            value=datos.get("hechos_resumidos", ""),
            height=150
        )

        guardar = st.form_submit_button("Guardar y evaluar expediente")

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
            "peticion": peticion,
            "hechos_resumidos": hechos_resumidos,
            "texto_original": st.session_state.texto_subido
        }

        st.session_state.expedientes.append(expediente_datos)

        resultado = evaluar_expediente(expediente_datos)
        borrador = generar_borrador(expediente_datos, resultado)

        st.success("Expediente guardado y evaluado correctamente.")

        st.subheader("Resultado del flujo")
        c1, c2, c3 = st.columns(3)
        c1.metric("Estado", resultado["estado"])
        c2.metric("Riesgo", pintar_riesgo(resultado["riesgo"]))
        c3.metric(
            "Días restantes",
            resultado["dias_restantes"] if resultado["dias_restantes"] is not None else "-"
        )

        st.write(f"**Artículo aplicable:** {resultado['articulo']}")
        st.write(f"**Qué dice el artículo:** {resultado['resumen_articulo']}")
        st.write(f"**Explicación del caso:** {resultado['explicacion']}")
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
# PANTALLA 2: DASHBOARD
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
            "Próxima actuación": res["accion"]
        })

    df = pd.DataFrame(filas)

    st.dataframe(df, use_container_width=True, hide_index=True)

    col1, col2, col3, col4 = st.columns(4)

    total = len(filas)
    vencidos = sum(1 for fila in filas if "Vencido" in fila["Riesgo"])
    alto = sum(1 for fila in filas if "Alto" in fila["Riesgo"])
    oposiciones = sum(
        1 for exp in st.session_state.expedientes
        if exp.get("demandado_oposicion")
    )

    col1.metric("Total expedientes", total)
    col2.metric("Vencidos", vencidos)
    col3.metric("Riesgo alto", alto)
    col4.metric("Con oposición", oposiciones)


# ============================================================
# PANTALLA 3: DETALLE
# ============================================================

elif pagina == "3. Detalle de expediente":
    st.header("🔎 Detalle de expediente")

    if not st.session_state.expedientes:
        st.warning("No hay expedientes cargados.")
    else:
        opciones = [exp["expediente"] for exp in st.session_state.expedientes]
        seleccionado = st.selectbox("Selecciona expediente", opciones)

        exp = next(e for e in st.session_state.expedientes if e["expediente"] == seleccionado)
        res = evaluar_expediente(exp)
        borrador = generar_borrador(exp, res)

        c1, c2, c3 = st.columns(3)
        c1.metric("Estado", res["estado"])
        c2.metric("Riesgo", pintar_riesgo(res["riesgo"]))
        c3.metric(
            "Días restantes",
            res["dias_restantes"] if res["dias_restantes"] is not None else "-"
        )

        st.subheader("Datos principales")

        datos_tabla = {
            "Expediente": exp["expediente"],
            "Demandante": exp["demandante"],
            "Demandado": exp["demandado"],
            "Cuantía": f"{exp['cuantia']:,.2f} lempiras",
            "Fecha notificación deudor": str(exp["fecha_notificacion_deudor"]),
            "Fecha traslado demandante": str(exp["fecha_traslado_demandante"])
        }

        st.json(datos_tabla)

        st.subheader("Evaluación jurídica")
        st.write(f"**Estado:** {res['estado']}")
        st.write(f"**Artículo aplicable:** {res['articulo']}")
        st.write(f"**Resumen del artículo:** {res['resumen_articulo']}")
        st.write(f"**Explicación aplicada al caso:** {res['explicacion']}")
        st.write(f"**Próxima actuación:** {res['accion']}")

        if res["fecha_vencimiento"]:
            st.write(f"**Fecha de vencimiento:** {res['fecha_vencimiento']}")

        if exp.get("hechos_resumidos"):
            st.subheader("Hechos resumidos")
            st.write(exp["hechos_resumidos"])

        st.subheader("Borrador generado")
        st.text_area("Borrador revisable", borrador, height=450)

        st.download_button(
            "Descargar borrador TXT",
            data=borrador,
            file_name=f"borrador_{exp['expediente']}.txt",
            mime="text/plain"
        )


# ============================================================
# PANTALLA 4: MAPA DE REGLAS
# ============================================================

elif pagina == "4. Mapa de reglas":
    st.header("🧠 Mapa de reglas procesales")

    st.markdown(
        """
        ## Flujo lógico del proceso monitorio

        ### 1. Presentación de demanda
        Si no existe demanda, el proceso no se inicia.

        ### 2. Examen de competencia
        Si el tribunal no es competente, se rechaza de plano y se remite al tribunal competente.

        ### 3. Revisión de requisitos
        Si la demanda no cumple requisitos, se requiere subsanación en cinco días.
        Si no se subsana, se archiva definitivamente.

        ### 4. Admisión y requerimiento
        Si la demanda cumple requisitos, se admite y se requiere al deudor para que pague u oponga en veinte días.

        ### 5. Respuesta del deudor
        El deudor puede:
        - pagar;
        - oponerse;
        - no pagar ni comparecer.

        ### 6. Si paga
        Se entrega comprobante y se archiva.

        ### 7. Si no paga ni se opone
        Vencido el plazo, procede ejecución por vía de apremio.

        ### 8. Si se opone
        Se da traslado al demandante y se determina el procedimiento por cuantía.

        ### 9. Cuantía
        - Si no excede de 50.000 lempiras: procedimiento abreviado.
        - Si excede de 50.000 lempiras: procedimiento ordinario.

        ### 10. Falta de demanda posterior
        Si tras la oposición no se presenta la demanda correspondiente en plazo, procede sobreseimiento.
        """
    )

    st.subheader("Reglas simplificadas en pseudocódigo")

    st.code(
        """
if not demanda_presentada:
    estado = "Sin demanda presentada"

elif not tribunal_competente:
    estado = "Rechazo por falta de competencia"

elif not cumple_requisitos:
    estado = "Subsanación pendiente"
    plazo = 5

elif cumple_requisitos and not demanda_admitida:
    estado = "Demanda lista para admisión"

elif demanda_admitida and not deudor_notificado:
    estado = "Pendiente de notificación al deudor"

elif deudor_notificado and demandado_pago:
    estado = "Pago realizado y archivo"

elif deudor_notificado and not pago and not oposicion and plazo_20_dias_vencido:
    estado = "Ejecución por vía de apremio"

elif demandado_oposicion and cuantia <= 50000:
    estado = "Procedimiento abreviado"

elif demandado_oposicion and cuantia > 50000:
    estado = "Procedimiento ordinario"

elif oposicion and cuantia > 50000 and no_demanda_ordinaria_en_plazo:
    estado = "Sobreseimiento"
        """,
        language="python"
    )

    st.subheader("Artículos cargados en el sistema")

    tabla_articulos = []

    for clave, valor in ARTICULOS.items():
        tabla_articulos.append({
            "Regla": clave,
            "Referencia": valor["referencia"],
            "Resumen operativo": valor["resumen"]
        })

    st.dataframe(
        pd.DataFrame(tabla_articulos),
        use_container_width=True,
        hide_index=True
    )
