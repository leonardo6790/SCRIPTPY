import os
import re
import sys
from datetime import datetime

import fitz
import pandas as pd


# =========================
# CONFIGURACION
# =========================
CARPETA_PDFS = r"C:\Users\david.bautista\Downloads\mayo\extraccion"
SALIDA_EXCEL = r"C:\Users\david.bautista\Downloads\mayo\extraccion\zeus_importacion_fuente_55.xlsx"

UNIDAD_NEGOCIO_DEFECTO = "P-BOG-ADM"


ZDOCUMENT_COLUMNS = [
    "Año y Mes",
    "Fuente",
    "Documento",
    "Fecha del Documento",
    "Número de Documento",
    "Total Débitos",
    "Total Créditos",
    "Descripción",
    "Número de Concepto",
    "Tercero",
    "Unidad de Negocio",
]

ZTRANSAC_COLUMNS = [
    "Año y Mes",
    "Fuente",
    "Documento",
    "Fecha",
    "Cuenta",
    "Ind. Tipo de Cuenta",
    "Cco",
    "Tercero",
    "Cliente",
    "Proveedor",
    "Descripción",
    "Numero de Concepto",
    "Valor",
    "Aplicación",
    "Tipo Factura",
    "Factura",
    "Fecha Vencimiento",
    "Referencia",
    "Unidad de Negocio",
    "Base",
    "Porcentaje",
    "Tasa de Cambio",
    "Débito",
    "Crédito",
]


# =========================
# UTILIDADES
# =========================
def limpiar_espacios(valor):
    return re.sub(r"\s+", " ", str(valor or "")).strip()


def buscar(patron, texto, default=""):
    m = re.search(patron, texto, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return m.group(1).strip() if m else default


def solo_nit(valor):
    valor = str(valor or "")
    valor = valor.split("-")[0]
    return re.sub(r"\D", "", valor)


def limpiar_tercero(valor):
    valor = str(valor or "").strip()
    valor = valor.split("-")[0].strip()
    return re.sub(r"[^A-Z0-9]", "", valor.upper())


def numero_a_float(valor):
    if valor is None or valor == "":
        return 0.0

    v = str(valor).strip()
    v = v.replace("$", "").replace("COP", "").replace(" ", "")

    if "," in v and "." in v:
        if v.rfind(",") > v.rfind("."):
            v = v.replace(".", "").replace(",", ".")
        else:
            v = v.replace(",", "")
    elif "," in v:
        v = v.replace(",", ".")

    try:
        return float(v)
    except Exception:
        return 0.0


def formatear_fecha(fecha):
    fecha = str(fecha or "").strip()

    formatos = [
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d/%m/%y",
        "%d-%m-%y",
    ]

    for fmt in formatos:
        try:
            return datetime.strptime(fecha[:10], fmt).strftime("%Y/%m/%d")
        except Exception:
            pass

    return fecha


def ano_mes(fecha):
    fecha = formatear_fecha(fecha)
    return fecha[:7].replace("/", "") if fecha else ""


def separar_fuente_documento(numero_doc):
    numero_doc = str(numero_doc or "").replace(" ", "")

    if "-" in numero_doc:
        fuente, documento = numero_doc.split("-", 1)
        return fuente, documento

    return "", numero_doc


def extraer_detalles_soporte(texto):
    lineas = [limpiar_espacios(x) for x in texto.splitlines() if limpiar_espacios(x)]
    dinero_re = re.compile(r"-?\d{1,3}(?:,\d{3})+\.\d{2}|-?\d+\.\d{2}")
    cantidad_re = re.compile(r"-?\d+(?:\.\d+)?")
    codigo_re = re.compile(r"\d{3,6}")
    detalles = []

    inicio = None

    for idx, linea in enumerate(lineas):
        if linea.upper() == "DESC. IVA":
            inicio = idx + 1
            break

    if inicio is None:
        return detalles

    fin = len(lineas)

    for idx in range(inicio, len(lineas)):
        linea_upper = lineas[idx].upper()

        if linea_upper.startswith(("SON:", "OBSERVACIONES", "TOTAL COMPRA")):
            fin = idx
            break

    detalle_lineas = lineas[inicio:fin]

    def es_inicio_fila(pos):
        return (
            pos + 2 < len(detalle_lineas)
            and dinero_re.fullmatch(detalle_lineas[pos])
            and dinero_re.fullmatch(detalle_lineas[pos + 1])
            and cantidad_re.fullmatch(detalle_lineas[pos + 2])
        )

    def hay_otra_fila_desde(pos):
        return any(es_inicio_fila(i) for i in range(pos, len(detalle_lineas)))

    i = 0

    while i < len(detalle_lineas):
        if not es_inicio_fila(i):
            i += 1
            continue

        valor = numero_a_float(detalle_lineas[i])
        descripcion_partes = []
        numero_concepto = ""
        j = i + 3

        while j < len(detalle_lineas):
            if es_inicio_fila(j):
                break

            actual = detalle_lineas[j]
            siguiente = detalle_lineas[j + 1] if j + 1 < len(detalle_lineas) else ""

            if (
                codigo_re.fullmatch(actual)
                and descripcion_partes
                and siguiente != "/"
                and (es_inicio_fila(j + 1) or not hay_otra_fila_desde(j + 1))
            ):
                numero_concepto = actual
                j += 1
                break

            descripcion_partes.append(actual)
            j += 1

        descripcion = limpiar_espacios(" ".join(descripcion_partes))

        if numero_concepto or descripcion:
            detalles.append({
                "numero_concepto": numero_concepto,
                "descripcion": descripcion,
                "valor": valor,
            })

        i = max(j, i + 1)

    return detalles


def extraer_detalle_soporte(texto):
    detalles = extraer_detalles_soporte(texto)

    if not detalles:
        return "", ""

    return detalles[0].get("numero_concepto", ""), detalles[0].get("descripcion", "")


def normalizar_factura(valor):
    return re.sub(r"[^A-Z0-9]", "", str(valor or "").upper())


def imprimir_seguro(texto):
    try:
        print(texto)
    except UnicodeEncodeError:
        print(str(texto).encode("ascii", "backslashreplace").decode("ascii"))


def ruta_excel_disponible(ruta_excel):
    if not os.path.exists(ruta_excel):
        return ruta_excel

    try:
        with open(ruta_excel, "ab"):
            pass
        return ruta_excel
    except PermissionError:
        base, ext = os.path.splitext(ruta_excel)
        sufijo = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base}_{sufijo}{ext}"


def es_soporte_adjunto(nombre_archivo):
    nombre = nombre_archivo.upper()

    if nombre.startswith("DOCUMENTO_SOPORTE") or "DOC._CONTABLE" in nombre:
        return False

    return bool(re.search(r"SOPO*RTE", nombre))


# =========================
# LECTURA PDF
# =========================
def extraer_texto_pdf(ruta_pdf):
    texto = ""

    with fitz.open(ruta_pdf) as pdf:
        for pagina in pdf:
            texto += pagina.get_text("text") + "\n"

    return texto


# =========================
# DETECCION TIPO
# =========================
def detectar_tipo(texto):
    t = texto.upper()

    if "FACTURA ELECTRÓNICA" in t or "FACTURA ELECTRONICA" in t:
        return "Factura Electronica"

    if "DOCUMENTO SOPORTE ELECTRONICO" in t or "OCUMENTO SOPORTE ELECTRONICO" in t:
        return "Documento Zeus"

    if "DOC. CONTABLE" in t and "DOCUMENTO SOPORTE" in t:
        return "Documento Soporte"

    tipos = [
        "CAUSACIONES ADMINISTRATIVAS",
        "CAUSACIONES TURISTICOS",
        "CAUSACIONES TURÍSTICOS",
        "CAUSACIONES IMPUESTOS",
        "CAUSACIONES INCAPACIDADES",
        "COMPROBANTE RECLASIFICACIÓN PROVEEDORES",
        "COMPROBANTE RECLASIFICACION PROVEEDORES",
        "COMPROBANTES DE EGRESO",
        "LEG ANTICIPO PROVEEDORES TESORERIA",
        "LEGALIZACIÓN CAJA MENOR",
        "LEGALIZACION CAJA MENOR",
        "NOTA CREDITO FACTURA PROVEEDOR",
        "NOTAS CONTABLES",
        "AJUSTES CONTABLE",
        "AMORTIZACIONES",
        "RECIBO DE CAJA",
    ]

    for tipo in tipos:
        if tipo in t:
            return tipo.title()

    if "RECIBO DE CAJA" in t and "TOTALES" in t:
        return "Recibo De Caja"

    if "CUENTA AUX/CCO/TER/CLI/PROV" in t and "TOTALES" in t:
        return "Documento Zeus"

    return "Desconocido"


# =========================
# FACTURA ELECTRONICA
# =========================
def extraer_factura_electronica(texto, archivo):
    factura = (
        buscar(r"Número de Factura:\s*([A-Z0-9\-]+)", texto)
        or buscar(r"Numero de Factura:\s*([A-Z0-9\-]+)", texto)
        or buscar(r"No\.\s*([A-Z0-9\-]+)", texto)
    )

    nit = (
        solo_nit(buscar(r"Nit del Emisor:\s*([0-9\-\.]+)", texto))
        or solo_nit(buscar(r"NIT\s*([0-9\-\.]+)", texto))
    )

    proveedor = (
        buscar(r"Razón Social:\s*(.+?)\s*Nombre Comercial", texto)
        or buscar(r"Razon Social:\s*(.+?)\s*Nombre Comercial", texto)
    )

    fecha = (
        buscar(r"Fecha de Emisión:\s*([0-9/]{10})", texto)
        or buscar(r"Fecha de Emision:\s*([0-9/]{10})", texto)
    )

    vencimiento = buscar(r"Fecha de Vencimiento:\s*([0-9/]{10})", texto)

    valor = (
        buscar(r"Total factura\s*\(=\)\s*COP\s*\$?\s*([\d\.\,]+)", texto)
        or buscar(r"Total Factura Electrónica\s*\$?\s*([\d\.\,]+)", texto)
    )

    cufe = (
        buscar(r"CUFE\s*:\s*([a-f0-9]{50,})", texto)
        or buscar(r"Código Único de Factura\s*-\s*CUFE\s*:\s*([a-f0-9]{50,})", texto)
    )

    return {
        "archivo": archivo,
        "tipo": "Factura Electronica",
        "factura": factura,
        "factura_key": normalizar_factura(factura),
        "nit": nit,
        "proveedor": limpiar_espacios(proveedor),
        "fecha": formatear_fecha(fecha),
        "vencimiento": formatear_fecha(vencimiento),
        "valor_factura": numero_a_float(valor),
        "cufe": cufe,
    }


# =========================
# ENCABEZADO ZEUS
# =========================
def extraer_encabezado_zeus(texto, archivo):
    tipo = detectar_tipo(texto)

    numero_doc = (
        buscar(r"No:\s*([0-9]+\s*-\s*0*\d+)", texto)
        or buscar(r"RECIBO DE CAJA\s+No\.\s*([0-9]+\s*-\s*0*\d+)", texto)
        or buscar(r"DOC\. CONTABLE:\s*([0-9]+\s*-\s*0*\d+)", texto)
        or buscar(r"ORDEN DE COMPRA:\s*([0-9]+\s*-\s*0*\d+)", texto)
        or buscar(r"FACT\.DE PROVEEDOR:\s*([A-Z0-9\-]+)", texto)
    )

    numero_doc = numero_doc.replace(" ", "")
    fuente, documento = separar_fuente_documento(numero_doc)

    fecha = (
        buscar(r"\n\s*Fecha:\s*([0-9]{4}/[0-9]{2}/[0-9]{2})", texto)
        or buscar(r"\n\s*FECHA:\s*([0-9]{4}/[0-9]{2}/[0-9]{2})", texto)
        or buscar(r"Fecha\s*:\s*([0-9]{4}/[0-9]{2}/[0-9]{2})", texto)
        or buscar(r"FECHA:\s*([0-9]{4}/[0-9]{2}/[0-9]{2})", texto)
    )

    unidad = buscar(r"Unidad de Negocio:\s*([A-Z0-9\-]+)", texto) or UNIDAD_NEGOCIO_DEFECTO

    nit = (
        limpiar_tercero(buscar(r"Tercero\s*:\s*\n?\s*([A-Z0-9]+)", texto))
        or limpiar_tercero(buscar(r"Cliente\s*:\s*([A-Z0-9]+)", texto))
        or solo_nit(buscar(r"NIT\s*:\s*([0-9]+)", texto))
        or solo_nit(buscar(r"Proveedor\s*:\s*\n?\s*([0-9]+)", texto))
    )

    proveedor = (
        buscar(r"Nombre\s*:\s*\n?\s*(.+?)\s*Tasa de Cambio", texto)
        or buscar(r"Nombre\s*:\s*\n?\s*(.+?)\s*Ciudad", texto)
        or buscar(r"Proveedor\s*:\s*\n?\s*[0-9]+\s+(.+?)\s+NIT", texto)
    )

    concepto = (
        buscar(r"Por Concepto:\s*(.+?)\s*Cuenta\s+Aux", texto)
        or buscar(r"Concepto\s*:\s*(.+?)\s*Cliente", texto)
        or buscar(r"Observaciones:\s*(.+?)\s*Total Compra", texto)
    )

    return {
        "archivo": archivo,
        "tipo": tipo,
        "numero_doc": numero_doc,
        "fuente": fuente,
        "documento": documento,
        "fecha": formatear_fecha(fecha),
        "unidad_negocio": unidad,
        "nit": nit,
        "proveedor": limpiar_espacios(proveedor),
        "concepto": limpiar_espacios(concepto),
    }


# =========================
# LINEAS CONTABLES ZEUS
# =========================
def extraer_lineas_zeus(texto):
    lineas = [limpiar_espacios(x) for x in texto.splitlines() if limpiar_espacios(x)]
    movimientos = []

    dinero_re = re.compile(r"-?\d{1,3}(?:,\d{3})+\.\d{2}|-?\d+\.\d{2}")
    cuenta_re = re.compile(r"^\d{4,12}\b")
    cuenta_linea_re = re.compile(r"^\d{4,12}$")

    def es_dinero(texto_valor):
        return bool(dinero_re.fullmatch(texto_valor) or texto_valor in {"0", "-0"})

    def extraer_datos_descripcion(descripcion):
        tipo_factura = ""
        factura = ""
        fecha_venc = ""
        referencia = ""

        mtipo = re.search(r"\b(FP|FA|CA|NC|AT|CE)\b", descripcion, re.IGNORECASE)
        if mtipo:
            tipo_factura = mtipo.group(1).upper()

        mfactura = re.search(
            r"\b(?:FP|FA|CA|NC|AT|CE)\s+([A-Z0-9\-]+)",
            descripcion,
            re.IGNORECASE
        )
        if mfactura:
            factura = mfactura.group(1)

        mfecha = re.search(
            r"(20\d{2}/\d{2}/\d{2}|20\d{2}-\d{2}-\d{2}|\d{2}/\d{2}/20\d{2}|\d{2}-\d{2}-20\d{2}|\d{2}-\d{2}-\d{2})",
            descripcion
        )
        if mfecha:
            fecha_venc = formatear_fecha(mfecha.group(1))

        mref = re.search(
            r"\b(FEPN-\d+|LOC\.?[A-Z0-9]+|CRM\d+|PNV\d+|[A-Z]{2,5}-\d+|[A-Z0-9]{5,15})\b",
            descripcion,
            re.IGNORECASE
        )
        if mref:
            referencia = mref.group(1)

        return tipo_factura, factura, fecha_venc, referencia

    def agregar_movimiento(cuenta, auxiliar, descripcion, debito, credito):
        tipo_factura, factura, fecha_venc, referencia = extraer_datos_descripcion(descripcion)

        movimientos.append({
            "Cuenta": cuenta,
            "Auxiliar": auxiliar,
            "Descripcion": descripcion,
            "Debito": debito,
            "Credito": credito,
            "TipoFactura": tipo_factura,
            "Factura": factura,
            "FechaVencimiento": fecha_venc,
            "Referencia": referencia,
        })

    # Formato de RECIBO DE CAJA:
    # Cuenta / Credito / Descripcion / Debito. PyMuPDF puede omitir columnas
    # vacias, por eso se procesa como bloques de cuatro lineas.
    if any("RECIBO DE CAJA" in linea.upper() for linea in lineas):
        i = 0
        while i <= len(lineas) - 4:
            cuenta = lineas[i]
            credito_txt = lineas[i + 1]
            descripcion = lineas[i + 2]
            debito_txt = lineas[i + 3]

            if (
                cuenta_linea_re.match(cuenta)
                and es_dinero(credito_txt)
                and not descripcion.upper().startswith(("CUENTA", "TOTALES", "FORMA PAGO", "RESUMEN"))
                and es_dinero(debito_txt)
            ):
                agregar_movimiento(
                    cuenta=cuenta,
                    auxiliar="",
                    descripcion=descripcion,
                    debito=numero_a_float(debito_txt),
                    credito=numero_a_float(credito_txt),
                )
                i += 4
                continue

            i += 1

        if movimientos:
            return movimientos

    i = 0
    while i <= len(lineas) - 5:
        descripcion = lineas[i]
        cuenta = lineas[i + 1]
        auxiliar = lineas[i + 2]
        debito_txt = lineas[i + 3]
        credito_txt = lineas[i + 4]

        if (
            not descripcion.upper().startswith(("CUENTA", "TOTALES", "ELABOR"))
            and cuenta_linea_re.match(cuenta)
            and dinero_re.fullmatch(debito_txt)
            and dinero_re.fullmatch(credito_txt)
        ):
            agregar_movimiento(
                cuenta=cuenta,
                auxiliar=auxiliar,
                descripcion=descripcion,
                debito=numero_a_float(debito_txt),
                credito=numero_a_float(credito_txt),
            )
            i += 5
            continue

        i += 1

    if movimientos:
        return movimientos

    for linea in lineas:
        linea_upper = linea.upper()

        if linea_upper.startswith("TOTALES"):
            continue

        if linea_upper.startswith("CUENTA"):
            continue

        if not cuenta_re.match(linea):
            continue

        valores = list(dinero_re.finditer(linea))

        if len(valores) < 2:
            continue

        debito_txt = valores[-2].group()
        credito_txt = valores[-1].group()

        debito = numero_a_float(debito_txt)
        credito = numero_a_float(credito_txt)

        texto_antes_valores = linea[:valores[-2].start()].strip()
        partes = texto_antes_valores.split()

        if len(partes) < 2:
            continue

        cuenta = partes[0]
        auxiliar = partes[1]
        descripcion = " ".join(partes[2:]).strip()

        agregar_movimiento(cuenta, auxiliar, descripcion, debito, credito)

    return movimientos


# =========================
# DOCUMENTO SOPORTE
# =========================
def extraer_documento_soporte(texto, archivo):
    encabezado = extraer_encabezado_zeus(texto, archivo)
    detalles = extraer_detalles_soporte(texto)
    numero_concepto, descripcion_concepto = extraer_detalle_soporte(texto)

    valor = (
        buscar(r"Total Factura\s*(-?[\d,]+\.\d{2})", texto)
        or buscar(r"Total Compra Neta\s*(-?[\d,]+\.\d{2})", texto)
        or buscar(r"Valor a Pagar:\s*([\d,]+\.\d{2})", texto)
    )

    factura = buscar(r"FACT\.DE PROVEEDOR:\s*([A-Z0-9\-]+)", texto)

    total = numero_a_float(valor)

    descripcion = (
        descripcion_concepto
        or buscar(r"Observaciones:\s*(.+?)\s*Total Compra", texto)
        or encabezado.get("concepto")
        or encabezado.get("proveedor")
    )

    movimientos = []

    if detalles:
        for detalle in detalles:
            valor_detalle = detalle.get("valor", 0.0)
            debito = valor_detalle if valor_detalle > 0 else 0.0
            credito = abs(valor_detalle) if valor_detalle < 0 else 0.0

            movimientos.append({
                "Cuenta": "",
                "Auxiliar": encabezado.get("nit", ""),
                "Descripcion": limpiar_espacios(detalle.get("descripcion", "")),
                "NumeroConcepto": detalle.get("numero_concepto", ""),
                "Debito": debito,
                "Credito": credito,
                "TipoFactura": "",
                "Factura": factura,
                "FechaVencimiento": encabezado.get("fecha", ""),
                "Referencia": factura,
            })
    elif total != 0:
        debito = total if total > 0 else 0.0
        credito = abs(total) if total < 0 else 0.0

        movimientos.append({
            "Cuenta": "",
            "Auxiliar": encabezado.get("nit", ""),
            "Descripcion": limpiar_espacios(descripcion),
            "NumeroConcepto": numero_concepto,
            "Debito": debito,
            "Credito": credito,
            "TipoFactura": "",
            "Factura": factura,
            "FechaVencimiento": encabezado.get("fecha", ""),
            "Referencia": factura,
        })

    encabezado.update({
        "factura": factura,
        "factura_key": normalizar_factura(factura),
        "referencia": factura,
        "numero_concepto": numero_concepto,
        "descripcion_concepto": limpiar_espacios(descripcion_concepto),
        "debito_total": sum(m["Debito"] for m in movimientos),
        "credito_total": sum(m["Credito"] for m in movimientos),
        "transacciones": movimientos,
    })

    return encabezado


# =========================
# DOCUMENTO CONTABLE
# =========================
def extraer_documento_contable(texto, archivo):
    tipo = detectar_tipo(texto)

    if tipo == "Documento Soporte":
        return extraer_documento_soporte(texto, archivo)

    encabezado = extraer_encabezado_zeus(texto, archivo)
    movimientos = extraer_lineas_zeus(texto)

    debito_total = sum(m["Debito"] for m in movimientos)
    credito_total = sum(m["Credito"] for m in movimientos)

    if debito_total == 0:
        debito_total = numero_a_float(buscar(r"Totales\s+([\d,]+\.\d{2})", texto))

    if credito_total == 0:
        credito_total = numero_a_float(buscar(r"Totales\s+[\d,]+\.\d{2}\s+([\d,]+\.\d{2})", texto))

    factura = ""
    referencia = ""

    for mov in movimientos:
        if not factura and mov.get("Factura"):
            factura = mov.get("Factura")

        if not referencia and mov.get("Referencia"):
            referencia = mov.get("Referencia")

    encabezado.update({
        "factura": factura,
        "factura_key": normalizar_factura(factura),
        "referencia": referencia,
        "debito_total": debito_total,
        "credito_total": credito_total,
        "transacciones": movimientos,
    })

    return encabezado


# =========================
# FILAS EXCEL
# =========================
def descripcion_documento(doc):
    if doc.get("fuente") in {"55", "56"}:
        return (
            doc.get("descripcion_concepto", "")
            or doc.get("concepto", "")
            or doc.get("proveedor", "")
        )

    return doc.get("proveedor", "") or doc.get("concepto", "")


def fila_zdocument(doc):
    return {
        "Año y Mes": ano_mes(doc.get("fecha", "")),
        "Fuente": doc.get("fuente", ""),
        "Documento": doc.get("documento", ""),
        "Fecha del Documento": doc.get("fecha", ""),
        "Número de Documento": doc.get("numero_doc", ""),
        "Total Débitos": doc.get("debito_total", 0.0),
        "Total Créditos": doc.get("credito_total", 0.0),
        "Descripción": descripcion_documento(doc),
        "Número de Concepto": doc.get("numero_concepto", ""),
        "Tercero": doc.get("nit", ""),
        "Unidad de Negocio": doc.get("unidad_negocio", ""),
    }


def fila_ztransac(doc, mov):
    debito = mov.get("Debito", 0.0)
    credito = mov.get("Credito", 0.0)

    valor = debito if debito > 0 else -credito

    return {
        "Año y Mes": ano_mes(doc.get("fecha", "")),
        "Fuente": doc.get("fuente", ""),
        "Documento": doc.get("documento", ""),
        "Fecha": doc.get("fecha", ""),
        "Cuenta": mov.get("Cuenta", ""),
        "Ind. Tipo de Cuenta": "",
        "Cco": mov.get("Cco", "") or mov.get("Auxiliar", ""),
        "Tercero": doc.get("nit", ""),
        "Cliente": "",
        "Proveedor": doc.get("nit", ""),
        "Descripción": mov.get("Descripcion", ""),
        "Numero de Concepto": mov.get("NumeroConcepto", "") or doc.get("numero_concepto", ""),
        "Valor": valor,
        "Aplicación": "",
        "Tipo Factura": mov.get("TipoFactura", ""),
        "Factura": mov.get("Factura", "") or doc.get("factura", ""),
        "Fecha Vencimiento": mov.get("FechaVencimiento", ""),
        "Referencia": mov.get("Referencia", "") or doc.get("referencia", ""),
        "Unidad de Negocio": doc.get("unidad_negocio", ""),
        "Base": "",
        "Porcentaje": "",
        "Tasa de Cambio": "0.0000",
        "Débito": debito,
        "Crédito": credito,
    }


# =========================
# MAIN
# =========================
def main():
    facturas = {}
    documentos = []
    auditoria = []

    archivos_pdf = []

    for raiz, carpetas, archivos in os.walk(CARPETA_PDFS):
        carpetas[:] = [
            carpeta for carpeta in carpetas
            if carpeta.upper() not in {"EGRESOS EXTRAIDO", "EGRESOS EXTRAIDOS"}
            and not carpeta.upper().startswith("EXTRACTO")
        ]

        for archivo in archivos:
            if es_soporte_adjunto(archivo):
                continue

            if archivo.lower().endswith(".pdf"):
                ruta = os.path.join(raiz, archivo)
                archivo_relativo = os.path.relpath(ruta, CARPETA_PDFS)
                archivos_pdf.append((archivo_relativo, ruta))

    archivos_pdf.sort(key=lambda item: item[0].lower())

    for idx, (archivo, ruta) in enumerate(archivos_pdf, 1):
        if idx == 1 or idx == len(archivos_pdf) or idx % 100 == 0:
            imprimir_seguro(f"[{idx}/{len(archivos_pdf)}] Procesando: {archivo}")

        try:
            texto = extraer_texto_pdf(ruta)
            tipo = detectar_tipo(texto)

            if tipo == "Factura Electronica":
                factura = extraer_factura_electronica(texto, archivo)

                if factura.get("factura_key"):
                    facturas[factura["factura_key"]] = factura

                auditoria.append({
                    "archivo": archivo,
                    "tipo": tipo,
                    "factura": factura.get("factura", ""),
                    "nit": factura.get("nit", ""),
                    "proveedor": factura.get("proveedor", ""),
                    "fecha": factura.get("fecha", ""),
                    "valor": factura.get("valor_factura", 0),
                    "estado": "Factura leida",
                })

            elif tipo != "Desconocido":
                doc = extraer_documento_contable(texto, archivo)
                documentos.append(doc)

                auditoria.append({
                    "archivo": archivo,
                    "tipo": doc.get("tipo", ""),
                    "numero_doc": doc.get("numero_doc", ""),
                    "fuente": doc.get("fuente", ""),
                    "documento": doc.get("documento", ""),
                    "fecha": doc.get("fecha", ""),
                    "nit": doc.get("nit", ""),
                    "proveedor": doc.get("proveedor", ""),
                    "numero_concepto": doc.get("numero_concepto", ""),
                    "descripcion_concepto": doc.get("descripcion_concepto", ""),
                    "unidad_negocio": doc.get("unidad_negocio", ""),
                    "debito_total": doc.get("debito_total", 0),
                    "credito_total": doc.get("credito_total", 0),
                    "lineas_contables": len(doc.get("transacciones", [])),
                    "estado": "Documento leido",
                })

            else:
                auditoria.append({
                    "archivo": archivo,
                    "tipo": "Desconocido",
                    "estado": "No se reconocio el PDF",
                })

        except Exception as e:
            auditoria.append({
                "archivo": archivo,
                "tipo": "Error",
                "estado": str(e),
            })

    zdocument_rows = []
    ztransac_rows = []
    consolidado_rows = []

    for doc in documentos:
        factura_info = facturas.get(doc.get("factura_key", ""), {})

        if factura_info:
            if not doc.get("proveedor"):
                doc["proveedor"] = factura_info.get("proveedor", "")

            if not doc.get("nit"):
                doc["nit"] = factura_info.get("nit", "")

        zdocument_rows.append(fila_zdocument(doc))

        for mov in doc.get("transacciones", []):
            ztransac_rows.append(fila_ztransac(doc, mov))

        consolidado_rows.append({
            "Archivo PDF": doc.get("archivo", ""),
            "Tipo": doc.get("tipo", ""),
            "Número Documento": doc.get("numero_doc", ""),
            "Fuente": doc.get("fuente", ""),
            "Documento": doc.get("documento", ""),
            "Fecha": doc.get("fecha", ""),
            "Unidad de Negocio": doc.get("unidad_negocio", ""),
            "Proveedor": doc.get("proveedor", ""),
            "NIT": doc.get("nit", ""),
            "Numero Concepto": doc.get("numero_concepto", ""),
            "Descripcion Concepto": doc.get("descripcion_concepto", ""),
            "Factura": doc.get("factura", ""),
            "Referencia": doc.get("referencia", ""),
            "Débito Total": doc.get("debito_total", 0),
            "Crédito Total": doc.get("credito_total", 0),
            "Líneas Contables": len(doc.get("transacciones", [])),
        })

    ruta_salida_excel = ruta_excel_disponible(SALIDA_EXCEL)

    with pd.ExcelWriter(ruta_salida_excel, engine="openpyxl") as writer:
        pd.DataFrame(zdocument_rows, columns=ZDOCUMENT_COLUMNS).to_excel(
            writer,
            sheet_name="ZDocument_Temp1",
            index=False
        )

        pd.DataFrame(ztransac_rows, columns=ZTRANSAC_COLUMNS).to_excel(
            writer,
            sheet_name="ZTransac_Temp1",
            index=False
        )

        pd.DataFrame(consolidado_rows).to_excel(
            writer,
            sheet_name="Consolidado",
            index=False
        )

        pd.DataFrame(auditoria).to_excel(
            writer,
            sheet_name="Auditoria",
            index=False
        )

    print("")
    print("Excel generado correctamente:")
    print(ruta_salida_excel)
    print("")
    print(f"PDF procesados: {len(archivos_pdf)}")
    print(f"Documentos contables: {len(documentos)}")
    print(f"Movimientos exportados: {len(ztransac_rows)}")
    print(f"Facturas electronicas: {len(facturas)}")


def nombre_mes(carpeta_mes):
    nombre = os.path.basename(carpeta_mes).strip()
    nombre = re.sub(r"\s*FUENTE\s+\d+\s*$", "", nombre, flags=re.IGNORECASE).strip()

    return nombre[:1].upper() + nombre[1:].lower()


def generar_plantillas_por_mes():
    global CARPETA_PDFS, SALIDA_EXCEL

    raiz_base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    fuentes = {
        "55": os.path.join(raiz_base, "FUENTE 55"),
        "56": os.path.join(raiz_base, "FUENTE 56"),
    }

    for fuente, carpeta_fuente in fuentes.items():
        if not os.path.isdir(carpeta_fuente):
            continue

        for entrada in sorted(os.listdir(carpeta_fuente)):
            carpeta_mes = os.path.join(carpeta_fuente, entrada)

            if not os.path.isdir(carpeta_mes):
                continue

            if entrada.upper().startswith(("EGRESOS", "EXTRACTO")):
                continue

            destino = os.path.join(carpeta_mes, "EGRESOS EXTRAIDO")
            os.makedirs(destino, exist_ok=True)

            mes = nombre_mes(carpeta_mes)
            CARPETA_PDFS = carpeta_mes
            SALIDA_EXCEL = os.path.join(destino, f"Plantilla{mes}Fuente{fuente}.xlsx")

            print("")
            print(f"Generando plantilla {mes} fuente {fuente}")
            main()


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--por-mes":
        generar_plantillas_por_mes()
        sys.exit(0)

    if len(sys.argv) >= 2:
        CARPETA_PDFS = sys.argv[1]

    if len(sys.argv) >= 3:
        SALIDA_EXCEL = sys.argv[2]

    main()

