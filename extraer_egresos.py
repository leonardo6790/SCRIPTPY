from pathlib import Path
import shutil
import csv
import re
from datetime import datetime
from collections import Counter, defaultdict
import fitz  # PyMuPDF

RUTA_ORIGEN = r"C:\Users\david.bautista\Downloads\mayo\FUENTE 56\ABRIL"

CARPETA_DESTINO = "EGRESOS EXTRAIDO"

EXTENSIONES_PERMITIDAS = [".pdf"]

# Si esta en True, solo copia PDF que internamente sean documentos validos:
# - Comprobantes de egreso
# - Anticipos proveedores / empleados
# - Documento soporte / doc. contable
COPIAR_SOLO_COINCIDENCIAS = True

# Si esta en True, al copiar el PDF le antepone el tipo y el numero detectado.
RENOMBRAR_CON_NUMERO_DOCUMENTO = True

# Tipos de documento validos por texto interno.
PATRONES_TEXTO_DOCUMENTO_VALIDO = [
    r"COMPROBANTES\s+DE\s+EGRESO",
    r"COMPROBANTE\s+DE\s+EGRESO",
    r"COMPROBANTES DE EGRESO",
    r"COMPROBANTE DE EGRESO",
    r"ANTICIPOS\s+PROVEEDORES\s*/\s*EMPLEADOS",
    r"ANTICIPOS PROVEEDORES / EMPLEADOS",
    r"DOCUMENTO\s+SOPORTE\s+EN\s+ADQUISICIONES",
    r"DOCUMENTO SOPORTE EN ADQUISICIONES",
    r"DOC\.\s*CONTABLE\s*:\s*\d{2}\s*-\s*\d+",
    r"FACT\.DE\s+PROVEEDOR",
    r"FACT.DE PROVEEDOR",
    r"No:\s*\d{2}\s*-\s*\d+",
    r"Débito",
    r"Debito",
    r"Crédito",
    r"Credito",
    r"Valor\s+a\s+Pagar",
    r"Total\s+Factura",
]

# Patrones por nombre, solo como referencia en el informe.
PATRONES_NOMBRE_DOCUMENTO = [
    r"egreso",
    r"egresos",
    r"comprobante",
    r"anticipos",
    r"anticipo",
    r"proveedores",
    r"empleados",
    r"documento",
    r"soporte",
    r"doc",
    r"contable",
    r"fact",
    r"factura",
    r"impricto",
    r"impridcto",
    r"impreso",
    r"imprimir",
    r"\bEGR\b",
    r"\bCE\b",
]


def leer_texto_pdf(ruta_pdf: Path) -> str:
    texto = ""

    try:
        with fitz.open(ruta_pdf) as documento:
            for pagina in documento:
                texto += pagina.get_text("text") + "\n"
    except Exception as e:
        texto = f"ERROR_LEYENDO_PDF: {e}"

    return texto


def detectar_tipo_documento(texto_pdf: str) -> str:
    if re.search(r"COMPROBANTE(S)?\s+DE\s+EGRESO", texto_pdf, re.IGNORECASE):
        return "COMPROBANTE DE EGRESO"

    if re.search(r"ANTICIPOS\s+PROVEEDORES\s*/\s*EMPLEADOS", texto_pdf, re.IGNORECASE):
        return "ANTICIPOS PROVEEDORES / EMPLEADOS"

    if (
        re.search(r"DOCUMENTO\s+SOPORTE\s+EN\s+ADQUISICIONES", texto_pdf, re.IGNORECASE)
        or re.search(r"DOC\.\s*CONTABLE\s*:\s*\d{2}\s*-\s*\d+", texto_pdf, re.IGNORECASE)
        or re.search(r"FACT\.DE\s+PROVEEDOR", texto_pdf, re.IGNORECASE)
    ):
        return "DOCUMENTO SOPORTE / DOC. CONTABLE"

    return ""


def validar_texto_documento(texto_pdf: str) -> tuple[bool, str, str]:
    encontrados = []

    for patron in PATRONES_TEXTO_DOCUMENTO_VALIDO:
        if re.search(patron, texto_pdf, re.IGNORECASE):
            encontrados.append(patron)

    tipo_documento = detectar_tipo_documento(texto_pdf)

    tiene_titulo_valido = bool(tipo_documento)

    tiene_numero_comun = re.search(
        r"No:\s*\d{2}\s*-\s*\d+",
        texto_pdf,
        re.IGNORECASE
    )

    tiene_doc_contable = re.search(
        r"DOC\.\s*CONTABLE\s*:\s*\d{2}\s*-\s*\d+",
        texto_pdf,
        re.IGNORECASE
    )

    tiene_debito_credito = (
        re.search(r"Débito|Debito", texto_pdf, re.IGNORECASE)
        and re.search(r"Crédito|Credito", texto_pdf, re.IGNORECASE)
    )

    tiene_totales_factura = (
        re.search(r"Total\s+Factura", texto_pdf, re.IGNORECASE)
        and re.search(r"Valor\s+a\s+Pagar", texto_pdf, re.IGNORECASE)
    )

    coincide = bool(
        tiene_titulo_valido
        or (tiene_numero_comun and tiene_debito_credito)
        or (tiene_doc_contable and tiene_totales_factura)
    )

    return coincide, ", ".join(encontrados), tipo_documento


def validar_nombre_documento(nombre_archivo: str) -> tuple[bool, str]:
    encontrados = []

    for patron in PATRONES_NOMBRE_DOCUMENTO:
        if re.search(patron, nombre_archivo, re.IGNORECASE):
            encontrados.append(patron)

    return bool(encontrados), ", ".join(encontrados)


def extraer_dato(patron: str, texto: str) -> str:
    resultado = re.search(patron, texto, re.IGNORECASE)

    if resultado:
        if resultado.groups():
            return resultado.group(1).strip()
        return resultado.group(0).strip()

    return ""


def extraer_numero_documento(texto: str) -> tuple[str, str, str, int | None]:
    patrones = [
        r"DOC\.\s*CONTABLE\s*:\s*(\d{2})\s*-\s*(\d+)",
        r"No:\s*(\d{2})\s*-\s*(\d+)",
        r"COMPROBANTES?\s+DE\s+EGRESO\s+No:\s*(\d{2})\s*-\s*(\d+)",
        r"COMPROBANTES?\s+DE\s+EGRESO\s+No:\s*(\d{2})-(\d+)",
        r"ANTICIPOS\s+PROVEEDORES\s*/\s*EMPLEADOS\s+No:\s*(\d{2})\s*-\s*(\d+)",
        r"ANTICIPOS\s+PROVEEDORES\s*/\s*EMPLEADOS\s+No:\s*(\d{2})-(\d+)",
    ]

    for patron in patrones:
        resultado = re.search(patron, texto, re.IGNORECASE)

        if resultado:
            fuente = resultado.group(1).strip()
            consecutivo_texto = resultado.group(2).strip()
            numero_completo = f"{fuente}-{consecutivo_texto}"

            try:
                consecutivo_numero = int(consecutivo_texto)
            except ValueError:
                consecutivo_numero = None

            return numero_completo, fuente, consecutivo_texto, consecutivo_numero

    return "", "", "", None


def extraer_total(texto: str) -> str:
    resultado_totales = re.search(
        r"Totales\s+([\d,\.]+)\s+([\d,\.]+)",
        texto,
        re.IGNORECASE
    )

    if resultado_totales:
        debito = resultado_totales.group(1).strip()
        credito = resultado_totales.group(2).strip()

        if debito == credito:
            return debito

        return f"Debito: {debito} / Credito: {credito}"

    resultado_valor_pagar = re.search(
        r"Valor\s+a\s+Pagar\s*:?\s*([\d,\.]+)",
        texto,
        re.IGNORECASE
    )

    if resultado_valor_pagar:
        return resultado_valor_pagar.group(1).strip()

    resultado_total_factura = re.search(
        r"Total\s+Factura\s+([\d,\.]+)",
        texto,
        re.IGNORECASE
    )

    if resultado_total_factura:
        return resultado_total_factura.group(1).strip()

    return ""


def limpiar_nombre_archivo(texto: str) -> str:
    texto = texto.strip()
    texto = re.sub(r'[\\/:*?"<>|]', "-", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto


def nombre_unico(destino: Path) -> Path:
    if not destino.exists():
        return destino

    base = destino.stem
    ext = destino.suffix
    carpeta = destino.parent
    contador = 1

    while True:
        nuevo = carpeta / f"{base}_{contador}{ext}"
        if not nuevo.exists():
            return nuevo
        contador += 1


def escribir_csv(ruta_csv: Path, campos: list[str], filas: list[dict]):
    with ruta_csv.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=campos, delimiter=";")
        writer.writeheader()
        writer.writerows(filas)


def calcular_consecutivos_faltantes(filas_documentos: list[dict]) -> list[dict]:
    por_tipo_y_fuente = defaultdict(list)

    for fila in filas_documentos:
        tipo_documento = fila.get("tipo_documento_detectado", "")
        fuente = fila.get("fuente", "")
        consecutivo_numero = fila.get("consecutivo_numero", "")

        if not fuente:
            continue

        try:
            consecutivo_numero = int(consecutivo_numero)
        except:
            continue

        clave = (tipo_documento, fuente)
        por_tipo_y_fuente[clave].append(consecutivo_numero)

    faltantes = []

    for (tipo_documento, fuente), numeros in por_tipo_y_fuente.items():
        numeros_unicos = sorted(set(numeros))

        if not numeros_unicos:
            continue

        minimo = min(numeros_unicos)
        maximo = max(numeros_unicos)

        for numero in range(minimo, maximo + 1):
            if numero not in numeros_unicos:
                consecutivo_texto = str(numero).zfill(10)
                faltantes.append({
                    "tipo_documento": tipo_documento,
                    "fuente": fuente,
                    "consecutivo_numero_faltante": numero,
                    "consecutivo_texto_faltante": consecutivo_texto,
                    "numero_documento_faltante": f"{fuente}-{consecutivo_texto}",
                    "rango_validado_desde": f"{fuente}-{str(minimo).zfill(10)}",
                    "rango_validado_hasta": f"{fuente}-{str(maximo).zfill(10)}",
                })

    return faltantes


def detectar_duplicados(filas_documentos: list[dict]) -> list[dict]:
    contador = Counter()

    for fila in filas_documentos:
        clave = (
            fila.get("tipo_documento_detectado", ""),
            fila.get("numero_documento_detectado", "")
        )

        if clave[1]:
            contador[clave] += 1

    duplicados = []

    for (tipo_documento, numero_documento), cantidad in contador.items():
        if cantidad <= 1:
            continue

        registros = [
            f for f in filas_documentos
            if f.get("tipo_documento_detectado") == tipo_documento
            and f.get("numero_documento_detectado") == numero_documento
        ]

        for fila in registros:
            duplicados.append({
                "tipo_documento": tipo_documento,
                "numero_documento_duplicado": numero_documento,
                "cantidad_veces": cantidad,
                "archivo": fila.get("archivo", ""),
                "carpeta_origen": fila.get("carpeta_origen", ""),
                "ruta_origen": fila.get("ruta_origen", ""),
                "ruta_destino": fila.get("ruta_destino", ""),
            })

    return duplicados


def generar_resumen_txt(
    ruta_resumen: Path,
    total_revisados: int,
    total_copiados: int,
    total_no_validos: int,
    total_error_lectura: int,
    filas_documentos: list[dict],
    faltantes: list[dict],
    duplicados: list[dict],
    sin_numero: list[dict]
):
    tipos = sorted(set([f["tipo_documento_detectado"] for f in filas_documentos if f.get("tipo_documento_detectado")]))

    lineas = []
    lineas.append("RESUMEN DE DOCUMENTOS EXTRAIDOS")
    lineas.append("=" * 60)
    lineas.append(f"Fecha de proceso: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lineas.append(f"PDF revisados: {total_revisados}")
    lineas.append(f"PDF copiados como documentos validos: {total_copiados}")
    lineas.append(f"PDF no identificados como documentos validos: {total_no_validos}")
    lineas.append(f"PDF con error de lectura: {total_error_lectura}")
    lineas.append(f"Documentos con numero detectado: {len([f for f in filas_documentos if f.get('numero_documento_detectado')])}")
    lineas.append(f"Documentos sin numero detectado: {len(sin_numero)}")
    lineas.append(f"Consecutivos faltantes detectados: {len(faltantes)}")
    lineas.append(f"Duplicados detectados: {len(set([(d['tipo_documento'], d['numero_documento_duplicado']) for d in duplicados]))}")
    lineas.append("")
    lineas.append("TIPOS DE DOCUMENTO DETECTADOS")
    lineas.append("-" * 60)

    for tipo in tipos:
        cantidad_tipo = len([f for f in filas_documentos if f.get("tipo_documento_detectado") == tipo])
        lineas.append(f"{tipo}: {cantidad_tipo}")

    lineas.append("")
    lineas.append("DETALLE POR TIPO Y FUENTE")
    lineas.append("-" * 60)

    claves = sorted(set([
        (f.get("tipo_documento_detectado", ""), f.get("fuente", ""))
        for f in filas_documentos
        if f.get("fuente")
    ]))

    for tipo_documento, fuente in claves:
        numeros = []

        for fila in filas_documentos:
            if (
                fila.get("tipo_documento_detectado") == tipo_documento
                and fila.get("fuente") == fuente
                and fila.get("consecutivo_numero") != ""
            ):
                try:
                    numeros.append(int(fila.get("consecutivo_numero")))
                except:
                    pass

        numeros_unicos = sorted(set(numeros))

        if numeros_unicos:
            minimo = min(numeros_unicos)
            maximo = max(numeros_unicos)
            faltantes_grupo = [
                f for f in faltantes
                if f["tipo_documento"] == tipo_documento and f["fuente"] == fuente
            ]
            duplicados_grupo = [
                d for d in duplicados
                if d["tipo_documento"] == tipo_documento
                and d["numero_documento_duplicado"].startswith(f"{fuente}-")
            ]

            lineas.append(f"Tipo: {tipo_documento}")
            lineas.append(f"Fuente: {fuente}")
            lineas.append(f"  Primer consecutivo encontrado: {fuente}-{str(minimo).zfill(10)}")
            lineas.append(f"  Ultimo consecutivo encontrado: {fuente}-{str(maximo).zfill(10)}")
            lineas.append(f"  Cantidad de consecutivos unicos encontrados: {len(numeros_unicos)}")
            lineas.append(f"  Cantidad de faltantes en el rango: {len(faltantes_grupo)}")
            lineas.append(f"  Cantidad de duplicados en el grupo: {len(set([d['numero_documento_duplicado'] for d in duplicados_grupo]))}")
            lineas.append("")

    if faltantes:
        lineas.append("CONSECUTIVOS FALTANTES")
        lineas.append("-" * 60)

        for faltante in faltantes:
            lineas.append(f"{faltante['tipo_documento']} - {faltante['numero_documento_faltante']}")

        lineas.append("")

    if duplicados:
        lineas.append("CONSECUTIVOS DUPLICADOS")
        lineas.append("-" * 60)

        claves_dup = sorted(set([
            (d["tipo_documento"], d["numero_documento_duplicado"])
            for d in duplicados
        ]))

        for tipo_documento, numero in claves_dup:
            cantidad = len([
                d for d in duplicados
                if d["tipo_documento"] == tipo_documento
                and d["numero_documento_duplicado"] == numero
            ])
            lineas.append(f"{tipo_documento} - {numero} - aparece {cantidad} veces")

        lineas.append("")

    if sin_numero:
        lineas.append("PDF IDENTIFICADOS COMO VALIDOS PERO SIN NUMERO DETECTADO")
        lineas.append("-" * 60)

        for fila in sin_numero:
            lineas.append(f"Tipo: {fila.get('tipo_documento_detectado', '')}")
            lineas.append(f"Archivo: {fila.get('archivo', '')}")
            lineas.append(f"Ruta: {fila.get('ruta_origen', '')}")
            lineas.append("")

    ruta_resumen.write_text("\n".join(lineas), encoding="utf-8")


def main():
    origen = Path(RUTA_ORIGEN)

    if not origen.exists():
        print(f"ERROR: La ruta no existe: {origen}")
        input("Presiona Enter para cerrar...")
        return

    destino_general = origen / CARPETA_DESTINO
    destino_general.mkdir(exist_ok=True)

    filas_detalle = []
    filas_documentos = []
    total_revisados = 0
    total_copiados = 0
    total_no_validos = 0
    total_error_lectura = 0

    for archivo in origen.rglob("*"):
        if not archivo.is_file():
            continue

        if destino_general in archivo.parents:
            continue

        if archivo.suffix.lower() not in EXTENSIONES_PERMITIDAS:
            continue

        total_revisados += 1

        texto_pdf = leer_texto_pdf(archivo)

        if texto_pdf.startswith("ERROR_LEYENDO_PDF"):
            total_error_lectura += 1

        coincide_texto, motivo_texto, tipo_documento = validar_texto_documento(texto_pdf)
        coincide_nombre, motivo_nombre = validar_nombre_documento(archivo.name)

        coincide_final = coincide_texto

        estado = "NO COPIADO"
        ruta_destino = ""

        numero_documento, fuente, consecutivo_texto, consecutivo_numero = extraer_numero_documento(texto_pdf)

        fecha_documento = extraer_dato(
            r"FECHA\s*:\s*(\d{4}/\d{2}/\d{2})",
            texto_pdf
        )

        if not fecha_documento:
            fecha_documento = extraer_dato(
                r"Fecha\s*:\s*(\d{4}/\d{2}/\d{2})",
                texto_pdf
            )

        tercero = extraer_dato(
            r"Tercero\s*:\s*([^\n\r]+)",
            texto_pdf
        )

        if not tercero:
            tercero = extraer_dato(
                r"NIT\s*:\s*([^\n\r]+)",
                texto_pdf
            )

        nombre_tercero = extraer_dato(
            r"Nombre\s*:\s*([^\n\r]+)",
            texto_pdf
        )

        if not nombre_tercero:
            nombre_tercero = extraer_dato(
                r"Proveedor\s*:\s*([^\n\r]+)",
                texto_pdf
            )

        unidad_negocio = extraer_dato(
            r"Unidad\s+de\s+Negocio\s*:\s*([^\n\r]+)",
            texto_pdf
        )

        concepto = extraer_dato(
            r"Por\s+Concepto\s*:\s*([^\n\r]+)",
            texto_pdf
        )

        if not concepto:
            concepto = extraer_dato(
                r"Observaciones\s*:\s*([^\n\r]+)",
                texto_pdf
            )

        factura_proveedor = extraer_dato(
            r"FACT\.DE\s+PROVEEDOR\s*:\s*([^\n\r]+)",
            texto_pdf
        )

        estado_documento = extraer_dato(
            r"ESTADO\s*:\s*([^\n\r]+)",
            texto_pdf
        )

        total_comprobante = extraer_total(texto_pdf)

        if coincide_final:
            if RENOMBRAR_CON_NUMERO_DOCUMENTO and numero_documento:
                tipo_limpio = limpiar_nombre_archivo(tipo_documento.replace(" ", "_").replace("/", "-"))
                numero_limpio = limpiar_nombre_archivo(numero_documento.replace(" ", ""))
                nombre_destino = f"{tipo_limpio}_{numero_limpio}_{archivo.name}"
            else:
                nombre_destino = archivo.name

            archivo_destino = nombre_unico(destino_general / nombre_destino)
            shutil.copy2(archivo, archivo_destino)

            total_copiados += 1
            estado = "COPIADO"
            ruta_destino = str(archivo_destino)
        else:
            total_no_validos += 1

        carpeta_relativa = archivo.parent.relative_to(origen)

        fila = {
            "carpeta_origen": str(carpeta_relativa),
            "archivo": archivo.name,
            "extension": archivo.suffix.lower(),
            "es_documento_valido_por_texto_interno": "SI" if coincide_texto else "NO",
            "tipo_documento_detectado": tipo_documento,
            "coincide_nombre_archivo": "SI" if coincide_nombre else "NO",
            "motivo_texto_interno": motivo_texto,
            "motivo_nombre_archivo": motivo_nombre,
            "numero_documento_detectado": numero_documento,
            "fuente": fuente,
            "consecutivo_texto": consecutivo_texto,
            "consecutivo_numero": consecutivo_numero if consecutivo_numero is not None else "",
            "fecha_documento_detectada": fecha_documento,
            "tercero_detectado": tercero,
            "nombre_tercero_detectado": nombre_tercero,
            "unidad_negocio_detectada": unidad_negocio,
            "concepto_detectado": concepto,
            "factura_proveedor_detectada": factura_proveedor,
            "estado_documento_detectado": estado_documento,
            "total_documento_detectado": total_comprobante,
            "estado": estado,
            "ruta_origen": str(archivo),
            "ruta_destino": ruta_destino,
        }

        filas_detalle.append(fila)

        if coincide_final:
            filas_documentos.append(fila)

    filas_documentos_ordenadas = sorted(
        filas_documentos,
        key=lambda x: (
            x.get("tipo_documento_detectado", ""),
            x.get("fuente", ""),
            int(x["consecutivo_numero"]) if str(x.get("consecutivo_numero", "")).isdigit() else 999999999999
        )
    )

    faltantes = calcular_consecutivos_faltantes(filas_documentos_ordenadas)
    duplicados = detectar_duplicados(filas_documentos_ordenadas)
    sin_numero = [f for f in filas_documentos_ordenadas if not f.get("numero_documento_detectado")]

    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")

    reporte_detallado_csv = destino_general / f"informe_documentos_detallado_{fecha}.csv"
    reporte_existentes_csv = destino_general / f"informe_consecutivos_existentes_{fecha}.csv"
    reporte_faltantes_csv = destino_general / f"informe_consecutivos_faltantes_{fecha}.csv"
    reporte_duplicados_csv = destino_general / f"informe_consecutivos_duplicados_{fecha}.csv"
    reporte_sin_numero_csv = destino_general / f"informe_documentos_sin_numero_detectado_{fecha}.csv"
    resumen_txt = destino_general / f"resumen_consecutivos_documentos_{fecha}.txt"

    campos_detalle = [
        "carpeta_origen",
        "archivo",
        "extension",
        "es_documento_valido_por_texto_interno",
        "tipo_documento_detectado",
        "coincide_nombre_archivo",
        "motivo_texto_interno",
        "motivo_nombre_archivo",
        "numero_documento_detectado",
        "fuente",
        "consecutivo_texto",
        "consecutivo_numero",
        "fecha_documento_detectada",
        "tercero_detectado",
        "nombre_tercero_detectado",
        "unidad_negocio_detectada",
        "concepto_detectado",
        "factura_proveedor_detectada",
        "estado_documento_detectado",
        "total_documento_detectado",
        "estado",
        "ruta_origen",
        "ruta_destino",
    ]

    campos_faltantes = [
        "tipo_documento",
        "fuente",
        "consecutivo_numero_faltante",
        "consecutivo_texto_faltante",
        "numero_documento_faltante",
        "rango_validado_desde",
        "rango_validado_hasta",
    ]

    campos_duplicados = [
        "tipo_documento",
        "numero_documento_duplicado",
        "cantidad_veces",
        "archivo",
        "carpeta_origen",
        "ruta_origen",
        "ruta_destino",
    ]

    escribir_csv(reporte_detallado_csv, campos_detalle, filas_detalle)
    escribir_csv(reporte_existentes_csv, campos_detalle, filas_documentos_ordenadas)
    escribir_csv(reporte_faltantes_csv, campos_faltantes, faltantes)
    escribir_csv(reporte_duplicados_csv, campos_duplicados, duplicados)
    escribir_csv(reporte_sin_numero_csv, campos_detalle, sin_numero)

    generar_resumen_txt(
        resumen_txt,
        total_revisados,
        total_copiados,
        total_no_validos,
        total_error_lectura,
        filas_documentos_ordenadas,
        faltantes,
        duplicados,
        sin_numero
    )

    print("Proceso terminado")
    print(f"PDF revisados: {total_revisados}")
    print(f"PDF copiados como documentos validos: {total_copiados}")
    print(f"PDF no identificados como documentos validos: {total_no_validos}")
    print(f"PDF con error de lectura: {total_error_lectura}")
    print(f"Consecutivos faltantes detectados: {len(faltantes)}")
    print(f"Duplicados detectados: {len(set([(d['tipo_documento'], d['numero_documento_duplicado']) for d in duplicados]))}")
    print(f"Documentos sin numero detectado: {len(sin_numero)}")
    print(f"Carpeta destino: {destino_general}")
    print(f"Informe detallado: {reporte_detallado_csv}")
    print(f"Informe existentes: {reporte_existentes_csv}")
    print(f"Informe faltantes: {reporte_faltantes_csv}")
    print(f"Informe duplicados: {reporte_duplicados_csv}")
    print(f"Informe sin numero: {reporte_sin_numero_csv}")
    print(f"Resumen TXT: {resumen_txt}")
    print("Los PDF copiados quedaron directamente dentro de EGRESOS EXTRAIDO.")
    print("La validacion se hizo leyendo el texto interno de cada PDF.")
    input("Presiona Enter para cerrar...")


if __name__ == "__main__":
    main()