import pdfplumber
import re
import pandas as pd
import os


def normalizar_locus(texto):
    texto = texto.upper().replace(" ", "").replace("\n", "")
    if "DR" in texto: return "DR"
    if "DQ" in texto or "DO" in texto: return "DQ"
    return texto.strip()


def limpiar_codigo_corto(codigo):
    if codigo.endswith("I") or codigo.endswith("l"):
        return codigo[:-1] + "1"
    return codigo


def es_alelo_especifico(codigo):
    return not codigo.upper().startswith("XX") and not es_nmdp(codigo)


def es_nmdp(codigo):
    """
    Un código NMDP empieza con una letra (no es XX ni empieza con número).
    Ejemplos válidos: BPMXB, FH, ABCD
    """
    codigo = codigo.strip()
    if not codigo:
        return False
    primera = codigo[0].upper()
    if not primera.isalpha():
        return False
    if codigo.upper().startswith("XX"):
        return False
    return True


def extraer_alelos_nmdp(texto_seccion, codigo_nmdp, grupo):
    """
    Busca el bloque CODIGO_NMDP:=:alelo1/alelo2/... en el texto de la sección
    y construye los alelos específicos combinando el grupo del título con cada alelo.
    Retorna una lista de strings tipo "15:08", "15:479".
    """
    # Escapamos el código NMDP y buscamos el patrón CODIGO:=:num/num/...
    patron = re.escape(codigo_nmdp) + r"\s*:=:\s*([\d/]+)"
    match = re.search(patron, texto_seccion, re.IGNORECASE)
    if not match:
        return []
    alelos_raw = match.group(1).split("/")
    alelos_construidos = []
    for alelo in alelos_raw:
        alelo = alelo.strip()
        if alelo:
            alelos_construidos.append(f"{grupo}:{alelo}")
    return alelos_construidos


def extraer_con_separacion_forzada(texto_seccion, codigo_corto, prefijo_grupo):
    variantes = [codigo_corto]
    if codigo_corto.endswith("1"):
        base = codigo_corto[:-1]
        variantes.append(base + "I")
        variantes.append(base + "l")
    indice_inicio = -1
    longitud_codigo_real = 0
    for var in variantes:
        patron_busqueda = re.escape(var) + r"\W{0,15}\d"
        match = re.search(patron_busqueda, texto_seccion, re.IGNORECASE)
        if match:
            indice_inicio = match.start()
            longitud_codigo_real = len(var)
            break
    if indice_inicio == -1: return "No hallado"
    texto_sucio = texto_seccion[indice_inicio:]
    patron_freno = r"(Locus:|Sample|[ABC]\*|D[RQP][AB]?\d\*|XX[0-9I]+)"
    match_freno = re.search(patron_freno, texto_sucio[longitud_codigo_real:])
    if match_freno:
        punto_de_corte = match_freno.start() + longitud_codigo_real
        texto_sucio = texto_sucio[:punto_de_corte]
    candidatos = re.findall(r'(\d+:\d+[A-Z]*)', texto_sucio)
    alelos_finales = []
    try:
        prefijo_int = int(prefijo_grupo)
    except:
        return "/".join(candidatos)
    for cand in candidatos:
        parte_izq = cand.split(":")[0]
        if parte_izq.isdigit() and int(parte_izq) == prefijo_int:
            alelos_finales.append(cand)
    return "/".join(alelos_finales) if alelos_finales else "Sin datos válidos"


def procesar_pdf_dq_ampliado(ruta_pdf):
    if not os.path.exists(ruta_pdf): return None
    texto_completo = ""
    with pdfplumber.open(ruta_pdf) as pdf:
        for page in pdf.pages:
            texto_completo += page.extract_text(x_tolerance=2) + "\n"
    datos = {"Archivo": os.path.basename(ruta_pdf)}

    match_id = re.search(r"Sample ID\s*[:\.]?\s*(.*?)(?:\n|Local)", texto_completo, re.IGNORECASE)
    datos["Sample ID"] = match_id.group(1).strip() if match_id else "No_Identificado"

    fragmentos = re.split(r"(Locus:\s*[A-Z]+)", texto_completo)
    secciones = {}
    locus_actual = None
    for frag in fragmentos:
        if "Locus:" in frag:
            match_locus = re.search(r"Locus:\s*([A-Z]+)", frag)
            if match_locus: locus_actual = normalizar_locus(match_locus.group(1))
        else:
            if locus_actual:
                if locus_actual not in secciones: secciones[locus_actual] = ""
                secciones[locus_actual] += frag

    loci_std = ["A", "B", "C", "DR", "DQ"]
    for l in loci_std:
        limite_huecos = 4 if l == "DQ" else 2
        for n in range(1, limite_huecos + 1):
            datos[f"{l}_{n}"] = ""
            datos[f"{l}_{n}_Detalle"] = ""

    for locus_nombre, texto_seccion in secciones.items():
        if locus_nombre not in loci_std: continue

        patron_asignacion = r"([A-Z0-9]+)\s*\*\s*(\d+)\s*:\s*([A-Z0-9I:]{2,})"
        coincidencias = re.findall(patron_asignacion, texto_seccion, re.IGNORECASE)
        coincidencias = [m for m in coincidencias if normalizar_locus(m[0]) == locus_nombre]

        limite_extraccion = 4 if locus_nombre == "DQ" else 2

        for idx, match in enumerate(coincidencias[:limite_extraccion]):
            num = idx + 1
            locus_raw, grupo, codigo_sucio = match
            locus_original = locus_raw.replace(" ", "").upper()

            if es_alelo_especifico(codigo_sucio):
                # Caso 1: alelo específico directo, ej: A*02:01
                datos[f"{locus_nombre}_{num}"] = f"{locus_original}*{grupo}:{codigo_sucio}"

            elif es_nmdp(codigo_sucio):
                # Caso 2: código NMDP, ej: B*15:BPMXB
                codigo_nmdp = codigo_sucio.strip()
                datos[f"{locus_nombre}_{num}"] = f"{locus_original}*{grupo}:{codigo_nmdp}"

                alelos_resueltos = extraer_alelos_nmdp(texto_seccion, codigo_nmdp, grupo)
                if alelos_resueltos:
                    datos[f"{locus_nombre}_{num}_Detalle"] = "/".join(alelos_resueltos)
                else:
                    datos[f"{locus_nombre}_{num}_Detalle"] = "Sin datos NMDP"

            else:
                # Caso 3: código ambiguo XX, ej: DRB1*11:XX01
                codigo_corto = limpiar_codigo_corto(codigo_sucio)
                datos[f"{locus_nombre}_{num}"] = f"{locus_original}*{grupo}:{codigo_corto}"
                datos[f"{locus_nombre}_{num}_Detalle"] = extraer_con_separacion_forzada(
                    texto_seccion, codigo_corto, grupo
                )

    return datos


def ejecutar_procesamiento(lista_archivos):
    resultados = []
    nombre_paciente = None

    for archivo in lista_archivos:
        info = procesar_pdf_dq_ampliado(archivo)
        if info:
            resultados.append(info)
            if nombre_paciente is None:
                nombre_paciente = info.get("Sample ID", "Paciente_Desconocido")

    if resultados:
        df = pd.DataFrame(resultados)
        nombre_csv = "resultados_hla_final_limpio.csv"
        df.to_csv(nombre_csv, index=False, encoding='utf-8-sig')
        return nombre_paciente
    return None