from docxtpl import DocxTemplate, RichText
import os
from datetime import datetime

def determinar_saludo_y_rol(genero):
    if not genero:
        return "doña", "de la funcionaria"
    g = genero.strip().upper()
    if g == "M":
        return "don", "del funcionario"
    elif g == "F":
        return "doña", "de la funcionaria"
    else:
        return "don/doña", "del funcionario/de la funcionaria"


def limpiar_valor_excel(valor):
    import math
    if valor is None or (isinstance(valor, float) and math.isnan(valor)):
        return ""
    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return str(valor)
    return str(valor).strip()


def es_valido(valor):
    return valor and str(valor).strip().lower() != "nan"


def generar_documento_desde_plantilla(licencia: dict, datos_extra: dict):
    """
    Formato 1: “Solo Compin (sin subrogancia)”
    - licencia: diccionario con llaves:
        id, nombre_titular, rut_titular, escalafon, grado_raw,
        periodo_inicio, periodo_fin, decreto_aut_excel, fecha_decreto_excel, genero, dias
    - datos_extra: dict con llaves:
        decreto_aut_excel, fecha_decreto_excel, secretario
    """

    # Ruta a la plantilla “plantilla_decreto.docx”
    plantilla_path = os.path.join("templates", "plantilla_decreto.docx")
    doc = DocxTemplate(plantilla_path)

    # Número y fecha de decreto (extraídos de datos_extra)
    raw_da = datos_extra.get("decreto_aut_excel", "")
    raw_fecha_da = datos_extra.get("fecha_decreto_excel", "")
    if raw_da:
        try:
            num_da = str(int(float(raw_da)))
        except:
            num_da = str(raw_da)
    else:
        num_da = ""
    fecha_da = raw_fecha_da or ""

    # Año del decreto (para el nombre de archivo, etc.)
    try:
        anio = datetime.strptime(fecha_da, "%d/%m/%Y").year
    except:
        anio = datetime.now().year

    # ID de la licencia
    try:
        id_lic = str(int(float(licencia["id"])))
    except:
        id_lic = str(licencia["id"])

    # Grado entero
    try:
        grado_int = str(int(float(licencia["grado_raw"])))
    except:
        grado_int = str(licencia["grado_raw"])

    saludo_tit, rol_tit = determinar_saludo_y_rol(licencia.get("genero", ""))

    # VIÑETA A
    rt_v_a = RichText()
    rt_v_a.add("a) ", bold=True, font="Arial", size=22)
    decreto_excel = limpiar_valor_excel(licencia.get("decreto_aut_excel"))
    fecha_decreto_excel = limpiar_valor_excel(licencia.get("fecha_decreto_excel"))
    if es_valido(decreto_excel) and es_valido(fecha_decreto_excel):
        rt_v_a.add(
            f"Decreto Alcaldicio Nº {decreto_excel} de fecha {fecha_decreto_excel}, que tramita Licencia Médica Nº {id_lic} de {saludo_tit} ",
            font="Arial", size=22
        )
    else:
        rt_v_a.add(
            f"Que tramita Licencia Médica Nº {id_lic} de {saludo_tit} ",
            font="Arial", size=22
        )
    rt_v_a.add(licencia["nombre_titular"], bold=True, font="Arial", size=22)
    rt_v_a.add(
        f" por {licencia.get('dias','')} días a contar del {licencia.get('periodo_inicio','')} "
        f"hasta el {licencia.get('periodo_fin','')} ambas fechas inclusive.",
        font="Arial", size=22
    )

    # VIÑETA B
    rt_v_b = RichText()
    rt_v_b.add("b) ", bold=True, font="Arial", size=22)
    rt_v_b.add("Informe de Página virtual de ", font="Arial", size=22)
    rt_v_b.add("COMPIN", bold=True, font="Arial", size=22)
    rt_v_b.add(f" que autoriza Licencia Médica Nº {id_lic}.", font="Arial", size=22)

    # VIÑETA C
    rt_v_c = RichText()
    rt_v_c.add("c) ", bold=True, font="Arial", size=22)
    rt_v_c.add("Lo dispuesto en el Decreto Alcaldicio N° 788/83 que autoriza al ", font="Arial", size=22)
    rt_v_c.add("Secretario", font="Arial", size=22)
    rt_v_c.add(" Municipal para firmar Decretos de resoluciones de licencias médicas.", font="Arial", size=22)

    # VIÑETA D
    rt_v_d = RichText()
    rt_v_d.add("d) ", bold=True, font="Arial", size=22)
    rt_v_d.add(
        "Resolución N° 573 de fecha 13.12.2014 de la Contraloría General de la República, en relación a los Actos Administrativos a través del Sistema de Registro Electrónico Municipal SIAPER.",
        font="Arial", size=22
    )

    # TEXTO DECRETO
    rt_decreto = RichText()
    rt_decreto.add("1. ", bold=True, font="Arial", size=22)
    rt_decreto.add("AUTORIZASE", bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Licencia Médica Nº {id_lic} que otorga reposo médico por {licencia.get('dias','')} días a contar del {licencia.get('periodo_inicio','')} hasta el {licencia.get('periodo_fin','')}, a nombre {rol_tit} ",
        font="Arial", size=22
    )
    rt_decreto.add(licencia["nombre_titular"], bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Rut {licencia.get('rut_titular','')}, {licencia.get('escalafon','')}, grado {grado_int}° de la escala municipal, por informe electrónico emitida por ",
        font="Arial", size=22
    )
    rt_decreto.add("COMPIN", bold=True, font="Arial", size=22)
    rt_decreto.add(".", font="Arial", size=22)

    texto_extra = (
        "Las facultades que me confiere lo establecido en la Ley N° 18.883/89 Estatuto "
        "Administrativo y en la Ley N° 18.695/92 (Refundida) Orgánica Constitucionales "
        "de Municipalidades."
    )
    distribucion_final = (
        "· - Interesado – Registro SIAPER de la Contraloría General de la República – "
        "Departamento Gestión de Personas – Departamento de Remuneraciones- Oficina de Partes e Informaciones."
    )

    contexto = {
        "DECRETO_NUM": num_da,
        "DECRETO_FECHA": fecha_da,
        "ANIO": anio,
        "CIUDAD": "PUDAHUEL",
        "VIÑETA_A": rt_v_a,
        "VIÑETA_B": rt_v_b,
        "VIÑETA_C": rt_v_c,
        "VIÑETA_D": rt_v_d,
        "TEXTO_VISTOS_EXTRA": texto_extra,
        "TEXTO_DECRETO": rt_decreto,
        "SECRETARIO_NOMBRE": datos_extra.get("secretario", ""),
        "DISTRIBUCION": distribucion_final,
        "NOMBRE_TITULAR": licencia.get("nombre_titular", ""),
        "RUT_TITULAR": licencia.get("rut_titular", ""),
        "ESCALAFON": licencia.get("escalafon", ""),
        "GRADO": licencia.get("grado_raw", ""),
        "PERIODO_INICIO": licencia.get("periodo_inicio", ""),
        "PERIODO_FIN": licencia.get("periodo_fin", ""),
        "DIAS": licencia.get("dias", ""),
        "ROL_TITULAR": rol_tit,
    }

    doc.render(contexto)
    carpeta_decretos = "decretos"
    if not os.path.exists(carpeta_decretos):
        os.makedirs(carpeta_decretos)

    nombre_archivo = f"D.A. Nº {num_da} de fecha {fecha_da} que autoriza Licencia Médica de {licencia['nombre_titular']}.docx"
    nombre_archivo = nombre_archivo.replace("/", "-").replace(":", "").replace("|", "")
    ruta_salida = os.path.join(carpeta_decretos, nombre_archivo)
    doc.save(ruta_salida)
    return ruta_salida


def generar_documento_desde_plantilla2(licencia: dict, datos_extra: dict, datos_subrogancia: dict):
    """
    Formato 2: “Solo Compin (con subrogancia)”
    - licencia: diccionario con llaves:
        id, nombre_titular, rut_titular, escalafon, grado_raw,
        periodo_inicio, periodo_fin, decreto_aut_excel, fecha_decreto_excel, genero, dias
    - datos_extra: dict con llaves:
        decreto_aut_excel, fecha_decreto_excel, secretario
    - datos_subrogancia: dict con llaves:
        nombre_subrogante, genero_subrogante, trato_subrogante,
        cargo_subrogante, direccion_subrogada,
        decreto_subrogancia, fecha_decreto_subrogancia,
        desde_subrogancia, hasta_subrogancia
    """

    # Ruta a la plantilla “plantilla_decreto2.docx”
    plantilla_path = os.path.join("templates", "plantilla_decreto2.docx")
    doc = DocxTemplate(plantilla_path)

    raw_da = datos_extra.get("decreto_aut_excel", "")
    raw_fecha_da = datos_extra.get("fecha_decreto_excel", "")
    if raw_da:
        try:
            num_da = str(int(float(raw_da)))
        except:
            num_da = str(raw_da)
    else:
        num_da = ""
    fecha_da = raw_fecha_da or ""

    try:
        anio = datetime.now().year  # en Formato 2 siempre usamos año actual
    except:
        anio = datetime.now().year

    saludo_tit, rol_tit = determinar_saludo_y_rol(licencia.get("genero", ""))

    try:
        id_lic = str(int(float(licencia["id"])))
    except:
        id_lic = str(licencia["id"])

    # VIÑETA C (subrogancia)
    rt_v_c = RichText()
    rt_v_c.add("c) ", bold=True, font="Arial", size=22)
    rt_v_c.add(
        f"Decreto Alcaldicio N° {datos_subrogancia.get('decreto_subrogancia', '[N°]')} de fecha {datos_subrogancia.get('fecha_decreto_subrogancia', '[fecha]')}, "
        f"Designa como {datos_subrogancia.get('cargo_subrogante', '[cargo]')} de {datos_subrogancia.get('direccion_subrogada', '[dirección]')} "
        f"a {datos_subrogancia.get('trato_subrogante', '[Sr/Sra]')} {datos_subrogancia.get('nombre_subrogante', '[nombre]')}, "
        f"a contar del día {datos_subrogancia.get('desde_subrogancia', '[desde]')} hasta el día {datos_subrogancia.get('hasta_subrogancia', '[hasta]')} "
        f"y mientras dure la ausencia del titular.",
        font="Arial", size=22
    )

    # VIÑETA E
    rt_v_e = RichText()
    rt_v_e.add("e) ", bold=True, font="Arial", size=22)
    rt_v_e.add("Lo dispuesto en el Decreto Alcaldicio N° 788/83 que autoriza al ", font="Arial", size=22)
    rt_v_e.add("Secretario", font="Arial", size=22)
    rt_v_e.add(" Municipal para firmar Decretos de resoluciones de licencias médicas.", font="Arial", size=22)

    # VIÑETA A
    rt_v_a = RichText()
    rt_v_a.add("a) ", bold=True, font="Arial", size=22)
    decreto_excel = limpiar_valor_excel(licencia.get("decreto_aut_excel"))
    fecha_decreto_excel = limpiar_valor_excel(licencia.get("fecha_decreto_excel"))
    if es_valido(decreto_excel) and es_valido(fecha_decreto_excel):
        rt_v_a.add(
            f"Decreto Alcaldicio Nº {decreto_excel} de fecha {fecha_decreto_excel}, que tramita Licencia Médica Nº {id_lic} de {saludo_tit} ",
            font="Arial", size=22
        )
    else:
        rt_v_a.add(
            f"Que tramita Licencia Médica Nº {id_lic} de {saludo_tit} ",
            font="Arial", size=22
        )
    rt_v_a.add(licencia["nombre_titular"], bold=True, font="Arial", size=22)
    rt_v_a.add(
        f" por {licencia.get('dias','')} días a contar del {licencia.get('periodo_inicio','')} hasta el {licencia.get('periodo_fin','')} ambas fechas inclusive.",
        font="Arial", size=22
    )

    # VIÑETA B
    rt_v_b = RichText()
    rt_v_b.add("b) ", bold=True, font="Arial", size=22)
    rt_v_b.add("Informe de Página virtual de ", font="Arial", size=22)
    rt_v_b.add("COMPIN", bold=True, font="Arial", size=22)
    rt_v_b.add(f" que autoriza Licencia Médica Nº {id_lic}.", font="Arial", size=22)

    # VIÑETA D
    rt_v_d = RichText()
    rt_v_d.add("d) ", bold=True, font="Arial", size=22)
    rt_v_d.add(
        "Resolución N° 573 de fecha 13.12.2014 de la Contraloría General de la República, en relación a los Actos Administrativos a través del Sistema de Registro Electrónico Municipal SIAPER.",
        font="Arial", size=22
    )

    # TEXTO DECRETO
    rt_decreto = RichText()
    rt_decreto.add("1. ", bold=True, font="Arial", size=22)
    rt_decreto.add("AUTORIZASE", bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Licencia Médica Nº {id_lic} que otorga reposo médico por {licencia.get('dias','')} días a contar del {licencia.get('periodo_inicio','')} hasta el {licencia.get('periodo_fin','')}, a nombre {saludo_tit} ",
        font="Arial", size=22
    )
    rt_decreto.add(licencia["nombre_titular"], bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Rut {licencia.get('rut_titular','')}, {licencia.get('escalafon','')}, grado {int(float(licencia.get('grado_raw',0)))}° de la escala municipal, por informe electrónico emitido por COMPIN y subrogancia mencionada en la viñeta letra c).",
        font="Arial", size=22
    )

    texto_extra = (
        "Las facultades que me confiere lo establecido en la Ley N° 18.883/89 Estatuto "
        "Administrativo y en la Ley N° 18.695/92 (Refundida) Orgánica Constitucionales "
        "de Municipalidades."
    )
    distribucion_final = (
        "· - Interesado – Registro SIAPER de la Contraloría General de la República – "
        "Departamento Gestión de Personas – Departamento de Remuneraciones- Oficina de Partes e Informaciones."
    )

    contexto = {
        "DECRETO_NUM": num_da,
        "DECRETO_FECHA": fecha_da,
        "ANIO": anio,
        "CIUDAD": "PUDAHUEL",
        "VIÑETA_A": rt_v_a,
        "VIÑETA_B": rt_v_b,
        "VIÑETA_C": rt_v_c,
        "VIÑETA_D": rt_v_d,
        "VIÑETA_E": rt_v_e,
        "TEXTO_VISTOS_EXTRA": texto_extra,
        "TEXTO_DECRETO": rt_decreto,
        "SECRETARIO_NOMBRE": datos_extra.get("secretario", ""),
        "DISTRIBUCION": distribucion_final,
        "NOMBRE_TITULAR": licencia.get("nombre_titular", ""),
        "RUT_TITULAR": licencia.get("rut_titular", ""),
        "ESCALAFON": licencia.get("escalafon", ""),
        "GRADO": licencia.get("grado_raw", ""),
        "PERIODO_INICIO": licencia.get("periodo_inicio", ""),
        "PERIODO_FIN": licencia.get("periodo_fin", ""),
        "DIAS": licencia.get("dias", ""),
        "ROL_TITULAR": rol_tit,
    }

    doc.render(contexto)
    carpeta_decretos = "decretos"
    if not os.path.exists(carpeta_decretos):
        os.makedirs(carpeta_decretos)

    nombre_archivo = f"{anio} D.A. Nº {num_da} de fecha {fecha_da} que autoriza Licencia Médica de {licencia['nombre_titular']} (Subrogancia).docx"
    nombre_archivo = nombre_archivo.replace("/", "-").replace(":", "").replace("|", "")
    ruta_salida = os.path.join(carpeta_decretos, nombre_archivo)
    doc.save(ruta_salida)
    return ruta_salida
