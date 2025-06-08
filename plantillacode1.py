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


def generar_documento_formato3(lista_licencias: list, datos_extra: dict):
    """
    Formato 3: “Múltiple Compin (sin subrogancia)”
    - lista_licencias: lista de diccionarios, cada uno con llaves:
        id, nombre_titular, rut_titular, escalafon, grado_raw,
        periodo_inicio, periodo_fin, decreto_aut_excel, fecha_decreto_excel, genero, dias
    - datos_extra: dict con llaves:
        decreto_aut_excel, fecha_decreto_excel, secretario
    """
    # Cargamos la plantilla
    plantilla_path = os.path.join("templates", "plantilla_decreto3.docx")
    doc = DocxTemplate(plantilla_path)

    # Extraemos el N° y Fecha de Decretos (desde datos_extra)
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

    # Calculamos ANIO
    try:
        anio = datetime.strptime(fecha_da, "%d/%m/%Y").year
    except:
        anio = datetime.now().year

    # Preparar la sección de “VISTOS:” (letra a, b, c, d, …) con RichText
    vistos_rt = RichText()
    letra_ord = 0
    for lic in lista_licencias:
        letra = chr(ord('a') + letra_ord)
        vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)

        decreto_excel = limpiar_valor_excel(lic.get('decreto_aut_excel'))
        fecha_decreto_excel = limpiar_valor_excel(lic.get('fecha_decreto_excel'))
        if es_valido(decreto_excel) and es_valido(fecha_decreto_excel):
            vistos_rt.add(
                f"Decreto Alcaldicio Nº {decreto_excel} de fecha {fecha_decreto_excel}, "
                f"que tramita Licencia Médica Nº {int(float(lic['id']))} de ",
                font="Arial", size=22
            )
        else:
            vistos_rt.add(
                f"Que tramita Licencia Médica Nº {int(float(lic['id']))} de ",
                font="Arial", size=22
            )
        vistos_rt.add(lic["nombre_titular"], bold=True, font="Arial", size=22)
        vistos_rt.add(
            f" por {lic.get('dias', '')} días a contar del {lic.get('periodo_inicio', '')} "
            f"hasta el {lic.get('periodo_fin', '')} ambas fechas inclusive.\n",
            font="Arial", size=22
        )
        letra_ord += 1

        letra = chr(ord('a') + letra_ord)
        vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)
        vistos_rt.add("Informe de Página virtual de ", font="Arial", size=22)
        vistos_rt.add("COMPIN", bold=True, font="Arial", size=22)
        vistos_rt.add(
            f" que autoriza Licencia Médica Nº {int(float(lic['id']))}.\n",
            font="Arial", size=22
        )
        letra_ord += 1

    # “Y TENIENDO PRESENTE:”
    texto_extra = (
        "Las facultades que me confiere lo establecido en la Ley N° 18.883/89 Estatuto "
        "Administrativo y en la Ley N° 18.695/92 (Refundida) Orgánica Constitucionales "
        "de Municipalidades."
    )

    # Preparar el Decreto principal
    lic0 = lista_licencias[0]
    saludo_tit, rol_tit = determinar_saludo_y_rol(lic0.get("genero", ""))

    if len(lista_licencias) >= 2:
        vistos_letras = "a y c"
        vistos_compin_letras = "b y d"
    else:
        vistos_letras = "a"
        vistos_compin_letras = "b"

    rt_decreto = RichText()
    rt_decreto.add("1. ", bold=True, font="Arial", size=22)
    rt_decreto.add("AUTORIZASE", bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Licencias Médicas que se individualizan en los Vistos letra {vistos_letras} "
        f"a nombre del funcionario ",
        font="Arial", size=22
    )
    rt_decreto.add(lic0["nombre_titular"], bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Rut {lic0.get('rut_titular','')}, {lic0.get('escalafon','')}, grado {int(float(lic0.get('grado_raw',0)))}° de la escala municipal, "
        f"por informes mencionados en los Vistos letra {vistos_compin_letras} emitidos por ",
        font="Arial", size=22
    )
    rt_decreto.add("COMPIN", bold=True, font="Arial", size=22)
    rt_decreto.add(".", font="Arial", size=22)

    distribucion_final = (
        "· - Interesado – Registro SIAPER de la Contraloría General de la República – "
        "Departamento Gestión de Personas – Departamento de Remuneraciones- Oficina de Partes e Informaciones."
    )

    ciudad_base = "PUDAHUEL"
    ciudad_fecha = f"{ciudad_base}, {fecha_da}" if fecha_da else f"{ciudad_base}, {datetime.now().strftime('%d/%m/%Y')}"
    mat_texto = "AUTORIZA LICENCIA MÉDICA"

    contexto = {
        "CIUDAD_FECHA": ciudad_fecha,
        "MAT": mat_texto,
        "VISTOS_RT": vistos_rt,
        "Y_TENIENDO_PRESENTE": texto_extra,
        "TEXTO_DECRETO": rt_decreto,
        "DECRETO_NUM": num_da,
        "DECRETO_FECHA": fecha_da,
        "ANIO": anio,
        "SECRETARIO_NOMBRE": datos_extra.get("secretario", ""),
        "DISTRIBUCION": distribucion_final,
    }

    doc.render(contexto)
    carpeta_decretos = "decretos"
    if not os.path.exists(carpeta_decretos):
        os.makedirs(carpeta_decretos)

    nombre_archivo = f"{anio} D.A. Nº {num_da} de fecha {fecha_da} que autoriza Licencias Médicas Múltiples.docx"
    nombre_archivo = nombre_archivo.replace("/", "-").replace(":", "").replace("|", "")
    ruta_salida = os.path.join(carpeta_decretos, nombre_archivo)
    doc.save(ruta_salida)
    return ruta_salida


def generar_documento_formato4(lista_licencias: list, datos_extra: dict, datos_subrogancia: dict):
    """
    Formato 4: “Subrogancia Múltiple Compin”
    - lista_licencias: lista de diccionarios con varias licencias
    - datos_extra: dict con llaves: decreto_aut_excel, fecha_decreto_excel, secretario
    - datos_subrogancia: dict con llaves:
        nombre_subrogante, genero_subrogante, trato_subrogante,
        cargo_subrogante, direccion_subrogada,
        decreto_subrogancia, fecha_decreto_subrogancia,
        desde_subrogancia, hasta_subrogancia
    """

    plantilla_path = os.path.join("templates", "plantilla_decreto4.docx")
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
        anio = datetime.strptime(fecha_da, "%d/%m/%Y").year
    except:
        anio = datetime.now().year

    vistos_rt = RichText()
    letra_ord = 0

    for lic in lista_licencias:
        letra = chr(ord('a') + letra_ord)
        vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)
        decreto_excel = limpiar_valor_excel(lic.get('decreto_aut_excel'))
        fecha_decreto_excel = limpiar_valor_excel(lic.get('fecha_decreto_excel'))
        if es_valido(decreto_excel) and es_valido(fecha_decreto_excel):
            vistos_rt.add(
                f"Decreto Alcaldicio Nº {decreto_excel} de fecha {fecha_decreto_excel}, "
                f"que tramita Licencia Médica Nº {int(float(lic['id']))} de ",
                font="Arial", size=22
            )
        else:
            vistos_rt.add(
                f"Que tramita Licencia Médica Nº {int(float(lic['id']))} de ",
                font="Arial", size=22
            )
        vistos_rt.add(lic["nombre_titular"], bold=True, font="Arial", size=22)
        vistos_rt.add(
            f" por {lic.get('dias','')} días a contar del {lic.get('periodo_inicio','')} "
            f"hasta el {lic.get('periodo_fin','')} ambas fechas inclusive.\n",
            font="Arial", size=22
        )
        letra_ord += 1

        letra = chr(ord('a') + letra_ord)
        vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)
        vistos_rt.add("Informe de Página virtual de ", font="Arial", size=22)
        vistos_rt.add("COMPIN", bold=True, font="Arial", size=22)
        vistos_rt.add(
            f" que autoriza Licencia Médica Nº {int(float(lic['id']))}.\n",
            font="Arial", size=22
        )
        letra_ord += 1

    letra = chr(ord('a') + letra_ord)
    vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)
    nds = datos_subrogancia.get("decreto_subrogancia", "").strip()
    fds = datos_subrogancia.get("fecha_decreto_subrogancia", "").strip()
    nom_sub = datos_subrogancia.get("nombre_subrogante", "").strip()
    gen_sub = datos_subrogancia.get("genero_subrogante", "").strip()
    trato_sub = datos_subrogancia.get("trato_subrogante", "").strip()
    cargo_sub = datos_subrogancia.get("cargo_subrogante", "").strip()
    dir_sub = datos_subrogancia.get("direccion_subrogada", "").strip()
    desde_sub = datos_subrogancia.get("desde_subrogancia", "").strip()
    hasta_sub = datos_subrogancia.get("hasta_subrogancia", "").strip()

    vistos_rt.add(
        f"Decreto Alcaldicio N° {nds if nds else '[N°]'} de fecha {fds if fds else '[fecha]'}, "
        f"Designa como {cargo_sub if cargo_sub else '[cargo]'} de {dir_sub if dir_sub else '[dirección]'} "
        f"a {trato_sub if trato_sub else '[Sr/Sra]'} {nom_sub if nom_sub else '[nombre]'}, "
        f"a contar del día {desde_sub if desde_sub else '[desde]'} hasta el día {hasta_sub if hasta_sub else '[hasta]'} "
        f"y mientras dure la ausencia del titular.\n",
        font="Arial", size=22
    )
    letra_ord += 1

    letra = chr(ord('a') + letra_ord)
    vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)
    vistos_rt.add(
        "Lo dispuesto en el Decreto Alcaldicio N° 788/83 que autoriza al Secretario Municipal para firmar Decretos de resoluciones de licencias médicas.\n",
        font="Arial", size=22
    )
    letra_ord += 1

    letra = chr(ord('a') + letra_ord)
    vistos_rt.add(f"{letra}) ", bold=True, font="Arial", size=22)
    vistos_rt.add(
        "Resolución N° 573 de fecha 13.12.2014 de la Contraloría General de la República, en relación a los Actos Administrativos a través del Sistema de Registro Electrónico Municipal SIAPER.",
        font="Arial", size=22
    )

    texto_extra = (
        "Las facultades que me confiere lo establecido en la Ley N° 18.883/89 Estatuto "
        "Administrativo y en la Ley N° 18.695/92 (Refundida) Orgánica Constitucionales "
        "de Municipalidades."
    )

    lic0 = lista_licencias[0]
    saludo_tit, rol_tit = determinar_saludo_y_rol(lic0.get("genero", ""))
    if len(lista_licencias) >= 2:
        vistos_letras = "a y c"
        vistos_compin_letras = "b y d"
    else:
        vistos_letras = "a"
        vistos_compin_letras = "b"

    rt_decreto = RichText()
    rt_decreto.add("1. ", bold=True, font="Arial", size=22)
    rt_decreto.add("AUTORIZASE", bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Licencias Médicas que se individualizan en los Vistos letra {vistos_letras} "
        f"a nombre del funcionario ",
        font="Arial", size=22
    )
    rt_decreto.add(lic0["nombre_titular"], bold=True, font="Arial", size=22)
    rt_decreto.add(
        f", Rut {lic0.get('rut_titular','')}, {lic0.get('escalafon','')}, grado {int(float(lic0.get('grado_raw',0)))}° de la escala municipal, "
        f"por informes mencionados en los Vistos letra {vistos_compin_letras} emitidos por ",
        font="Arial", size=22
    )
    rt_decreto.add("COMPIN", bold=True, font="Arial", size=22)
    rt_decreto.add(".", font="Arial", size=22)

    distribucion_final = (
        "· - Interesado – Registro SIAPER de la Contraloría General de la República – "
        "Departamento Gestión de Personas – Departamento de Remuneraciones- Oficina de Partes e Informaciones."
    )

    ciudad_base = "PUDAHUEL"
    ciudad_fecha = f"{ciudad_base}, {fecha_da}" if fecha_da else f"{ciudad_base}, {datetime.now().strftime('%d/%m/%Y')}"
    mat_texto = "AUTORIZA LICENCIA MÉDICA"

    contexto = {
        "CIUDAD_FECHA": ciudad_fecha,
        "MAT": mat_texto,
        "VISTOS_RT": vistos_rt,
        "Y_TENIENDO_PRESENTE": texto_extra,
        "TEXTO_DECRETO": rt_decreto,
        "DECRETO_NUM": num_da,
        "DECRETO_FECHA": fecha_da,
        "ANIO": anio,
        "SECRETARIO_NOMBRE": datos_extra.get("secretario", ""),
        "DISTRIBUCION": distribucion_final,
    }

    doc.render(contexto)
    carpeta_decretos = "decretos"
    if not os.path.exists(carpeta_decretos):
        os.makedirs(carpeta_decretos)

    nombre_archivo = f"{anio} D.A. Nº {num_da} de fecha {fecha_da} que autoriza Licencias Médicas Múltiples (Subrogancia).docx"
    nombre_archivo = nombre_archivo.replace("/", "-").replace(":", "").replace("|", "")
    ruta_salida = os.path.join(carpeta_decretos, nombre_archivo)
    doc.save(ruta_salida)
    return ruta_salida
