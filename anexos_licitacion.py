"""Anexos listos para rellenar dentro de «Sobres para preparar».

Para cada documento de un sobre (Bloque 5 de la revisión) que cita un anexo («Modelo del Anexo 4», «Anexos II y III»):
  1. si el órgano publicó ese anexo como archivo aparte (Anexo II.docx, deuc.xml, anexos.zip…) se copia tal cual
     (si hay Word y PDF del mismo anexo, entra el Word, que se puede rellenar);
  2. si no, se extraen del PCAP las páginas de ese anexo (cabecera «ANEXO N» al inicio de página);
  3. si no se encuentra, no devuelve nada y el documento sigue siendo un .txt, como antes.
Los archivos publicados que no se pueden asignar a ningún documento se devuelven aparte.

Todo es determinista: no gasta tokens de IA.
"""
import io
import re
import zipfile

import requests
from pypdf import PdfReader, PdfWriter

import fetch_licitaciones as f

EDITABLES = {".docx", ".doc", ".odt", ".ods", ".xlsx", ".xls", ".xml", ".rtf", ".dotx", ".ott"}
EXTENSIONES = EDITABLES | {".pdf", ".zip", ".bc3", ".pptx", ".txt", ".csv", ".dwg", ".dxf", ".png", ".jpg", ".jpeg"}
MAX_ARCHIVO = 25 * 1024 * 1024    # anexos y modelos
MAX_OTROS = 15 * 1024 * 1024      # el resto de documentos (precios, planos…): si pesa más, solo se deja el enlace
MAX_TOTAL = 90 * 1024 * 1024
MAX_FICHEROS = 40
MAX_PAGINAS_ANEXO = 40
RE_MODELO = re.compile(r"anexo|deuc|modelo|formulario|declaraci", re.IGNORECASE)

# Número de anexo: arábigo, romano (I, V, X) o letra suelta. Los romanos se pasan a arábigos para que «Anexo IV»
# del PCAP y «Anexo 4» del Bloque 5 sean el mismo.
_NUM = r"(?:N[º°.oO]*\s*)?([IVX]+|\d+|[A-Z])(?![A-Za-z0-9])"
RE_CABECERA = re.compile(r"^\s*(?:ANEXO|Anexo)\s*" + _NUM)
RE_CITA = re.compile(r"(?i:\banexos?)\s*" + _NUM)
RE_SIGUIENTE = re.compile(r"\s*(?:,|;|\by\b|\be\b|\bo\b)\s*(?i:anexos?\s*)?" + _NUM)
RE_ARCHIVO = re.compile(r"(?i:anexos?)[\s_\-.]*" + _NUM)
_ROMANOS = {"I": 1, "V": 5, "X": 10}


def normalizar_numeral(txt):
    t = txt.strip().upper()
    if t.isdigit():
        return str(int(t))
    if re.fullmatch(r"[IVX]+", t):
        total = 0
        for i, c in enumerate(t):
            v = _ROMANOS[c]
            total += -v if i + 1 < len(t) and _ROMANOS[t[i + 1]] > v else v
        return str(total)
    return t


def anexos_citados(texto):
    """Números de anexo citados en un texto, sin repetir y en orden: «Anexos II y III» -> ['2', '3']."""
    salida = []
    for m in RE_CITA.finditer(texto or ""):
        numeros = [m.group(1)]
        pos = m.end()
        while True:
            sig = RE_SIGUIENTE.match(texto, pos)
            if not sig:
                break
            numeros.append(sig.group(1))
            pos = sig.end()
        for n in numeros:
            n = normalizar_numeral(n)
            if n not in salida:
                salida.append(n)
    return salida


def numeral_de_archivo(nombre):
    m = RE_ARCHIVO.search(nombre or "")
    return normalizar_numeral(m.group(1)) if m else None


# ------------------------------------------------------------------ anexos dentro del PCAP
def indice_anexos_pcap(datos):
    """{número de anexo: (primera página, última página)} (páginas desde 1), o {} si no se detecta ninguno.
    Una cabecera cuenta si está en las primeras líneas de la página y es corta; se descartan las páginas de
    índice (varias líneas «ANEXO…») y las citas del índice repetidas más adelante."""
    try:
        pdf = PdfReader(io.BytesIO(datos))
        n = len(pdf.pages)
    except Exception:
        return {}, None
    cabeceras = []
    for i in range(n):
        try:
            texto = pdf.pages[i].extract_text() or ""
        except Exception:
            continue
        lineas = [l for l in texto[:3000].splitlines() if l.strip()]
        if sum(1 for l in lineas if re.match(r"\s*(ANEXO|Anexo)\b", l)) >= 3:
            continue  # índice
        for l in lineas[:8]:
            m = RE_CABECERA.match(l)
            if m and len(l.strip()) <= 110:
                cabeceras.append((i + 1, normalizar_numeral(m.group(1))))
                break
    limpias = [(p, num) for p, num in cabeceras
               if not (p <= 0.4 * n and any(p2 > p and n2 == num for p2, n2 in cabeceras))]
    indice = {}
    for i, (p, num) in enumerate(limpias):
        if num in indice:
            continue
        fin = n
        for p2, n2 in limpias[i + 1:]:
            if n2 != num:
                fin = max(p, p2 - 1)
                break
        indice[num] = (p, min(fin, p + MAX_PAGINAS_ANEXO - 1))
    return indice, pdf


def extraer_paginas(pdf, rangos):
    """PDF (bytes) con las páginas de los rangos (inicio, fin) indicados."""
    escritor = PdfWriter()
    for ini, fin in rangos:
        for p in range(ini - 1, fin):
            escritor.add_page(pdf.pages[p])
    salida = io.BytesIO()
    escritor.write(salida)
    return salida.getvalue()


# ------------------------------------------------------------------ archivos publicados aparte
def _extension_real(datos):
    if datos[:4] == b"%PDF":
        return ".pdf"
    if datos[:4] == b"PK\x03\x04":
        try:
            nombres = zipfile.ZipFile(io.BytesIO(datos)).namelist()
        except Exception:
            return ".zip"
        if any(n.startswith("word/") for n in nombres):
            return ".docx"
        if any(n.startswith("xl/") for n in nombres):
            return ".xlsx"
        if "mimetype" in nombres:
            tipo = zipfile.ZipFile(io.BytesIO(datos)).read("mimetype")
            return ".ods" if b"spreadsheet" in tipo else ".odt"
        return ".zip"
    if datos[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return ".doc"
    if datos.lstrip()[:5] == b"<?xml":
        return ".xml"
    if datos[:5] == b"{\\rtf":
        return ".rtf"
    return ""


def _nombre_con_extension(nombre, datos):
    ext = re.search(r"\.[A-Za-z0-9]{2,5}$", nombre or "")
    if ext and ext.group(0).lower() in EXTENSIONES:
        return nombre
    return (nombre or "documento") + _extension_real(datos)


def descargar(url, limite):
    """(datos, motivo): datos es None si no se pudo bajar o pesa más que el límite."""
    try:
        r = requests.get(url, headers=f.HEADERS, timeout=120, stream=True)
        r.raise_for_status()
        if int(r.headers.get("Content-Length") or 0) > limite:
            return None, "demasiado grande"
        trozos, total = [], 0
        for trozo in r.iter_content(256 * 1024):
            total += len(trozo)
            if total > limite:
                return None, "demasiado grande"
            trozos.append(trozo)
        return b"".join(trozos), None
    except Exception as e:
        return None, f"no se pudo descargar ({e.__class__.__name__})"


def _expandir_zip(datos, tope=40 * 1024 * 1024):
    salida, total = [], 0
    try:
        z = zipfile.ZipFile(io.BytesIO(datos))
        for info in z.infolist():
            base = info.filename.rsplit("/", 1)[-1]
            if info.is_dir() or not base or info.filename.startswith("__MACOSX") or base in (".DS_Store", "Thumbs.db"):
                continue
            total += info.file_size
            if total > tope:
                break
            salida.append((base, z.read(info)))
    except Exception:
        pass
    return salida


def _fichero(nombre, datos, origen):
    nombre = _nombre_con_extension(nombre, datos)
    ext = re.search(r"\.[A-Za-z0-9]{2,5}$", nombre)
    return {
        "nombre": nombre, "ext": ext.group(0).lower() if ext else "", "datos": datos, "origen": origen,
        "numeral": numeral_de_archivo(nombre),
        "deuc": "deuc" in nombre.lower() and "instruc" not in nombre.lower(),  # las instrucciones no son el formulario
    }


def bajar_publicados(docs, avisos):
    """Descarga los documentos adicionales de la licitación y despliega los .zip de anexos/modelos."""
    ficheros, enlaces, total = [], [], 0
    for d in (docs or [])[:MAX_FICHEROS]:
        nombre, url = d.get("nombre") or "documento", d.get("url")
        if not url:
            continue
        es_modelo = bool(RE_MODELO.search(nombre))
        datos, motivo = descargar(url, MAX_ARCHIVO if es_modelo else MAX_OTROS)
        if datos is None:
            enlaces.append((nombre, url))
            avisos.append(f"Documento «{nombre}»: {motivo}; se deja el enlace en «Pliegos/Otros documentos».")
            continue
        total += len(datos)
        if total > MAX_TOTAL:
            avisos.append("Se alcanzó el tope de tamaño de los documentos adicionales; no se incluyen todos.")
            break
        fichero = _fichero(nombre, datos, f"«{nombre}» publicado con el pliego")
        if fichero["ext"] == ".zip" and es_modelo:
            for sub_nombre, sub_datos in _expandir_zip(datos):
                ficheros.append(_fichero(sub_nombre, sub_datos, f"«{sub_nombre}» dentro de «{nombre}»"))
        else:
            ficheros.append(fichero)
    return ficheros, enlaces


# ------------------------------------------------------------------ emparejado con los documentos de los sobres
def _quedarse_con_editables(candidatos):
    """Si un mismo anexo está en Word (editable) y en PDF, solo el Word."""
    grupos = {}
    for c in candidatos:
        grupos.setdefault(c["numeral"] or "deuc", []).append(c)
    salida = []
    for lista in grupos.values():
        editables = [c for c in lista if c["ext"] in EDITABLES]
        salida += editables or lista
    return salida


def preparar_anexos(registro, documentos, pcap, avisos):
    """documentos: lista de documentos (dicts con «nombre» y «detalle») de todos los sobres y del bloque posterior.
    Devuelve (por_documento, otros_anexos, otros_documentos, enlaces):
      por_documento: {id(documento): [{"ext", "datos", "origen"}]}  (solo los que tienen anexo)
      otros_anexos / otros_documentos: [(nombre, datos)] de lo publicado que no se asignó a ningún documento."""
    ficheros, enlaces = bajar_publicados(registro.get("docs_adicionales"), avisos)
    indice, pdf = indice_anexos_pcap(pcap) if pcap else ({}, None)
    usados = set()
    por_documento = {}
    for doc in documentos:
        texto = doc["nombre"] + " " + " ".join(doc.get("detalle") or [])
        citados = anexos_citados(texto)
        es_deuc = "deuc" in doc["nombre"].lower()
        items = []
        candidatos = [c for c in ficheros if (c["numeral"] in citados) or (es_deuc and c["deuc"])]
        candidatos = _quedarse_con_editables(candidatos)
        for c in candidatos:
            usados.add(id(c))
            items.append({"ext": c["ext"], "datos": c["datos"], "origen": c["origen"]})
        cubiertos = {c["numeral"] for c in candidatos if c["numeral"]}
        rangos, etiquetas = [], []
        for n in citados:
            if n not in cubiertos and n in indice and pdf is not None:
                rangos.append(indice[n])
                etiquetas.append(f"Anexo {n}: págs. {indice[n][0]}-{indice[n][1]}")
        if rangos:
            try:
                items.append({"ext": ".pdf", "datos": extraer_paginas(pdf, rangos),
                              "origen": "páginas del PCAP (" + "; ".join(etiquetas) + ")"})
            except Exception as e:
                avisos.append(f"No se pudieron extraer del PCAP las páginas de «{doc['nombre']}» ({e.__class__.__name__}).")
        if items:
            por_documento[id(doc)] = items
    otros_anexos, otros_documentos = [], []
    for c in ficheros:
        if id(c) in usados:
            continue
        (otros_anexos if RE_MODELO.search(c["nombre"]) else otros_documentos).append((c["nombre"], c["datos"]))
    return por_documento, otros_anexos, otros_documentos, enlaces
