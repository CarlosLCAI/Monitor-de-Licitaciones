"""Prepara la carpeta (ZIP) de una licitación a partir de su Revisión IA.

Contenido del ZIP:
  Pliegos/                      PCAP y PPT en PDF, la Revisión IA en PDF y un acceso a la ficha
  <Sobre / Archivo ...>/        un .txt por cada documento que el Bloque 5 de la revisión pide en ese sobre
  Posterior adjudicatario/      un .txt por cada documento que solo se pide al propuesto como adjudicatario
  LEEME.txt                     qué hay en la carpeta y avisos

Los nombres de los sobres son los que da cada pliego (los recoge la IA en el Bloque 5).

Uso:  python preparar_carpeta.py <folder_id> --directorio <carpeta de salida>
Lo lanza el flujo "Preparar carpeta" (visor -> botón "Preparar carpeta"). Necesita Chrome (o Edge) para
convertir la revisión a PDF; si no lo encuentra, incluye la revisión en HTML y lo avisa en el LEEME.
"""
import argparse
import base64
import html as html_lib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import requests

import fetch_licitaciones as f

AQUI = Path(__file__).resolve().parent

# ----------------------------------------------------------------------------------------------
# Bloque 5 de la revisión: sobres y documentos
# ----------------------------------------------------------------------------------------------
RE_BLOQUE5 = re.compile(r"^##\s+Bloque\s*5\b", re.IGNORECASE)
RE_H2 = re.compile(r"^##\s+\S")
RE_H3 = re.compile(r"^###\s+(.*\S)\s*$")
RE_ITEM = re.compile(r"^(\s*)(?:[-*•]|\d+[.)])\s+(.*\S)\s*$")
RE_POSTERIOR = re.compile(
    r"posterior|propuesto como adjudicatari|tras el requerimiento|requerid[oa]|una vez adjudicad", re.IGNORECASE
)
RE_NO_LOCALIZADO = re.compile(r"no localizad", re.IGNORECASE)
# Viñetas que son condiciones o avisos del pliego y no un documento (informes con el formato antiguo, sin negritas)
RE_NOTA = re.compile(
    r"^(sin |no |la falta|los? .{0,40} no se |la documentación acreditativa|dentro de|cualquier|se admite|solo |"
    r"el incumplimiento|la oferta no|todos los|la presentación)", re.IGNORECASE)


def limpiar_md(texto):
    texto = re.sub(r"\*\*|__|`", "", texto)
    return re.sub(r"\s+", " ", texto).strip()


def _nuevo_documento(texto):
    """Nombre del documento = lo que va en **negrita** al principio; si no hay, lo que precede a «:» o « – »."""
    m = re.match(r"^\*\*(.+?)\*\*\s*(.*)$", texto)
    negrita = bool(m)
    if m:
        nombre, resto = m.group(1), m.group(2)
        if len(limpiar_md(nombre).rstrip(":. ")) <= 3 and resto:  # «V1», «A»…: se completa con lo que sigue
            resto_limpio = limpiar_md(resto).lstrip(" ,;:.-–—")
            cola = re.split(r"[.;]", resto_limpio, maxsplit=1)
            nombre = f"{limpiar_md(nombre).rstrip(':. ')} - {cola[0][:60].strip()}"
            resto = cola[1] if len(cola) > 1 else ""
    else:
        partes = re.split(r":\s", texto, maxsplit=1)
        nombre, resto = partes[0], (partes[1] if len(partes) > 1 else "")
        corte = re.split(r",\s", nombre, maxsplit=1)
        if len(corte) > 1 and len(limpiar_md(corte[0])) >= 8:  # «Modelo del Anexo 4, firmado y fechado…»
            nombre, resto = corte[0], corte[1] + (". " + resto if resto else "")
    nombre = limpiar_md(nombre).rstrip(":.,; ")
    if len(nombre) > 85:
        nombre = nombre[:85].rsplit(" ", 1)[0] + "…"
    resto = limpiar_md(resto).lstrip(" ,;:.-–—")
    return {"nombre": nombre or "Documento", "detalle": [resto] if resto else [], "negrita": negrita}


def _separar_notas(documentos):
    """Devuelve (documentos, notas). Con negritas, solo son documentos las viñetas con negrita; sin ellas,
    se apartan las que parecen condiciones o avisos."""
    con_negrita = sum(1 for d in documentos if d["negrita"])
    if documentos and con_negrita / len(documentos) >= 0.6:  # formato nuevo: cada documento va en negrita
        docs = [d for d in documentos if d["negrita"]]
    else:
        docs = [d for d in documentos if not RE_NOTA.search(d["nombre"])]
    notas = [d for d in documentos if d not in docs]
    return docs, [" ".join([n["nombre"]] + n["detalle"]).strip() for n in notas]


def parsear_bloque5(informe):
    """Devuelve {"sobres": [{"nombre", "documentos": [{"nombre", "detalle"}]}], "posterior": [...]}
    o None si el informe no tiene un Bloque 5 utilizable (ningún sobre con documentos)."""
    lineas = (informe or "").splitlines()
    inicio = next((i for i, l in enumerate(lineas) if RE_BLOQUE5.match(l)), None)
    if inicio is None:
        return None
    fin = next((i for i in range(inicio + 1, len(lineas)) if RE_H2.match(lineas[i])), len(lineas))

    elementos = []  # (tipo, sangría, texto)
    for linea in lineas[inicio + 1:fin]:
        h3 = RE_H3.match(linea)
        if h3:
            elementos.append(("h3", 0, h3.group(1)))
            continue
        it = RE_ITEM.match(linea)
        if it:
            elementos.append(("item", len(it.group(1).expandtabs(4)), it.group(2)))
    sangrias = [s for t, s, _ in elementos if t == "item"]
    base = min(sangrias) if sangrias else 0

    sobres, posterior = [], []
    sobre = None
    posterior_global = False
    contenedor_posterior = False
    sangria_hijos = None
    ultimo = None
    for idx, (tipo, sangria, texto) in enumerate(elementos):
        if tipo == "h3":
            titulo = limpiar_md(texto)
            contenedor_posterior = False
            ultimo = None
            if RE_POSTERIOR.search(titulo):
                sobre, posterior_global = None, True
            else:
                sobre = {"nombre": titulo, "documentos": []}
                sobres.append(sobre)
                posterior_global = False
            continue
        if sangria <= base:  # viñeta de primer nivel
            contenedor_posterior = False
            limpio = limpiar_md(texto)
            siguiente_mas_profunda = idx + 1 < len(elementos) and elementos[idx + 1][0] == "item" and elementos[idx + 1][1] > base
            if RE_POSTERIOR.search(limpio) and (siguiente_mas_profunda or limpio.rstrip().endswith(":")):
                contenedor_posterior, sangria_hijos, ultimo = True, None, None
                continue
            if posterior_global:
                destino = posterior
            else:
                if sobre is None:
                    sobre = {"nombre": "Documentación a presentar", "documentos": []}
                    sobres.append(sobre)
                destino = sobre["documentos"]
            ultimo = _nuevo_documento(texto)
            destino.append(ultimo)
        else:  # subviñeta
            if contenedor_posterior:
                if sangria_hijos is None:
                    sangria_hijos = sangria
                if sangria <= sangria_hijos:
                    ultimo = _nuevo_documento(texto)
                    posterior.append(ultimo)
                    continue
            if ultimo is not None:
                ultimo["detalle"].append(limpiar_md(texto))

    for sobre_i in sobres:
        sobre_i["documentos"], sobre_i["notas"] = _separar_notas(sobre_i["documentos"])
    posterior, notas_posterior = _separar_notas(posterior)
    sobres = [s for s in sobres if s["documentos"]]
    todos = [d for s in sobres for d in s["documentos"]]
    if not todos or all(RE_NO_LOCALIZADO.search(d["nombre"]) for d in todos):
        return None
    return {"sobres": sobres, "posterior": posterior, "notas_posterior": notas_posterior}


# ----------------------------------------------------------------------------------------------
# Revisión IA -> HTML -> PDF (mismo aspecto que el botón "Compartir" del visor)
# ----------------------------------------------------------------------------------------------
def _en_linea(texto):
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", html_lib.escape(texto, quote=True))


def markdown_a_html(texto):
    """Misma lógica que formatearInformeIA del visor: títulos ##/###, viñetas, listas y tablas."""
    html = []
    lista = None
    tabla = None

    def cerrar_lista():
        nonlocal lista
        if lista:
            html.append(f"</{lista}>")
            lista = None

    def cerrar_tabla():
        nonlocal tabla
        if not tabla:
            return
        filas = []
        for i, fila in enumerate(tabla):
            celdas = "".join(f"<td>{_en_linea(c.strip())}</td>" for c in fila)
            filas.append(f'<tr{" class=\"informe-tabla-cab\"" if i == 0 else ""}>{celdas}</tr>')
        html.append('<div class="informe-tabla-wrap"><table class="informe-tabla"><tbody>' + "".join(filas) + "</tbody></table></div>")
        tabla = None

    for linea in str(texto).splitlines():
        if re.match(r"^\s*\|.*\|\s*$", linea):
            es_separador = re.match(r"^[\s:|-]+$", linea) and "-" in linea
            if es_separador and tabla:
                continue
            celdas = re.sub(r"^\||\|$", "", linea.strip()).split("|")
            tabla = (tabla or []) + [celdas]
            continue
        if tabla:
            cerrar_tabla()
        m1, m2 = re.match(r"^##\s+(.*)$", linea), re.match(r"^###\s+(.*)$", linea)
        if m1:
            cerrar_lista()
            html.append(f'<p class="informe-h1">{_en_linea(m1.group(1))}</p>')
            continue
        if m2:
            cerrar_lista()
            html.append(f'<p class="informe-h2">{_en_linea(m2.group(1))}</p>')
            continue
        vineta = re.match(r"^\s*[-*]\s+(.*)$", linea)
        numero = re.match(r"^\s*\d+\.\s+(.*)$", linea)
        if vineta or numero:
            etiqueta = "ul" if vineta else "ol"
            if lista != etiqueta:
                cerrar_lista()
                html.append(f"<{etiqueta}>")
                lista = etiqueta
            html.append(f"<li>{_en_linea((vineta or numero).group(1))}</li>")
            continue
        cerrar_lista()
        if not linea.strip():
            continue
        clase = "informe-aviso" if "⚠ REVISAR" in linea else "informe-parrafo"
        html.append(f'<p class="{clase}">{_en_linea(linea)}</p>')
    cerrar_lista()
    cerrar_tabla()
    return "".join(html)


CSS_RESPALDO = """
  body { font-family: Arial, sans-serif; color: #1a1a1a; font-size: 13.5px; line-height: 1.55; max-width: 900px; margin: 0 auto; padding: 24px; }
  .informe-h1 { background: #07AAAA; color: #fff; text-transform: uppercase; font-weight: 700; padding: 7px 12px; margin: 20px 0 8px; }
  .informe-h2 { background: #07AAAA; color: #fff; font-weight: 700; padding: 6px 12px; margin: 12px 0 6px; }
  .informe-tabla { width: 100%; border-collapse: collapse; } .informe-tabla td { border: 1px solid #DADADA; padding: 4px 7px; vertical-align: top; }
  .informe-aviso { color: #ED2E63; font-weight: 600; }
"""


def css_del_visor():
    """El PDF se ve igual que el de «Compartir»: se reutiliza el CSS de index.html."""
    try:
        texto = (AQUI / "index.html").read_text(encoding="utf-8")
        m = re.search(r"const CSS_INFORME_EXPORTADO = `(.*?)`;", texto, re.S)
        if m:
            return m.group(1)
    except OSError:
        pass
    return CSS_RESPALDO


COLORES_SEMAFORO = {
    "verde": ("#DCEEE2", "Verde: nos podemos presentar con la solvencia actual"),
    "naranja": ("#FAE5CF", "Naranja: nos podemos presentar, pero falta algún perfil o parte de la solvencia"),
    "rojo": ("#F5D8DC", "Rojo: no cumplimos los requisitos"),
    "gris": ("#E6E8E9", "Gris: no se puede valorar (falta información en el informe o en el pliego)"),
}


def formato_importe(valor):
    try:
        n = float(valor)
    except (TypeError, ValueError):
        return None
    texto = f"{n:,.2f}"
    return texto.replace(",", "§").replace(".", ",").replace("§", ".") + " €"


def html_revision(registro, rev):
    logo = ""
    ruta_logo = AQUI / "logo-cai.png"
    if ruta_logo.exists():
        logo = "data:image/png;base64," + base64.b64encode(ruta_logo.read_bytes()).decode("ascii")
    datos = [("Expediente", registro.get("folder_id")), ("Órgano", registro.get("organo")),
             ("Importe", formato_importe(registro.get("importe"))), ("Fecha límite", registro.get("fecha_limite"))]
    filas = "".join(f"<tr><td>{html_lib.escape(k)}</td><td>{html_lib.escape(str(v))}</td></tr>" for k, v in datos if v)
    if registro.get("link"):
        filas += f'<tr><td>Ficha</td><td><a href="{html_lib.escape(registro["link"])}">{html_lib.escape(registro["link"])}</a></td></tr>'
    semaforo = ""
    if rev.get("semaforo") in COLORES_SEMAFORO:
        color, texto = COLORES_SEMAFORO[rev["semaforo"]]
        motivo = f"<br>{html_lib.escape(rev['semaforo_motivo'])}" if rev.get("semaforo_motivo") else ""
        semaforo = f'<div style="background:{color};padding:9px 14px;margin:10px 0;border-radius:6px;font-size:13px"><strong>{html_lib.escape(texto)}</strong>{motivo}</div>'
    fecha = (rev.get("fecha_analisis") or "")[:10]
    objeto = (registro.get("descripcion") or registro.get("titulo") or "").strip()
    return (
        '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>Revisión IA</title><style>'
        + css_del_visor() + "</style></head><body>"
        '<div class="cab">' + (f'<img src="{logo}" alt="CAI Consultores">' if logo else "<strong>CAI Consultores</strong>")
        + '<div class="tipo">Revisión IA de licitación</div></div>'
        f"<h1>{html_lib.escape(registro.get('folder_id') or '')}</h1><p class=\"objeto\">{html_lib.escape(objeto)}</p>"
        f'<table class="datos">{filas}</table>{semaforo}'
        + markdown_a_html(rev.get("informe") or "")
        + f'<p class="pie">Informe generado con IA{" el " + fecha if fecha else ""} a partir del pliego de la licitación. '
        "Debe contrastarse con la documentación oficial.</p></body></html>"
    )


def ejecutable_chrome():
    if os.environ.get("CHROME_PATH") and os.path.exists(os.environ["CHROME_PATH"]):
        return os.environ["CHROME_PATH"]
    for nombre in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome"):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    for ruta in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                 r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                 r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                 r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
        if os.path.exists(ruta):
            return ruta
    return None


def html_a_pdf(html_texto):
    """Devuelve los bytes del PDF, o None si no hay navegador o falla la conversión."""
    chrome = ejecutable_chrome()
    if not chrome:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        origen, destino = Path(tmp) / "revision.html", Path(tmp) / "revision.pdf"
        origen.write_text(html_texto, encoding="utf-8")
        comando = [chrome, "--headless=new", "--disable-gpu", "--no-sandbox", "--no-pdf-header-footer",
                   f"--user-data-dir={Path(tmp) / 'perfil'}", f"--print-to-pdf={destino}", origen.as_uri()]
        try:
            subprocess.run(comando, check=True, timeout=120, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            print(f"    Aviso: no se pudo generar el PDF de la revisión ({e}).")
            return None
        return destino.read_bytes() if destino.exists() and destino.stat().st_size > 1000 else None


# ----------------------------------------------------------------------------------------------
# Carpeta
# ----------------------------------------------------------------------------------------------
def sanear_nombre(texto, maximo=80):
    texto = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", str(texto or ""))
    texto = re.sub(r"\s+", " ", texto).strip().strip(".")
    return texto[:maximo].rstrip(" .") or "documento"


def nombre_archivo_zip(folder_id):
    """Nombre del ZIP en la release «carpetas»: el número de expediente. El visor calcula el mismo para enlazarlo."""
    return re.sub(r"[^A-Za-z0-9._-]+", "-", folder_id).strip("-") + ".zip"


def nombre_carpeta_expediente(folder_id):
    """La carpeta se llama como el expediente; los caracteres que Windows no admite (como la «/») pasan a «-»."""
    return re.sub(r'\s+', " ", re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", str(folder_id))).strip(" .-")[:80] or "expediente"


def descargar_pdf(url):
    if not url:
        return None
    try:
        r = requests.get(url, headers=f.HEADERS, timeout=120)
        r.raise_for_status()
        return r.content if r.content[:4] == b"%PDF" else None
    except Exception as e:
        print(f"    Aviso: no se pudo descargar {url[:60]}… ({e})")
        return None


def _txt_documento(registro, ambito, doc, ya_aviso):
    lineas = [doc["nombre"], "", f"Expediente: {registro.get('folder_id')}", f"Dónde se entrega: {ambito}", ""]
    if doc["detalle"]:
        lineas += ["Según la revisión IA del pliego:"] + [f"- {d}" for d in doc["detalle"] if d] + [""]
    lineas += ["(Documento pendiente de preparar. Contrasta siempre lo que pide el pliego.)"]
    return "\r\n".join(lineas) + "\r\n"


def _escribir_documentos(zf, raiz, carpeta, registro, ambito, documentos):
    ancho = max(2, len(str(len(documentos))))
    for i, doc in enumerate(documentos, 1):
        nombre = f"{str(i).zfill(ancho)} - {sanear_nombre(doc['nombre'], 75)}.txt"
        zf.writestr(f"{raiz}/{sanear_nombre(carpeta, 50)}/{nombre}", _txt_documento(registro, ambito, doc, False))


def construir_carpeta(folder_id, directorio):
    encontrados = f._historicos_con_licitacion(folder_id)
    if not encontrados:
        raise SystemExit(f"No se encontró la licitación {folder_id!r} en ningún histórico.")
    registros = [r for _, _, regs in encontrados for r in regs]
    con_revision = [r for r in registros if f.informe_ia_valido(r.get("revision_ia"))]
    if not con_revision:
        raise SystemExit(f"{folder_id}: no tiene una Revisión IA válida; no se puede preparar la carpeta.")
    registro = next((r for r in con_revision if r.get("pcap_url")), con_revision[0])
    rev = registro["revision_ia"]
    bloque5 = parsear_bloque5(rev.get("informe"))
    if bloque5 is None:
        raise SystemExit(f"{folder_id}: el Bloque 5 de la revisión no detalla los sobres y documentos; no se prepara la carpeta.")

    exp = sanear_nombre(folder_id, 40)
    raiz = nombre_carpeta_expediente(folder_id)
    avisos = []
    os.makedirs(directorio, exist_ok=True)
    destino = Path(directorio) / nombre_archivo_zip(folder_id)

    pcap = descargar_pdf(registro.get("pcap_url"))
    ppt = descargar_pdf(registro.get("ppt_url"))
    if not registro.get("pcap_url") or pcap is None:
        avisos.append("No se pudo descargar el PCAP: descárgalo desde la ficha.")
    if not registro.get("ppt_url"):
        avisos.append("La licitación no tiene PPT publicado en el monitor.")
    elif ppt is None:
        avisos.append("No se pudo descargar el PPT: descárgalo desde la ficha.")
    documento_html = html_revision(registro, rev)
    pdf_revision = html_a_pdf(documento_html)
    if pdf_revision is None:
        avisos.append("No había navegador para convertir la revisión a PDF: se incluye en HTML (ábrela en el navegador).")

    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zf:
        if pcap:
            zf.writestr(f"{raiz}/Pliegos/PCAP - {exp}.pdf", pcap)
        if ppt:
            zf.writestr(f"{raiz}/Pliegos/PPT - {exp}.pdf", ppt)
        if pdf_revision:
            zf.writestr(f"{raiz}/Pliegos/Revisión IA - {exp}.pdf", pdf_revision)
        else:
            zf.writestr(f"{raiz}/Pliegos/Revisión IA - {exp}.html", documento_html)
        if registro.get("link"):
            zf.writestr(f"{raiz}/Pliegos/Ficha de la licitación.url", f"[InternetShortcut]\r\nURL={registro['link']}\r\n")
        for sobre in bloque5["sobres"]:
            carpeta = sobre["nombre"].replace("–", "-").replace("—", "-")
            _escribir_documentos(zf, raiz, carpeta, registro, sobre["nombre"], sobre["documentos"])
            if sobre["notas"]:
                cabecera = f"Condiciones y avisos de este sobre según la revisión IA ({sobre['nombre']}):"
                zf.writestr(f"{raiz}/{sanear_nombre(carpeta, 50)}/Notas del pliego.txt",
                            "\r\n".join([cabecera, ""] + [f"- {n}" for n in sobre["notas"]]) + "\r\n")
        if bloque5["posterior"]:
            _escribir_documentos(zf, raiz, "Posterior adjudicatario", registro, "solo si eres el propuesto como adjudicatario (tras el requerimiento)", bloque5["posterior"])
            if bloque5["notas_posterior"]:
                zf.writestr(f"{raiz}/Posterior adjudicatario/Notas del pliego.txt",
                            "\r\n".join(["Condiciones y avisos de la documentación posterior:", ""] + [f"- {n}" for n in bloque5["notas_posterior"]]) + "\r\n")

        leeme = [
            f"CARPETA DE LA LICITACIÓN {folder_id}",
            f"{registro.get('descripcion') or registro.get('titulo') or ''}",
            "",
            f"Órgano de contratación: {registro.get('organo') or '—'}",
            f"Importe sin IVA: {formato_importe(registro.get('importe')) or '—'}",
            f"Fecha límite: {registro.get('fecha_limite') or '—'}",
            f"Ficha: {registro.get('link') or '—'}",
            f"Preparada el {datetime.now(timezone.utc):%d/%m/%Y a las %H:%M} UTC",
            "",
            "CONTENIDO",
            "- Pliegos: PCAP, PPT, Revisión IA y acceso a la ficha.",
        ]
        leeme += [f"- {s['nombre']}: {len(s['documentos'])} documento(s)." for s in bloque5["sobres"]]
        if bloque5["posterior"]:
            leeme.append(f"- Posterior adjudicatario: {len(bloque5['posterior'])} documento(s).")
        leeme += ["", "Los .txt de los sobres son marcadores: un archivo por cada documento que hay que presentar, "
                      "según la Revisión IA. El nombre de cada sobre es el que usa el pliego."]
        if avisos:
            leeme += ["", "AVISOS"] + [f"- {a}" for a in avisos]
        leeme += ["", "La Revisión IA es una ayuda: contrástala siempre con los pliegos oficiales."]
        zf.writestr(f"{raiz}/LEEME.txt", "\r\n".join(leeme) + "\r\n")

    n_docs = sum(len(s["documentos"]) for s in bloque5["sobres"]) + len(bloque5["posterior"])
    print(f"Carpeta preparada: {destino} ({destino.stat().st_size // 1024} KB, {len(bloque5['sobres'])} sobre(s), {n_docs} documento(s))")
    for a in avisos:
        print(f"  Aviso: {a}")
    return destino


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Prepara la carpeta (ZIP) de una licitación")
    parser.add_argument("folder_id")
    parser.add_argument("--directorio", default="carpetas", help="carpeta donde dejar el ZIP")
    args = parser.parse_args()
    construir_carpeta(args.folder_id, args.directorio)
