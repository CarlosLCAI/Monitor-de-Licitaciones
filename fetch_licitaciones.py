import requests
from lxml import etree
from pypdf import PdfReader
from io import BytesIO
import base64
import json
import os
import re
from datetime import datetime, timedelta, timezone

NS = {
    'atom': 'http://www.w3.org/2005/Atom',
    'cbc': 'urn:dgpe:names:draft:codice:schema:xsd:CommonBasicComponents-2',
    'cac': 'urn:dgpe:names:draft:codice:schema:xsd:CommonAggregateComponents-2',
    'cac-place-ext': 'urn:dgpe:names:draft:codice-place-ext:schema:xsd:CommonAggregateComponents-2',
    'cbc-place-ext': 'urn:dgpe:names:draft:codice-place-ext:schema:xsd:CommonBasicComponents-2',
}

FEED_URL = "https://contrataciondelestado.es/sindicacion/sindicacion_643/licitacionesPerfilesContratanteCompleto3.atom"
# Feeds de la plataforma. Cada uno publica licitaciones distintas: el de perfiles no incluye las
# de plataformas autonómicas agregadas (p. ej. toda la Junta de Andalucía, expedientes "CONTR ...")
# ni los contratos menores. El tercer valor indica si un fallo al leerlo aborta la ejecución
# (el principal sí; los secundarios solo avisan y se sigue con el resto).
FEED_AGREGADAS_URL = "https://contrataciondelestado.es/sindicacion/sindicacion_1044/PlataformasAgregadasSinMenores.atom"
FEED_MENORES_URL = "https://contrataciondelestado.es/sindicacion/sindicacion_1143/contratosMenoresPerfilesContratantes.atom"
FEEDS = [
    ("perfiles", FEED_URL, True),
    ("agregadas", FEED_AGREGADAS_URL, False),
    ("menores", FEED_MENORES_URL, False),
]
VISOR_URL = "https://carloslcai.github.io/Monitor-de-Licitaciones/"
VENTANA_HORAS = 72  # cubre con margen el hueco viernes tarde -> lunes madrugada (fines de semana con poca/nula publicación)
MAX_PAGINAS = 30
FILTROS_MANIFEST_FILE = "filtros.json"
FILTROS_CARPETA_BASE = "filtros"

# Nombres de archivo estándar dentro de la carpeta de cada filtro (filtros/<id>/...).
NOMBRE_CONFIG = "config.json"
NOMBRE_ESTADO = "estado.json"
NOMBRE_HISTORICO = "historico.json"
NOMBRE_RESULTADO_HOY = "resultado_hoy.json"
NOMBRE_ULTIMA_LECTURA = "ultima_lectura.json"

# Se usa solo si la carpeta del filtro "1" no tiene config.json todavía (primera ejecución).
# A partir de ahí, filtros/<id>/config.json es la fuente de verdad y se edita desde el visor.
FILTRO_CONFIG_POR_DEFECTO = {
    "nombre": "Diseño Urbano Andalucía",
    "nuts_prefix": "ES61",
    "region": "Andalucía",
    "cpv_permitidos": ["71400000", "71410000", "71420000", "71222000", "71222100", "71222200", "71240000", "71241000", "71243000", "71245000", "71510000", "90712100"],
    "estados_permitidos": ["PUB"],
    "tipos_contrato_permitidos": [],   # códigos de ProcurementProject/TypeCode; vacío = todos
    "procedimientos_permitidos": [],   # códigos de TenderingProcess/ProcedureCode; vacío = todos
    "importe_min": None,
    "importe_max": None,
}

# Config por defecto para un filtro nuevo añadido al manifiesto que aún no tiene archivo propio:
# sin ningún criterio activo (coincide con todo) hasta que se configure desde el visor.
FILTRO_CONFIG_VACIO = {
    "nombre": "Nuevo filtro",
    "nuts_prefix": "",
    "region": "",
    "cpv_permitidos": [],
    "estados_permitidos": [],
    "tipos_contrato_permitidos": [],
    "procedimientos_permitidos": [],
    "importe_min": None,
    "importe_max": None,
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "application/atom+xml, application/xml, text/xml, */*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
    "Referer": "https://contrataciondelestado.es/",
}

# Revisión de solvencia por IA (opcional): se activa solo si las variables de entorno
# SOLVENCIA_EMPRESA y ANTHROPIC_API_KEY están configuradas (secrets de GitHub Actions).
# Usa la API directa de Anthropic (Claude), mandando los PDF completos como documentos
# (no texto extraído) para que Claude pueda leer visualmente páginas escaneadas sin capa
# de texto — pypdf se quedaría en blanco con esos documentos.
ANTHROPIC_API_ENDPOINT = "https://api.anthropic.com/v1/messages"
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5")
# Modelo barato usado solo para localizar en qué páginas del PCAP está cada apartado, antes de
# mandarle el documento completo al modelo caro — ver localizar_paginas_relevantes().
ANTHROPIC_MODEL_LOCALIZADOR = os.environ.get("ANTHROPIC_MODEL_LOCALIZADOR", "claude-haiku-5-5")
ANTHROPIC_VERSION = "2023-06-01"

MARCA_INFORME_CORTADO = "⚠ REVISAR — informe cortado por límite de longitud de la IA, puede estar incompleto."

PROMPT_SISTEMA_INFORME = (
    "Eres un asistente técnico especializado en licitaciones públicas españolas de obra civil, "
    "edificación e infraestructuras, para el equipo de CAI Consultores (arquitectura/ingeniería). "
    "Analizas el PCAP (pliego de cláusulas administrativas particulares) de una licitación y, si se "
    "incluye, su PPT (pliego de prescripciones técnicas), y "
    "extraes la información clave en un informe estructurado en 6 bloques, en este orden. Genera "
    "los 6 bloques completos en una sola respuesta: esto se ejecuta de forma automática, sin "
    "nadie revisando ni validando entre bloques.\n\n"
    "El texto del PCAP que recibes puede venir recortado a las páginas más relevantes (verás "
    "marcas \"--- Página N ---\"); si no encuentras algo, no asumas que no existe en el documento "
    "completo — indícalo como no localizado en vez de darlo por ausente.\n\n"
    "Si recibes también el texto del PPT (aparece como «TEXTO DEL PPT»), úsalo junto con el PCAP: "
    "aporta el detalle técnico (trabajos y entregables, equipo mínimo y titulaciones, plazos "
    "parciales, medios, normativa). Intégralo en los bloques que correspondan indicando «(según el "
    "PPT)» cuando un dato proceda solo de él, y señala con \"⚠ REVISAR\" cualquier discrepancia entre "
    "PCAP y PPT. Si no se incluye PPT, di en el Bloque 1 que su detalle técnico no se ha analizado.\n\n"
    "Si el PCAP incluye un Cuadro de Características del Contrato (o cuadro resumen / anexo de "
    "características), úsalo como fuente principal: suele resumir casi todo el contrato. "
    "Contrástalo con las cláusulas del cuerpo del pliego y, si hay alguna discrepancia, "
    "señálala con \"⚠ REVISAR\" indicando qué dice cada parte.\n\n"
    "FORMATO: responde en Markdown. Usa exactamente '## Bloque N – Título' para cada uno de los "
    "6 títulos de bloque, y '### ' para los subtítulos interiores de cada bloque. Usa **negrita** "
    "para resaltar cifras y datos clave, viñetas '- ' o listas numeradas '1. ' donde corresponda, "
    "y tablas Markdown (fila de cabecera + fila separadora '---') para cualquier dato tabular "
    "(presupuestos por lote, baremos de puntuación, etc.). No uses HTML ni describas colores o "
    "tipografías — el formato visual lo aplica el visor, tú solo estructuras el contenido.\n\n"
    "## Bloque 1 – Características Principales del Contrato\n"
    "### Objeto de la licitación\n"
    "Descripción precisa (50-70 palabras) y lista de objetivos.\n"
    "### División en lotes\n"
    "Lista numerada de lotes, o indica que no hay división en lotes.\n"
    "### Características del contrato\n"
    "Presupuesto base (por lote si aplica), duración, plazo/lugar/forma de presentación de "
    "ofertas, y trabajos específicos obligatorios (arqueología, topografía, geotecnia, trámites "
    "ambientales…).\n\n"
    "## Bloque 2 – Solvencia\n"
    "### Económica y financiera\n"
    "Solvencia económica mínima exigida, cifra de negocio solicitada, y seguros u otra "
    "acreditación requerida.\n"
    "### Técnica y profesional\n"
    "Titulación exigida por perfil, experiencia mínima y su forma de acreditación, y proyectos de "
    "referencia exigidos (tipología, importe). Indica el equipo mínimo por lote y qué puestos "
    "pueden subcontratarse. Si existe un criterio de mejora por experiencia adicional del equipo "
    "distinto del mínimo de solvencia, márcalo aparte con \"⚠ REVISAR — criterio de mejora de "
    "equipo, no confundir con el mínimo de solvencia\".\n\n"
    "## Bloque 3 – Criterios de Adjudicación\n"
    "### Automáticos (fórmulas)\n"
    "Detalla la fórmula económica y cualquier otro criterio automático.\n"
    "### Juicio de valor\n"
    "Desglose y subapartados de metodología, con su puntuación.\n"
    "### Umbrales mínimos de puntuación\n"
    "Indica si existen, o que no los hay.\n\n"
    "## Bloque 4 – Forma de Pago\n"
    "Valoración de trabajos, certificaciones y facturación, y condiciones adicionales "
    "(penalizaciones, revisión de precios, etc.).\n\n"
    "## Bloque 5 – Documentación a Presentar\n"
    "Desglosa un sobre (o archivo electrónico) por cada uno que haya que entregar, con un subtítulo "
    "'### <nombre exacto que le da el pliego> – <qué contiene>' (por ejemplo '### Sobre A – "
    "Documentación administrativa', '### Sobre 2 – Proposición técnica' o '### Archivo electrónico "
    "nº 1 – Documentación administrativa'). Dentro de cada uno, una viñeta '- ' por cada documento "
    "que hay que presentar, con el nombre del documento en **negrita** (por ejemplo **DEUC**, "
    "**Anexo 2 – Modelo de oferta**, **Memoria técnica**) y, tras él, una breve descripción de lo "
    "que debe contener; si necesitas detallar, usa subviñetas. La documentación que solo se exige "
    "al licitador propuesto como adjudicatario (tras el requerimiento) va en un subtítulo aparte "
    "'### Documentación posterior (solo el propuesto como adjudicatario)', también en viñetas. Si el "
    "pliego no detalla los sobres, indícalo como no localizado en vez de inventarlo.\n\n"
    "## Bloque 6 – Conclusiones / Análisis Preliminar\n"
    "Analiza el precio de licitación en relación con la duración del contrato. Compara los "
    "requisitos de solvencia económica y técnica de los Bloques 2 y 3 contra el perfil de la "
    "empresa que se te facilita en el mensaje (equipo, experiencia, cifra de negocio), y valora "
    "si CAI Consultores cumple los mínimos o qué le faltaría acreditar.\n"
    "### Semáforo de presentación\n"
    "Valora SOLO si CAI puede presentarse con su perfil de solvencia (admisibilidad); no valores "
    "precio, baja, plazo ni probabilidad de ganar. Compara cada requisito (solvencia económica, "
    "experiencia en trabajos similares, equipo mínimo y titulaciones/colegiación, seguros, "
    "clasificación) con el perfil, aplicando los CRITERIOS DE LA DIRECCIÓN que figuran al principio "
    "del perfil. Asigna UN color:\n"
    "- VERDE: con el perfil cumplimos todos los requisitos y podemos presentarnos en solitario; o "
    "ya hemos ganado un contrato muy similar por objeto (consta en el perfil), por lo que la solvencia exigida "
    "es equivalente a la que ya acreditamos.\n"
    "- NARANJA: podemos presentarnos, pero falta algo que no es la mayoría de los requisitos "
    "(por ejemplo 1 de 3 o 2 de 5 perfiles, o una parte de la solvencia) y lo que falta es "
    "subcontratable o cubrible con medios externos o UTE según el propio pliego; o el único "
    "precedente similar lo ganamos yendo en UTE.\n"
    "- ROJO: no cumplimos: falta la mayoría de los requisitos, o falta algo imprescindible que el "
    "pliego no permite subcontratar ni integrar con medios externos, o la solvencia económica es insuficiente.\n"
    "- GRIS: no se puede valorar porque el texto recibido no incluye los requisitos de solvencia "
    "o de equipo (páginas no localizadas, pliego incompleto) o el perfil no permite compararlos. Gris NO "
    "significa que no cumplamos: úsalo en lugar de rojo cuando falte información.\n"
    "Detalla en el informe: el color elegido, cuántos requisitos se cumplen del total, qué falta y "
    "si el pliego permite subcontratarlo. Los datos del perfil marcados [SIN VERIFICAR] o que no "
    "consten son carencia por acreditar, no incumplimiento definitivo. Si el pliego exige "
    "clasificación empresarial y el perfil no consta que la tengamos, señálalo con \"⚠ REVISAR — "
    "criterio pendiente de decidir\" y no des rojo solo por esto.\n\n"
    "DATOS PARA EL EXCEL (obligatorio): justo ANTES de la línea final, añade un único bloque con "
    "este formato exacto (JSON válido: comillas dobles, sin comas finales, sin comentarios; usa null "
    "o \"\" cuando el dato no conste en los pliegos; textos breves y literales, sin repetir el informe; "
    "los saltos de línea dentro de un texto, como \\n):\n"
    "<<<DATOS_EXCEL\n"
    "{\"info\": {\"valor_estimado\": <número sin IVA o null>, \"iva_pct\": <número, p. ej. 21, o null>, "
    "\"fecha_presentacion\": \"AAAA-MM-DD\" o null, \"plazo_ejecucion\": \"\", \"consultas\": \"\", "
    "\"garantia_provisional\": \"\", \"garantia_definitiva\": \"\", \"seguro_rc\": \"\", \"lotes\": \"\", "
    "\"revision_precios\": \"\", \"plazo_garantia\": \"\", \"ute\": \"\", \"subcontratacion\": \"\", "
    "\"umbral_tecnico\": \"\"},\n"
    "\"solvencia\": {\"nota\": \"\", \"economica\": [ITEM], \"tecnica\": [ITEM], \"equipo\": [ITEM]},\n"
    "\"juicio_valor\": [{\"nombre\": \"\", \"puntos\": <número o null>, \"paginas_a4\": <número o null>, "
    "\"paginas_a3\": <número o null>, \"formato\": \"\", \"acreditacion\": \"\", \"apartados\": "
    "[{\"nombre\": \"\", \"puntos\": <número o null>}]}],\n"
    "\"formulas\": [{\"nombre\": \"\", \"puntos\": <número o null>, \"formula\": \"\", \"acreditacion\": \"\"}],\n"
    "\"otros\": [{\"nombre\": \"\", \"puntos\": <número o null>, \"acreditacion\": \"\", \"apartados\": "
    "[{\"nombre\": \"\", \"puntos\": <número o null>, \"regla\": \"\"}]}],\n"
    "\"umbral\": \"\"}\n"
    "DATOS_EXCEL>>>\n"
    "donde ITEM = {\"nombre\": \"requisito en pocas palabras\", \"valor\": \"cifra o exigencia\", "
    "\"acreditacion\": \"cómo se acredita\", \"cai\": \"cumple|parcial|no|no_consta\", \"cai_nota\": "
    "\"motivo en pocas palabras\"} (omite \"cai\" y \"cai_nota\" si no se puede comparar con el perfil; "
    "usa \"cai\" según tu comparación del Bloque 6). Criterios: solvencia.economica = volumen de "
    "negocio, seguros y similares; solvencia.tecnica = trabajos similares, certificaciones (ISO) y "
    "demás exigencias de la empresa; solvencia.equipo = un ITEM por puesto del equipo mínimo (nombre = "
    "el puesto, valor = titulación y experiencia, acreditacion = cómo se justifica); solvencia.nota = "
    "reglas generales (UTE, puestos no acumulables…). juicio_valor = un elemento por memoria o "
    "documento que se valora por juicio de valor, con su tope de puntos, el máximo de páginas A4 y A3 "
    "que fija el pliego, el formato exigido (letra, interlineado…) y un apartado por cada "
    "subcriterio con sus puntos; si el pliego no desglosa subcriterios, deja apartados vacío. "
    "formulas = criterios que se puntúan con una fórmula (precio, plazos…): escribe la fórmula en "
    "texto plano con sus variables definidas (por ejemplo \"PO = (PM × MO) / OV; PO = puntuación "
    "obtenida; PM = puntuación máxima…\"). otros = criterios automáticos que no son una fórmula "
    "(experiencia de perfiles, mejoras del equipo, compromisos…): un apartado por perfil o concepto, "
    "con su tope y la regla de puntuación (\"2 puntos/plan, hasta 4\").\n\n"
    "LÍNEA FINAL (obligatoria): termina tu respuesta con una última línea, sola, exactamente con "
    "este formato (sin negritas ni comillas): SEMAFORO: <verde|naranja|rojo|gris> | <motivo en una "
    "sola frase de 25 palabras como máximo>\n\n"
    "Reglas generales: responde en español técnico, basándote solo en el PCAP y el PPT proporcionados "
    "(nunca extrapoles de otros pliegos), sé riguroso con baremos y limitaciones de formato, y "
    "ante cualquier duda o incoherencia en el documento, no la inventes — indícalo y sugiere "
    "revisar ese punto concreto del PCAP."
)

# --- Localización de secciones relevantes (paso barato con Haiku) ------------------------------
# En vez de mandar el PCAP entero al modelo caro, primero se le pasa a un modelo barato un
# resumen (solo el inicio de cada página) para que diga en qué páginas está cada apartado del
# informe. Con eso se construye un documento reducido con solo esas páginas (+ margen), que es
# lo que de verdad se manda a generar el informe. Si esta llamada falla por lo que sea, se cae
# al recorte por Anexo I/III de siempre (ver recortar_preservando_anexos).

PROMPT_SISTEMA_LOCALIZADOR = (
    "Eres un asistente que localiza en qué páginas de un PCAP (pliego de cláusulas "
    "administrativas particulares) de una licitación pública española aparece cada uno de los "
    "apartados indicados en la herramienta. Se te da, página por página, solo el inicio del "
    "texto de cada una, no el documento completo. Devuelve los números de página (empezando en "
    "1) donde aparece cada apartado; una misma página puede repetirse en varios apartados, y "
    "deja la lista vacía si no la localizas. No inventes páginas que no veas en el texto "
    "proporcionado."
)

HERRAMIENTA_LOCALIZADOR = {
    "name": "localizar_secciones",
    "description": "Registra en qué páginas del PCAP aparece cada apartado del informe.",
    "input_schema": {
        "type": "object",
        "properties": {
            "objeto_y_caracteristicas": {
                "type": "array", "items": {"type": "integer"},
                "description": "Objeto, lotes, presupuesto, duración, plazos y trabajos "
                               "específicos obligatorios (arqueología, topografía, geotecnia, "
                               "trámites ambientales…).",
            },
            "solvencia": {
                "type": "array", "items": {"type": "integer"},
                "description": "Solvencia económica/financiera y técnica/profesional, equipo "
                               "mínimo, titulaciones y experiencia exigida.",
            },
            "criterios_adjudicacion": {
                "type": "array", "items": {"type": "integer"},
                "description": "Criterios de adjudicación automáticos (fórmulas) y de juicio de "
                               "valor.",
            },
            "forma_pago": {
                "type": "array", "items": {"type": "integer"},
                "description": "Forma de pago, valoración de trabajos, certificaciones, "
                               "facturación, penalizaciones, revisión de precios.",
            },
            "documentacion_a_presentar": {
                "type": "array", "items": {"type": "integer"},
                "description": "Sobres y documentación que hay que presentar para licitar.",
            },
        },
        "required": [
            "objeto_y_caracteristicas", "solvencia", "criterios_adjudicacion", "forma_pago",
            "documentacion_a_presentar",
        ],
    },
}


# Cuántos caracteres del comienzo de cada página ve el localizador (Haiku, barato: 60 páginas × 350
# caracteres son unos 6.000 tokens).
LONGITUD_INICIO_PAGINA = 350


def quitar_lineas_repetidas(paginas_texto, fraccion_minima=0.3, paginas_minimas=4):
    """Devuelve las páginas sin las líneas que se repiten en muchas de ellas (cabeceras y pies:
    sello de firma electrónica con CSV y enlace de validación, 'Página 3 de 60', nombre del
    órgano…). Se usa solo para el resumen del localizador: si la cabecera de cada página ocupa
    más de 220 caracteres, el inicio de todas las páginas sería idéntico y el localizador no vería
    nada útil (pasó con un PCAP de la Autoridad Portuaria de Sevilla). Las líneas se comparan sin
    cifras, para que 'Página 2 de 60' y 'Página 3 de 60' cuenten como la misma."""
    def clave(linea):
        return re.sub(r"\d+", "#", re.sub(r"\s+", " ", linea.strip()))

    if len(paginas_texto) < paginas_minimas:
        return list(paginas_texto)
    paginas_con_linea = {}
    for texto in paginas_texto:
        for k in {clave(l) for l in (texto or "").splitlines() if l.strip()}:
            paginas_con_linea[k] = paginas_con_linea.get(k, 0) + 1
    umbral = max(paginas_minimas, fraccion_minima * len(paginas_texto))
    repetidas = {k for k, n in paginas_con_linea.items() if n >= umbral}
    return [
        "\n".join(l for l in (texto or "").splitlines() if l.strip() and clave(l) not in repetidas)
        for texto in paginas_texto
    ]


def localizar_paginas_relevantes(paginas_texto, api_key):
    """Llamada barata (Haiku) que decide qué páginas hacen falta leer para cada apartado del
    informe, a partir de un resumen de cada página (no el documento entero). Lanza una excepción
    si algo falla; el llamador cae entonces al recorte por Anexo I/III de siempre."""
    inicios = [re.sub(r"\s+", " ", texto).strip()[:LONGITUD_INICIO_PAGINA] for texto in quitar_lineas_repetidas(paginas_texto)]
    indice = "\n".join(f"[Página {i + 1}] {inicio}" for i, inicio in enumerate(inicios))
    indice = indice[:60000]  # tope de seguridad para documentos con muchísimas páginas

    payload = {
        "model": ANTHROPIC_MODEL_LOCALIZADOR,
        "max_tokens": 1024,
        # Sin cache_control aquí: Haiku 4.5 exige 4.096 tokens mínimo para cachear, y este
        # prompt del sistema (~130 tokens) más la herramienta no se acercan a eso — cachearlo
        # no haría nada (la API lo ignora sin más), así que no lo declaramos.
        "system": PROMPT_SISTEMA_LOCALIZADOR,
        "tools": [HERRAMIENTA_LOCALIZADOR],
        "tool_choice": {"type": "tool", "name": "localizar_secciones"},
        "messages": [{
            "role": "user",
            "content": f"Documento de {len(paginas_texto)} páginas:\n\n{indice}",
        }],
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
        "Content-Type": "application/json",
    }
    resp = requests.post(ANTHROPIC_API_ENDPOINT, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    bloque = next(b for b in data["content"] if b.get("type") == "tool_use")
    return bloque["input"]


def construir_texto_reducido(paginas_texto, secciones, margen=1, max_chars=200000):
    """A partir de las páginas que ha localizado la IA (+ 1 página de margen a cada lado),
    construye el texto reducido que se manda a generar el informe. Siempre incluye las primeras
    páginas (portada/índice/objeto) y cualquier anexo I/III (su encabezado + 3 páginas siguientes),
    aunque el localizador no lo haya marcado, por si el localizador se equivoca."""
    total_paginas = len(paginas_texto)
    incluidas = set(range(1, min(3, total_paginas) + 1))

    for paginas in secciones.values():
        for p in paginas:
            for pp in range(p - margen, p + margen + 1):
                if 1 <= pp <= total_paginas:
                    incluidas.add(pp)

    for i, texto in enumerate(paginas_texto):
        if PATRON_CABECERA_ANEXO.search(texto or ""):
            for pp in range(i + 1, i + 2 + PAGINAS_POSTERIORES_ANEXO):
                if 1 <= pp <= total_paginas:
                    incluidas.add(pp)

    paginas_ordenadas = sorted(incluidas)
    texto = "\n\n".join(f"--- Página {p} ---\n{paginas_texto[p - 1]}" for p in paginas_ordenadas)
    if len(texto) > max_chars:
        texto = texto[:max_chars]
    return texto, paginas_ordenadas


# Patrones para localizar el Anexo I / Anexo III dentro del texto del PCAP — es donde suelen
# vivir los requisitos concretos de solvencia económica/técnica cuando el cuerpo de cláusulas
# se limita a remitir a ellos. \b evita que "anexo i" case dentro de "anexo iii".
PATRONES_ANEXO = [r'anexo\s+i\b', r'anexo\s+1\b', r'anexo\s+iii\b', r'anexo\s+3\b']


# Solo cuenta como Anexo I/III la página cuyo texto tiene una línea que EMPIEZA por "ANEXO I/III":
# contar cualquier mención en el cuerpo ("véase Anexo I") metía en algunos PCAP más de la mitad
# de las páginas. Un anexo ocupa varias páginas y las siguientes no repiten el encabezado, por
# eso construir_texto_reducido incluye también las 3 páginas posteriores a cada encabezado.
# El Cuadro Resumen / de Características (datos clave: presupuesto, plazos, solvencia, criterios) se
# trata igual: varios pliegos lo llevan como anexo y el localizador se lo dejaba fuera.
PATRON_CABECERA_ANEXO = re.compile(
    r'^\s*(?:anexo\s+(?:i|1|iii|3)\b|cuadro[\s-]+(?:resumen|de\s+caracter[ií]sticas))',
    re.IGNORECASE | re.MULTILINE,
)
PAGINAS_POSTERIORES_ANEXO = 3


def recortar_preservando_anexos(texto, max_chars):
    """Si el texto cabe entero, no se toca. Si hay que recortar, en vez de cortar a ciegas
    por los primeros caracteres (arriesgándose a perder el Anexo I/III, que suele estar pasada
    la página 50), se localiza dónde empieza y se preserva ese bloque completo, recortando de
    en medio si hace falta."""
    if len(texto) <= max_chars:
        return texto

    texto_lower = texto.lower()
    posiciones = [m.start() for patron in PATRONES_ANEXO for m in re.finditer(patron, texto_lower)]
    if not posiciones:
        return texto[:max_chars]

    inicio_anexo = min(posiciones)
    presupuesto_inicio = max_chars // 3
    inicio_doc = texto[:presupuesto_inicio]
    separador = "\n[...]\n"
    espacio_restante = max(max_chars - len(inicio_doc) - len(separador), 0)
    return inicio_doc + separador + texto[inicio_anexo:inicio_anexo + espacio_restante]


def obtener_pcap(url, api_key=None, min_chars_por_pagina=100, max_chars_texto=350000, max_bytes_documento=32 * 1024 * 1024):
    """Descarga el PCAP y decide cómo se mandará a la IA:
    - Si tiene una capa de texto razonable (documento generado digitalmente, el caso normal), se
      extrae el texto con pypdf — mucho más barato en tokens que mandar el PDF entero (Claude
      convierte cada página de un "documento" en una imagen, ~1.500-3.000+ tokens por página).
      Con la clave de la API disponible, primero se localizan (con Haiku, barato) las páginas
      relevantes para cada apartado del informe y solo esas se mandan al modelo caro — un PCAP de
      50+ páginas puede quedar reducido a un puñado de páginas reales. Si la localización falla
      por lo que sea, se cae al recorte por Anexo I/III de siempre (350.000 caracteres de tope).
    - Si el texto extraído es casi nulo (indicio de páginas escaneadas/rasterizadas), se manda
      el PDF completo en base64 como documento, para que Claude lo lea visualmente — esto sí es
      caro, pero solo ocurre cuando de verdad hace falta OCR.
    Devuelve (tipo, contenido) con tipo en {"texto", "documento", None}.
    """
    if not url:
        return None, None
    resp = requests.get(url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    contenido_bytes = resp.content

    try:
        lector = PdfReader(BytesIO(contenido_bytes))
        paginas_texto = [pagina.extract_text() or "" for pagina in lector.pages]
        num_paginas = len(paginas_texto)
        texto_completo = "\n".join(paginas_texto)
    except Exception:
        num_paginas = 1
        paginas_texto = []
        texto_completo = ""

    promedio_por_pagina = len(texto_completo) / max(num_paginas, 1)
    if promedio_por_pagina >= min_chars_por_pagina:
        if api_key and num_paginas > 1:
            try:
                secciones = localizar_paginas_relevantes(paginas_texto, api_key)
                texto_reducido, paginas_incluidas = construir_texto_reducido(
                    paginas_texto, secciones, max_chars=max_chars_texto
                )
                if num_paginas > 6 and (not any(secciones.values()) or len(paginas_incluidas) <= 3):
                    # El localizador no ha encontrado nada útil (o solo quedan las 3 páginas fijas
                    # del inicio). Mejor el documento completo (recortado) que un informe sin cláusulas.
                    raise ValueError(f"el localizador no encontró páginas útiles ({len(paginas_incluidas)}/{num_paginas})")
                print(
                    f"    Localizador: usando {len(paginas_incluidas)}/{num_paginas} páginas "
                    f"({len(texto_reducido)} caracteres) en vez del documento completo."
                )
                return "texto", texto_reducido
            except Exception as e:
                print(f"    Aviso: falló la localización de secciones ({e}); se usa el recorte por Anexo I/III de siempre.")
        if len(texto_completo) > max_chars_texto:
            print(f"    Aviso: PCAP de {len(texto_completo)} caracteres supera el tope de {max_chars_texto}; se recorta preservando el Anexo I/III si se localiza.")
        return "texto", recortar_preservando_anexos(texto_completo, max_chars_texto)

    if len(contenido_bytes) > max_bytes_documento:
        print(f"    Aviso: PCAP de {len(contenido_bytes)} bytes supera el máximo de {max_bytes_documento}; se omite.")
        return None, None
    return "documento", base64.b64encode(contenido_bytes).decode("ascii")


# Precios de ANTHROPIC_MODEL (Sonnet 5.5) en USD por millón de tokens, para estimar lo que cuesta
# cada informe. Si se cambia de modelo o de precios, actualizarlos aquí.
PRECIO_USD_M = {"entrada": 2.0, "salida": 10.0, "cache_leida": 0.20, "cache_escrita": 2.5}


def estimar_coste_usd(uso):
    total = (
        (uso.get("input_tokens") or 0) * PRECIO_USD_M["entrada"]
        + (uso.get("output_tokens") or 0) * PRECIO_USD_M["salida"]
        + (uso.get("cache_read_input_tokens") or 0) * PRECIO_USD_M["cache_leida"]
        + (uso.get("cache_creation_input_tokens") or 0) * PRECIO_USD_M["cache_escrita"]
    ) / 1_000_000
    return round(total, 4)


# El informe termina con una línea "SEMAFORO: color | motivo" (ver PROMPT_SISTEMA_INFORME): se separa
# del texto y se guarda aparte para que el visor pinte la celda del expediente. Tolera negritas,
# viñetas o comillas que el modelo añada por su cuenta.
RE_SEMAFORO = re.compile(
    r"^[ \t>*_`#\-]*SEMAFORO\s*:\s*[*_`]*\s*(verde|naranja|rojo|gris)\b[ \t*_`]*\|?[ \t]*(.*?)[ \t*_`\"]*$",
    re.IGNORECASE | re.MULTILINE,
)


def extraer_semaforo(texto):
    """Devuelve (texto sin la línea del semáforo, color o None, motivo o None). Si el modelo repite
    la línea, vale la última."""
    coincidencias = list(RE_SEMAFORO.finditer(texto))
    if not coincidencias:
        return texto, None, None
    ultima = coincidencias[-1]
    color = ultima.group(1).lower()
    motivo = ultima.group(2).strip().strip('"').strip() or None
    limpio = RE_SEMAFORO.sub("", texto).rstrip()
    return limpio, color, motivo


# Entre el informe y la línea del semáforo la IA añade un bloque JSON con los datos para el Excel de la
# carpeta (ver PROMPT_SISTEMA_INFORME). Se separa del texto del informe y se guarda en datos_excel.
RE_DATOS_EXCEL = re.compile(r"<<<\s*DATOS_EXCEL\s*(.*?)\s*DATOS_EXCEL\s*>>>", re.DOTALL)


def extraer_datos_excel(texto):
    """Devuelve (texto sin el bloque, datos o None). Si el JSON no es válido (o el bloque quedó cortado)
    el informe sigue siendo válido: simplemente no habrá Excel para esa licitación."""
    coincidencias = list(RE_DATOS_EXCEL.finditer(texto))
    if not coincidencias:
        return RE_DATOS_EXCEL_ABIERTO.sub("", texto).rstrip(), None
    crudo = coincidencias[-1].group(1).strip()
    crudo = re.sub(r"^```(?:json)?\s*|\s*```$", "", crudo).strip()
    limpio = RE_DATOS_EXCEL.sub("", texto).rstrip()
    try:
        datos = json.loads(crudo)
    except ValueError as e:
        print(f"    Aviso: el bloque de datos del Excel no es JSON válido ({e}); no se guarda.")
        return limpio, None
    if not isinstance(datos, dict) or not any(k in datos for k in ("info", "solvencia", "juicio_valor", "formulas", "otros")):
        return limpio, None
    return limpio, datos


# Bloque abierto y nunca cerrado (respuesta cortada): se quita igualmente para no ensuciar el informe.
RE_DATOS_EXCEL_ABIERTO = re.compile(r"<<<\s*DATOS_EXCEL.*", re.DOTALL)


def obtener_ppt(url, max_chars=90000, min_chars=500):
    """Texto del PPT (pliego de prescripciones técnicas) para mandarlo a la IA junto al PCAP. Es un
    documento corto, así que va entero (con tope de seguridad). Devuelve None si no hay enlace, si no
    se puede descargar o si es un escaneo sin texto: el informe se genera igualmente, sin PPT."""
    if not url:
        return None
    try:
        resp = requests.get(url, headers=HEADERS, timeout=60)
        resp.raise_for_status()
        paginas = [(p.extract_text() or "") for p in PdfReader(BytesIO(resp.content)).pages]
    except Exception as e:
        print(f"    Aviso: no se pudo leer el PPT ({e}); se analiza solo el PCAP.")
        return None
    texto = "\n\n".join(f"--- Página {i + 1} ---\n{t}" for i, t in enumerate(paginas))
    if len(texto.strip()) < min_chars:
        print("    Aviso: el PPT no tiene texto legible (¿escaneado?); se analiza solo el PCAP.")
        return None
    return texto[:max_chars]


def generar_informe_licitacion_ia(perfil_empresa, pcap_tipo, pcap_contenido, titulo, organo, api_key, ppt_texto=None):
    if not api_key:
        return {"informe": "No hay ANTHROPIC_API_KEY configurada para llamar a Claude."}
    if not pcap_tipo:
        return {"informe": "No se pudo descargar el PCAP."}

    fecha_hoy = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    # El perfil de empresa es idéntico en todas las licitaciones de una misma ejecución: se manda
    # como bloque cacheable aparte (cache_control) para que, a partir de la segunda licitación
    # analizada en el mismo día, ese bloque (y el prompt del sistema) salgan con el descuento de
    # caché de Anthropic en vez de pagarse entero en cada llamada.
    bloque_perfil = {
        "type": "text",
        "text": f"PERFIL DE LA EMPRESA (CAI Consultores):\n{perfil_empresa}",
        "cache_control": {"type": "ephemeral"},
    }
    texto_variable = (
        f"FECHA DE HOY: {fecha_hoy}\n"
        f"LICITACIÓN: {titulo}\n"
        f"ÓRGANO: {organo}\n\n"
        f"A continuación se adjunta el PCAP (o el fragmento relevante ya localizado) de esta "
        f"licitación" + (" y el texto de su PPT" if ppt_texto else " (sin PPT)") +
        f". Genera el informe según las instrucciones del sistema."
    )

    if pcap_tipo == "documento":
        bloque_pcap = {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": pcap_contenido},
        }
    else:
        bloque_pcap = {"type": "text", "text": f"TEXTO DEL PCAP:\n{pcap_contenido}"}

    contenido_mensaje = [bloque_perfil, {"type": "text", "text": texto_variable}, bloque_pcap]
    if ppt_texto:
        contenido_mensaje.append({"type": "text", "text": f"TEXTO DEL PPT:\n{ppt_texto}"})

    # Un informe de 6 bloques con tablas ronda los 6.000-9.000 tokens de salida. Con el tope
    # anterior (4.096) se cortaba a mitad de frase. Además, el razonamiento de Sonnet cuenta
    # contra max_tokens, así que con poco margen podía comerse casi todo el cupo (incluso dejar la
    # respuesta sin texto). Por eso: tope amplio, razonamiento adaptativo y esfuerzo medio para
    # acotar cuánto razona sin renunciar a que cruce cláusulas y revise cifras.
    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 16000,
        "thinking": {"type": "adaptive"},
        "output_config": {"effort": "medium"},
        "system": [{
            "type": "text", "text": PROMPT_SISTEMA_INFORME,
            "cache_control": {"type": "ephemeral"},
        }],
        "messages": [{"role": "user", "content": contenido_mensaje}],
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
        "Content-Type": "application/json",
    }

    # Diagnóstico de tamaños (nunca se imprime el contenido del perfil, es un secret).
    print(
        f"    Tamaños enviados a la IA -> perfil: {len(perfil_empresa)} car., "
        f"PCAP: modo '{pcap_tipo}', {len(pcap_contenido)} caracteres"
    )

    try:
        resp = requests.post(ANTHROPIC_API_ENDPOINT, headers=headers, json=payload, timeout=600)
        if not resp.ok:
            print(f"    Respuesta de error de Anthropic (HTTP {resp.status_code}): {resp.text[:2000]}")
        resp.raise_for_status()
        data = resp.json()
        stop_reason = data.get("stop_reason")
        uso = data.get("usage", {})
        print(
            f"    Tokens usados -> entrada: {uso.get('input_tokens')}, salida: {uso.get('output_tokens')}, "
            f"caché escrita: {uso.get('cache_creation_input_tokens')}, caché leída: {uso.get('cache_read_input_tokens')}, "
            f"fin: {stop_reason}"
        )
        texto = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text").strip()
        if not texto:
            return {"informe": f"Error al generar el informe con IA: la respuesta no contiene texto (stop_reason={stop_reason})."}
        texto, datos_excel = extraer_datos_excel(texto)
        texto, semaforo, motivo = extraer_semaforo(texto)
        if stop_reason == "max_tokens":
            texto += "\n\n" + MARCA_INFORME_CORTADO
        resultado = {"informe": texto, "usa_ppt": bool(ppt_texto)}
        # Consumo de la llamada principal (sin el localizador, que cuesta céntimos): permite al visor
        # mostrar cuánto se gasta en revisiones, por persona y por mes.
        resultado["uso"] = {
            "entrada": uso.get("input_tokens"),
            "salida": uso.get("output_tokens"),
            "cache_leida": uso.get("cache_read_input_tokens"),
            "cache_escrita": uso.get("cache_creation_input_tokens"),
        }
        resultado["coste_estimado_usd"] = estimar_coste_usd(uso)
        if semaforo:
            resultado["semaforo"] = semaforo
            resultado["semaforo_motivo"] = motivo
        if datos_excel:
            resultado["datos_excel"] = datos_excel
        return resultado
    except Exception as e:
        return {"informe": f"Error al generar el informe con IA: {e!r}"}


def analizar_licitacion_ia(r, perfil_empresa):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    try:
        pcap_tipo, pcap_contenido = obtener_pcap(r.get("pcap_url"), api_key=api_key)
    except Exception:
        pcap_tipo, pcap_contenido = None, None

    ppt_texto = obtener_ppt(r.get("ppt_url")) if pcap_tipo else None
    resultado = generar_informe_licitacion_ia(
        perfil_empresa, pcap_tipo, pcap_contenido, r.get("descripcion") or r.get("titulo"), r.get("organo"), api_key, ppt_texto
    )
    resultado["fecha_analisis"] = datetime.now(timezone.utc).isoformat()
    return resultado


def cargar_manifiesto_filtros():
    if os.path.exists(FILTROS_MANIFEST_FILE):
        with open(FILTROS_MANIFEST_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("filtros", [])
    # Primera vez que se ejecuta con soporte multi-filtro: se crea el manifiesto
    # con un único filtro "1" en filtros/1/.
    manifiesto = {"filtros": [{"id": "1", "carpeta": f"{FILTROS_CARPETA_BASE}/1"}]}
    with open(FILTROS_MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifiesto, f, ensure_ascii=False, indent=2)
    return manifiesto["filtros"]


def ruta_filtro(desc, nombre_archivo):
    carpeta = desc["carpeta"]
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, nombre_archivo)


def cargar_estado(archivo):
    if os.path.exists(archivo):
        with open(archivo, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"ids_vistos": []}


def guardar_estado(archivo, estado):
    with open(archivo, "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def cargar_filtro_config(archivo, config_por_defecto):
    if os.path.exists(archivo):
        with open(archivo, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return {**config_por_defecto, **cfg}
    with open(archivo, "w", encoding="utf-8") as f:
        json.dump(config_por_defecto, f, ensure_ascii=False, indent=2)
    return dict(config_por_defecto)


def calcular_resultado(codigos):
    """Resume los ResultCode (uno por lote, lista oficial TenderResultCode) en una etiqueta:
    formalizada / adjudicada / desierta / desistida, o None si aún no hay resultado."""
    codigos = [c for c in codigos if c]
    if not codigos:
        return None
    if "9" in codigos:
        return "formalizada"
    if any(c in ("1", "2", "8", "10") for c in codigos):
        return "adjudicada"
    if any(c in ("3", "6", "7") for c in codigos):
        return "desierta"
    if any(c in ("4", "5") for c in codigos):
        return "desistida"
    return None


def aplicar_cambio_estado(registro, data, hoy):
    """Si el feed trae una versión más reciente de una licitación ya guardada, actualiza su
    estado/resultado y su fecha límite (las ampliaciones de plazo cambian ambos). No toca
    'updated' (la fecha de la licitación que muestra el visor) sino 'estado_updated'.
    Devuelve la lista de campos que cambiaron: 'estado' y/o 'fecha_limite'."""
    version_nueva = data.get("updated")
    version_actual = registro.get("estado_updated") or registro.get("updated")
    try:
        if version_nueva and version_actual and datetime.fromisoformat(version_nueva) <= datetime.fromisoformat(version_actual):
            return []  # versión igual o más antigua: no puede revertir lo ya actualizado
    except ValueError:
        pass

    cambios = []
    nuevo_estado = data.get("estado")
    if nuevo_estado and (nuevo_estado != registro.get("estado") or data.get("resultado") != registro.get("resultado")):
        registro["estado_anterior"] = registro.get("estado")
        registro["estado"] = nuevo_estado
        registro["resultado"] = data.get("resultado")
        registro["estado_cambiado_el"] = hoy
        cambios.append("estado")

    nuevo_limite = data.get("fecha_limite")
    if nuevo_limite and nuevo_limite != registro.get("fecha_limite"):
        registro["fecha_limite_anterior"] = registro.get("fecha_limite")
        registro["fecha_limite"] = nuevo_limite
        registro["fecha_limite_cambiada_el"] = hoy
        cambios.append("fecha_limite")

    if cambios:
        registro["estado_updated"] = version_nueva
    return cambios


def cargar_historico(archivo):
    if not os.path.exists(archivo):
        return []
    with open(archivo, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as e:
            raise RuntimeError(
                f"{archivo} tiene un error de sintaxis JSON ({e}). "
                "No se puede continuar sin arreglarlo a mano (revisa comas sobrantes o "
                "corchetes/llaves sin cerrar) para no arriesgarse a perder el histórico."
            ) from e


def informe_ia_valido(revision):
    """True si es un informe real (no un mensaje de error) que merece reutilizarse."""
    if not isinstance(revision, dict):
        return False
    informe = revision.get("informe")
    return bool(informe) and MARCA_INFORME_CORTADO not in informe and not informe.startswith(
        ("Error al generar", "No hay ANTHROPIC", "No se pudo descargar")
    )


def normalizar_importe(txt):
    """'1.732,24' -> '1732.24'; los importes con punto decimal ('119296.8') se dejan igual."""
    t = (txt or "").strip().rstrip(".,")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    return t or None


# Las licitaciones de la Junta de Andalucía llegan por el feed de plataformas agregadas SIN órgano
# claro, importe ni enlaces a los pliegos. Su portal tiene un servicio de consulta (el que usa su
# propia web) con el detalle completo de cada expediente, incluidas las URL de descarga.
JUNTA_DETALLE_URL = "https://www.juntadeandalucia.es/haciendayadministracionpublica/apl/pdc-front-publico/elastic/sirec_pdc_expedientes_details/_search"
RE_ID_JUNTA = re.compile(r"juntadeandalucia\.es.*idExpediente=(\d+)")
MAX_CONSULTAS_JUNTA_RELLENO = 80


def id_expediente_junta(registro):
    m = RE_ID_JUNTA.search(registro.get("link") or "")
    return m.group(1) if m else None


def consultar_detalle_junta(id_expediente):
    r = requests.post(
        JUNTA_DETALLE_URL,
        json={"query": {"match": {"_id": id_expediente}}},
        headers={**HEADERS, "Content-Type": "application/json", "Accept": "application/json"},
        timeout=30,
    )
    r.raise_for_status()
    hits = r.json().get("hits", {}).get("hits", [])
    return hits[0]["_source"] if hits else None


def documento_junta(documentos, prefijo):
    """URL de descarga del documento activo más reciente cuya categoría empieza por PCAP / PPT."""
    activos = [d for d in documentos
               if (d.get("categoria") or "").upper().startswith(prefijo) and d.get("estado") == "Activo" and d.get("descarga")]
    if not activos:
        return None
    return max(activos, key=lambda d: d.get("fechaPublicacion") or "")["descarga"]


def enriquecer_registro_junta(registro, cache, hoy):
    """Completa órgano, importe y URL de PCAP/PPT de una licitación de la Junta (solo huecos).
    cache: {id_expediente: detalle o None} compartido entre filtros para consultar cada
    expediente una sola vez por ejecución. Devuelve True si el registro cambió."""
    id_exp = id_expediente_junta(registro)
    if not id_exp:
        return False
    if id_exp not in cache:
        try:
            cache[id_exp] = consultar_detalle_junta(id_exp)
        except Exception as e:
            print(f"    Aviso: no se pudo consultar el detalle de la Junta para {registro.get('folder_id')} ({e})")
            cache[id_exp] = None
        time.sleep(0.3)
    detalle = cache[id_exp]
    if not detalle:
        return False

    cambios = False

    def poner(campo, valor):
        nonlocal cambios
        if valor and not registro.get(campo):
            registro[campo] = valor
            cambios = True

    poner("organo", ((detalle.get("perfilContratante") or {}).get("descripcion") or "").strip())
    # En la Junta la descripción del expediente explica el objeto mucho mejor que su título
    # (que suele ser una abreviatura). Se guarda aunque venga vacía para no volver a consultarla.
    if registro.get("descripcion") is None:
        registro["descripcion"] = (detalle.get("descripcion") or "").strip()
        cambios = True
    if detalle.get("importeLicitacion") is not None:
        poner("importe", str(detalle["importeLicitacion"]))
    documentos = detalle.get("documentos") or []
    poner("pcap_url", documento_junta(documentos, "PCAP"))
    poner("ppt_url", documento_junta(documentos, "PPT"))
    if registro.get("junta_consultado") != hoy:
        registro["junta_consultado"] = hoy
        cambios = True
    return cambios


def parse_entry(entry, filtro_cfg):
    def find_text(path):
        el = entry.find(path, NS)
        return el.text if el is not None else None

    folder_id = find_text('.//cbc:ContractFolderID')
    estado_code = find_text('.//cbc-place-ext:ContractFolderStatusCode')
    titulo = find_text('atom:title')
    updated = find_text('atom:updated')
    summary_text = find_text('atom:summary')

    link_el = entry.find('atom:link', NS)
    link = link_el.get('href') if link_el is not None else None

    nuts_prefix = filtro_cfg.get("nuts_prefix") or ""
    nuts_codes = [e.text for e in entry.findall('.//cbc:CountrySubentityCode', NS)]
    es_andalucia = (not nuts_prefix) or any(c and c.startswith(nuts_prefix) for c in nuts_codes)

    cpv_permitidos = filtro_cfg.get("cpv_permitidos") or []
    cpv_codes = [e.text for e in entry.findall('.//cbc:ItemClassificationCode', NS)]
    cpv_match = any(c in cpv_permitidos for c in cpv_codes)

    pcap_el = entry.find('.//cac:LegalDocumentReference//cbc:URI', NS)
    pcap_url = pcap_el.text if pcap_el is not None else None

    ppt_el = entry.find('.//cac:TechnicalDocumentReference//cbc:URI', NS)
    ppt_url = ppt_el.text if ppt_el is not None else None
    
    deadline_el = entry.find('.//cac:TenderingProcess/cac:TenderSubmissionDeadlinePeriod/cbc:EndDate', NS)
    fecha_limite = deadline_el.text if deadline_el is not None else None

    organo, importe = None, None
    if summary_text:
        organo_match = re.search(r'rgano de [Cc]ontrataci.n:\s*(.*?)\s*;\s*Importe', summary_text)
        importe_match = re.search(r'Importe:\s*([\d.,]+)\s*(?:EUR|€)', summary_text)
        organo = organo_match.group(1).strip() if organo_match else None
        importe = normalizar_importe(importe_match.group(1)) if importe_match else None

    if importe is None:
        base_el = entry.find('.//cac:ProcurementProject/cac:BudgetAmount/cbc:TaxExclusiveAmount', NS)
        if base_el is not None and base_el.text:
            importe = base_el.text.strip()

    tipo_contrato = find_text('.//cac:ProcurementProject/cbc:TypeCode')
    procedimiento = find_text('.//cac:TenderingProcess/cbc:ProcedureCode')
    resultado = calcular_resultado([e.text for e in entry.findall('.//cac:TenderResult/cbc:ResultCode', NS)])

    importe_num = None
    importe_el = entry.find('.//cac:ProcurementProject/cac:BudgetAmount/cbc:EstimatedOverallContractAmount', NS)
    if importe_el is not None and importe_el.text:
        try:
            importe_num = float(importe_el.text)
        except ValueError:
            importe_num = None

    tipos_contrato_permitidos = filtro_cfg.get("tipos_contrato_permitidos") or []
    tipo_contrato_match = not tipos_contrato_permitidos or tipo_contrato in tipos_contrato_permitidos

    procedimientos_permitidos = filtro_cfg.get("procedimientos_permitidos") or []
    procedimiento_match = not procedimientos_permitidos or procedimiento in procedimientos_permitidos

    importe_min = filtro_cfg.get("importe_min")
    importe_max = filtro_cfg.get("importe_max")
    importe_match = True
    if importe_num is not None:
        if importe_min is not None and importe_num < importe_min:
            importe_match = False
        if importe_max is not None and importe_num > importe_max:
            importe_match = False

    return {
        "folder_id": folder_id,
        "estado": estado_code,
        "resultado": resultado,
        "titulo": titulo,
        "updated": updated,
        "link": link,
        "es_andalucia": es_andalucia,
        "cpv_match": cpv_match,
        "cpv_codes": cpv_codes,
        "pcap_url": pcap_url,
        "ppt_url": ppt_url,
        "organo": organo,
        "importe": importe,
        "importe_num": importe_num,
        "fecha_limite": fecha_limite,
        "tipo_contrato": tipo_contrato,
        "tipo_contrato_match": tipo_contrato_match,
        "procedimiento": procedimiento,
        "procedimiento_match": procedimiento_match,
        "importe_match": importe_match,
    }


import time

def fetch_pagina(url, intentos=3):
    for intento in range(1, intentos + 1):
        resp = None
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            return etree.fromstring(resp.content)
        except (requests.exceptions.RequestException, etree.XMLSyntaxError) as e:
            print(f"Intento {intento}/{intentos} fallido para {url}: {e}")
            if resp is not None:
                print(f"    HTTP {resp.status_code} | primeros 300 caracteres de la respuesta: {resp.text[:300]!r}")
            if intento == intentos:
                raise
            time.sleep(5 * intento)  # espera creciente: 5s, 10s, 15s...


def get_next_link(root):
    next_el = root.find('atom:link[@rel="next"]', NS)
    return next_el.get('href') if next_el is not None else None
    
def notificar_teams(resultados, paginas, total_entries, nombre_filtro="Monitor"):
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not webhook_url:
        print("Aviso: no se ha configurado TEAMS_WEBHOOK_URL, se omite notificación.")
        return

    ahora = datetime.now(timezone.utc).strftime("%H:%M UTC")

    texto_resumen = (
        f"**Lectura PLACSP — {nombre_filtro}** ({ahora})\n\n"
        f"Páginas leídas: {paginas}  \n"
        f"Entries totales leídas: {total_entries}  \n"
        f"Licitaciones nuevas filtradas: {len(resultados)}\n\n"
        f"[Abrir visor completo]({VISOR_URL})"
    )

    if resultados:
        detalle = "\n\n".join(
            f"- **[{r['folder_id']}]** {r['titulo']}  \n{r['link']}"
            for r in resultados[:15]
        )
        texto_resumen += f"\n\n{detalle}"
        if len(resultados) > 15:
            texto_resumen += f"\n\n_(y {len(resultados) - 15} más)_"

    adaptive_card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock",
                "text": texto_resumen,
                "wrap": True
            }
        ],
        "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "Abrir visor",
                "url": VISOR_URL
            }
        ]
    }

    payload = {
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": adaptive_card
            }
        ]
    }

    try:
        resp = requests.post(webhook_url, json=payload, timeout=15)
        resp.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error notificando a Teams: {e}")


def leer_feed(nombre, url_inicial, contextos, limite_fecha, fecha_captura_hoy):
    """Recorre un feed (de más reciente a más antiguo) hasta salir de la ventana de tiempo y
    evalúa cada entrada contra todos los filtros. Devuelve (páginas leídas, entradas leídas)."""
    url_actual = url_inicial
    pagina = 0
    entradas_leidas = 0
    versiones_vistas = set()
    ultima_entrada = None  # fecha de la entrada más reciente del feed: sirve para detectar feeds congelados

    while url_actual and pagina < MAX_PAGINAS:
        pagina += 1
        root = fetch_pagina(url_actual)
        entries = root.findall('atom:entry', NS)

        if not entries:
            break

        parar = False
        for entry in entries:
            entradas_leidas += 1
            updated_el = entry.find('atom:updated', NS)
            if updated_el is None or not updated_el.text:
                continue
            fecha_entry = datetime.fromisoformat(updated_el.text)
            if ultima_entrada is None or fecha_entry > ultima_entrada:
                ultima_entrada = fecha_entry
            if fecha_entry < limite_fecha:
                parar = True
                break

            # El feed va de más reciente a más antiguo: si una licitación aparece varias veces,
            # la primera es su versión actual y las demás son estados ya superados.
            folder_id_entry = entry.findtext('.//cbc:ContractFolderID', namespaces=NS)
            if folder_id_entry:
                if folder_id_entry in versiones_vistas:
                    continue
                versiones_vistas.add(folder_id_entry)

            for ctx in contextos:
                cfg = ctx["cfg"]
                data = parse_entry(entry, cfg)

                # Licitación ya guardada que reaparece en el feed: si ha cambiado de estado
                # (adjudicada, desierta, anulada...) o de fecha límite (ampliación de plazo) se
                # actualiza su registro, aunque el nuevo estado ya no entre en los estados
                # permitidos del filtro.
                existente = ctx["historico_idx"].get(data["folder_id"])
                if existente is not None:
                    estado_previo = existente.get("estado")
                    limite_previo = existente.get("fecha_limite")
                    cambios = aplicar_cambio_estado(existente, data, fecha_captura_hoy)
                    if cambios:
                        ctx["cambios_estado"].append(data["folder_id"])
                        detalle = []
                        if "estado" in cambios:
                            detalle.append(f"estado {estado_previo} -> {data['estado']}"
                                           f"{' (' + data['resultado'] + ')' if data['resultado'] else ''}")
                        if "fecha_limite" in cambios:
                            detalle.append(f"fecha límite {limite_previo} -> {data['fecha_limite']}")
                        print(f"    Cambio [{data['folder_id']}] ({nombre}): " + "; ".join(detalle))
                    continue

                cpv_permitidos = cfg.get("cpv_permitidos") or []
                estados_permitidos = cfg.get("estados_permitidos") or []

                if (data["es_andalucia"]
                    and (not cpv_permitidos or data["cpv_match"])
                    and (not estados_permitidos or data["estado"] in estados_permitidos)
                    and data["tipo_contrato_match"]
                    and data["procedimiento_match"]
                    and data["importe_match"]
                    and data["folder_id"] not in ctx["ids_vistos"]):
                    data["fuente"] = nombre  # feed del que sale (perfiles / agregadas / menores)
                    ctx["resultados"].append(data)
                    ctx["ids_vistos"].add(data["folder_id"])

        if parar:
            break

        url_actual = get_next_link(root)
        if url_actual:
            time.sleep(2)  # pausa entre páginas para no parecer scraping automático agresivo

    return pagina, entradas_leidas, (ultima_entrada.isoformat() if ultima_entrada else None)


# Marcas que pone el equipo desde el visor: "revisado", "presentado" y "descartado". Valen para la licitación en
# todos los filtros donde aparece; el visor las escribe en todos, y aquí se deja igual lo que haya
# quedado desigual (datos anteriores, un guardado a medias) y las licitaciones que entran nuevas
# en otro filtro heredan la marca. Si hay varias, manda la más reciente.
CAMPOS_MARCA_REVISION = (
    ("revisado_humano", "revisado_fecha", "revisado_por"),
    ("presentado", "presentado_fecha", "presentado_por"),
    ("carpeta_descargada", "carpeta_descargada_fecha", "carpeta_descargada_por"),
    ("descartado", "descartado_fecha", "descartado_por"),
)


def unificar_marcas_revision(contextos):
    ultima = {}  # (folder_id, campo_valor) -> registro con la marca más reciente
    for ctx in contextos:
        for r in ctx["historico"]:
            fid = r.get("folder_id")
            if not fid:
                continue
            for c_valor, c_fecha, _ in CAMPOS_MARCA_REVISION:
                if r.get(c_valor):
                    k = (fid, c_valor)
                    if k not in ultima or str(r.get(c_fecha) or "") > str(ultima[k].get(c_fecha) or ""):
                        ultima[k] = r
    propagadas = 0
    for ctx in contextos:
        for r in list(ctx["historico"]) + list(ctx["resultados"]):
            fid = r.get("folder_id")
            for c_valor, c_fecha, c_por in CAMPOS_MARCA_REVISION:
                origen = ultima.get((fid, c_valor))
                if origen is None or origen is r:
                    continue
                if any(r.get(c) != origen.get(c) for c in (c_valor, c_fecha, c_por)):
                    for c in (c_valor, c_fecha, c_por):
                        r[c] = origen.get(c)
                    propagadas += 1
    if propagadas:
        print(f"Marcas de revisión/descarte propagadas entre filtros: {propagadas}")


def main():
    filtros_desc = cargar_manifiesto_filtros()

    contextos = []
    for desc in filtros_desc:
        config_por_defecto = FILTRO_CONFIG_POR_DEFECTO if desc["id"] == "1" else FILTRO_CONFIG_VACIO
        cfg = cargar_filtro_config(ruta_filtro(desc, NOMBRE_CONFIG), config_por_defecto)
        estado = cargar_estado(ruta_filtro(desc, NOMBRE_ESTADO))
        # El histórico se carga ya (no al final) para poder actualizar el estado de
        # licitaciones ya guardadas cuando vuelvan a aparecer en el feed con otro estado.
        historico = cargar_historico(ruta_filtro(desc, NOMBRE_HISTORICO))
        contextos.append({
            "desc": desc,
            "cfg": cfg,
            "ids_vistos": set(estado.get("ids_vistos", [])),
            "total_entries_acumulado_previo": estado.get("total_entries_acumulado", 0),
            "resultados": [],
            "historico": historico,
            "historico_idx": {r["folder_id"]: r for r in historico if r.get("folder_id")},
            "cambios_estado": [],
        })

    fecha_hora_iso = datetime.now(timezone.utc).isoformat()
    fecha_captura_hoy = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    limite_fecha = datetime.now(timezone.utc) - timedelta(hours=VENTANA_HORAS)
    pagina = 0
    total_entries_leidas = 0
    feeds_con_error = []
    info_feeds = []  # por feed: páginas, entradas y fecha de su entrada más reciente (para el visor)

    for nombre_feed, url_feed, obligatorio in FEEDS:
        print(f"== Feed {nombre_feed}")
        try:
            paginas_feed, entradas_feed, ultima_entrada_feed = leer_feed(nombre_feed, url_feed, contextos, limite_fecha, fecha_captura_hoy)
        except Exception as e:
            if obligatorio:
                raise
            print(f"Aviso: no se pudo leer el feed '{nombre_feed}' ({e}); se continúa con el resto.")
            feeds_con_error.append(f"{nombre_feed}: {e}")
            info_feeds.append({"nombre": nombre_feed, "estado": "error", "error": str(e)[:200]})
            continue
        print(f"   páginas: {paginas_feed}, entradas: {entradas_feed}, última entrada: {ultima_entrada_feed}")
        info_feeds.append({"nombre": nombre_feed, "estado": "ok", "paginas": paginas_feed,
                           "entradas": entradas_feed, "ultima_entrada": ultima_entrada_feed})
        pagina += paginas_feed
        total_entries_leidas += entradas_feed

    print(f"Páginas leídas: {pagina}")
    print(f"Licitaciones leídas (total entries): {total_entries_leidas}")
    if feeds_con_error:
        notificar_fallo_teams("; ".join(feeds_con_error), parcial=True)

    # Completar las licitaciones de la Junta: las nuevas de hoy y, una sola vez, las ya guardadas
    # que se capturaron sin órgano, importe o pliegos.
    cache_junta = {}
    for ctx in contextos:
        for r in ctx["resultados"]:
            enriquecer_registro_junta(r, cache_junta, fecha_captura_hoy)
    relleno = 0
    for ctx in contextos:
        for r in ctx["historico"]:
            if not id_expediente_junta(r) or r.get("descripcion") is not None:
                continue
            if id_expediente_junta(r) not in cache_junta and relleno >= MAX_CONSULTAS_JUNTA_RELLENO:
                continue
            if id_expediente_junta(r) not in cache_junta:
                relleno += 1
            enriquecer_registro_junta(r, cache_junta, fecha_captura_hoy)
    if cache_junta:
        print(f"Detalle de la Junta consultado para {len(cache_junta)} expediente(s).")

    resultados_por_filtro = {}
    # La revisión IA ya no se genera sola: se solicita desde el visor (flujo "Revisión IA bajo
    # demanda"). Aquí solo se reutiliza, sin ninguna llamada a la API, un informe válido que la
    # misma licitación ya tenga en el histórico de cualquier otro filtro.
    informes_ia = {}
    for ctx in contextos:
        for registro in ctx["historico"]:
            if registro.get("folder_id") and informe_ia_valido(registro.get("revision_ia")):
                informes_ia.setdefault(registro["folder_id"], registro["revision_ia"])

    unificar_marcas_revision(contextos)

    for ctx in contextos:
        desc = ctx["desc"]
        resultados = ctx["resultados"]
        nombre_filtro = ctx["cfg"].get("nombre") or f"Filtro {desc['id']}"
        print(f"[{nombre_filtro}] licitaciones nuevas filtradas: {len(resultados)}"
              f" | cambios de estado: {len(ctx['cambios_estado'])}")

        for r in resultados:
            clave_informe = r.get("folder_id")
            if clave_informe and clave_informe in informes_ia:
                r["revision_ia"] = dict(informes_ia[clave_informe])
                print(f"    -> Informe IA reutilizado para [{r['folder_id']}] (ya existía en otro filtro)")

        total_entries_acumulado = ctx["total_entries_acumulado_previo"] + total_entries_leidas
        guardar_estado(ruta_filtro(desc, NOMBRE_ESTADO), {
            "ids_vistos": list(ctx["ids_vistos"]),
            "total_entries_acumulado": total_entries_acumulado,
        })

        metadata_lectura = {
            "fecha_hora": fecha_hora_iso,
            "paginas": pagina,
            "total_entries_leidas": total_entries_leidas,
            "nuevas_filtradas": len(resultados),
            "total_entries_acumulado": total_entries_acumulado,
            "feeds": info_feeds,
        }
        with open(ruta_filtro(desc, NOMBRE_ULTIMA_LECTURA), "w", encoding="utf-8") as f:
            json.dump(metadata_lectura, f, ensure_ascii=False, indent=2)

        for r in resultados:
            print(f"  - [{r['folder_id']}] {r['titulo']} | {r['link']}")

        with open(ruta_filtro(desc, NOMBRE_RESULTADO_HOY), "w", encoding="utf-8") as f:
            json.dump(resultados, f, ensure_ascii=False, indent=2)

        for r in resultados:
            r["fecha_captura"] = fecha_captura_hoy

        archivo_historico = ruta_filtro(desc, NOMBRE_HISTORICO)
        historico = ctx["historico"]
        historico.extend(resultados)
        with open(archivo_historico, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False, indent=2)

        notificar_teams(resultados, pagina, total_entries_leidas, nombre_filtro)
        resultados_por_filtro[desc["id"]] = resultados

    return resultados_por_filtro


def notificar_fallo_teams(mensaje_error, parcial=False):
    webhook_url = os.environ.get("TEAMS_WEBHOOK_URL")
    if not webhook_url:
        return

    ahora = datetime.now(timezone.utc).strftime("%H:%M UTC")
    texto = (
        f"**⚠ Lectura PLACSP fallida** ({ahora})\n\n"
        f"La ejecución de hoy no ha podido completarse:\n\n"
        f"`{mensaje_error}`\n\n"
        f"No se han detectado licitaciones nuevas en esta ejecución. "
        f"Se reintentará en la próxima ejecución programada."
    )
    if parcial:
        texto = (
            f"**⚠ Lectura PLACSP parcial** ({ahora})\n\n"
            f"No se pudo leer alguno de los feeds secundarios; el resto sí se procesó:\n\n"
            f"`{mensaje_error}`"
        )
    adaptive_card = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.4",
        "body": [{"type": "TextBlock", "text": texto, "wrap": True}],
    }
    payload = {
        "attachments": [
            {"contentType": "application/vnd.microsoft.card.adaptive", "content": adaptive_card}
        ]
    }
    try:
        requests.post(webhook_url, json=payload, timeout=15).raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error notificando el fallo a Teams: {e}")


def _historicos_con_licitacion(folder_id):
    """[(ruta, histórico completo, [registros de esa licitación])] de cada filtro que la contiene."""
    encontrados = []
    for desc in cargar_manifiesto_filtros():
        ruta = ruta_filtro(desc, NOMBRE_HISTORICO)
        historico = cargar_historico(ruta)
        registros = [r for r in historico if r.get("folder_id") == folder_id]
        if registros:
            encontrados.append((ruta, historico, registros))
    return encontrados


def generar_revision_bajo_demanda(folder_id, salida, forzar=False, solicitante=None):
    """Genera (o reutiliza, si ya existe uno válido) el informe IA de una licitación y lo deja en
    un archivo temporal. Separado de aplicar_revision_bajo_demanda para que, si el push falla
    porque otro proceso escribió antes, se pueda reaplicar sin volver a pagar la llamada a la IA."""
    encontrados = _historicos_con_licitacion(folder_id)
    if not encontrados:
        raise RuntimeError(f"No se encontró la licitación {folder_id!r} en ningún histórico.")
    registros = [r for _, _, regs in encontrados for r in regs]

    resultado = None
    if not forzar:
        previo = next((r["revision_ia"] for r in registros if informe_ia_valido(r.get("revision_ia"))), None)
        if previo:
            print("Ya existe un informe válido de esta licitación; se reutiliza sin llamar a la IA.")
            resultado = previo

    if resultado is None:
        perfil_empresa = os.environ.get("SOLVENCIA_EMPRESA")
        if not perfil_empresa:
            resultado = {
                "informe": "Error al generar el informe con IA: falta el secret SOLVENCIA_EMPRESA.",
                "fecha_analisis": datetime.now(timezone.utc).isoformat(),
            }
        else:
            base = next((r for r in registros if r.get("pcap_url")), registros[0])
            resultado = analizar_licitacion_ia(base, perfil_empresa)
        print(f"Informe de {folder_id}: {len(resultado['informe'])} caracteres")
        if solicitante:
            resultado["solicitada_por"] = solicitante[:60]

        # Al regenerar a la fuerza, un fallo (clave caducada, PCAP no descargable...) no debe
        # machacar un informe que ya era válido, aunque fuera incompleto.
        if forzar and not informe_ia_valido(resultado):
            previo = next((r["revision_ia"] for r in registros if informe_ia_valido(r.get("revision_ia"))), None)
            if previo:
                print("La regeneración ha fallado; se conserva el informe anterior.")
                resultado = previo

    with open(salida, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)


def aplicar_revision_bajo_demanda(folder_id, entrada):
    """Copia el informe generado a TODOS los registros de esa licitación, en todos los filtros."""
    with open(entrada, "r", encoding="utf-8") as f:
        resultado = json.load(f)
    total = 0
    for ruta, historico, registros in _historicos_con_licitacion(folder_id):
        for r in registros:
            r["revision_ia"] = dict(resultado)
            total += 1
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False, indent=2)
    print(f"Revisión IA aplicada a {total} registro(s) de {folder_id}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Monitor de licitaciones PLACSP")
    parser.add_argument("--generar-revision", metavar="FOLDER_ID", help="genera el informe IA de una licitación")
    parser.add_argument("--aplicar-revision", metavar="FOLDER_ID", help="guarda un informe ya generado en los históricos")
    parser.add_argument("--salida", help="archivo donde dejar el informe generado")
    parser.add_argument("--entrada", help="archivo con el informe a guardar")
    parser.add_argument("--forzar", action="store_true", help="regenera aunque ya exista un informe válido")
    parser.add_argument("--solicitante", default=None, help="nombre de quien pide la revisión (se guarda en el informe)")
    args = parser.parse_args()

    if args.generar_revision:
        generar_revision_bajo_demanda(args.generar_revision, args.salida, args.forzar, args.solicitante)
    elif args.aplicar_revision:
        aplicar_revision_bajo_demanda(args.aplicar_revision, args.entrada)
    else:
        try:
            main()
        except Exception as e:
            print(f"Fallo la ejecucion: {e}")
            notificar_fallo_teams(str(e))
            raise
