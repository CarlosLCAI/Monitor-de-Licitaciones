/* Capa única de acceso a los datos y a las operaciones del visor.

   Todo lo que el visor necesita del "servidor" pasa por el objeto Backend: leer datos, guardar marcas y
   criterios, lanzar la revisión IA o la preparación de carpetas, listar las carpetas ya preparadas y saber
   quién es el usuario. index.html no llama a GitHub directamente; si algún día el visor pasa a otro
   servidor (alojamiento propio con inicio de sesión de Microsoft 365, Supabase, etc.), basta con sustituir
   este archivo por otro que ofrezca las mismas funciones:

     Credenciales y usuario
       propietario                       usuario cuyas marcas se muestran con el check y no con la inicial
       tieneCredenciales()               true si esta persona puede operar (hoy: hay token guardado)
       nombreLocal()                     nombre que la persona ha escrito (hoy) o el de su cuenta
       usuarioActual()      -> Promise   nombre con el que se firman las marcas y las solicitudes
       cuentaRemota()       -> Promise   usuario de la cuenta en el servidor (hoy: login de GitHub)
       guardarCredenciales(clave, nombre)  solo en la versión con token

     Lectura
       leerJson(ruta)       -> Promise   JSON publicado (null si no existe)
       leerHistoricoFresco(ruta) -> Promise  histórico sin esperar a la publicación (null si falla)
       listarCarpetas()     -> Promise   Map nombreZip -> { actualizado (ms), url }
       historiaMonitor()    -> Promise   cifras de la pestaña Estadísticas

     Escritura y operaciones (lanzan Error con un mensaje legible si fallan)
       lanzarFlujo('revisionIA' | 'carpeta', entradas)
       guardarConfig(ruta, cfg)
       actualizarHistorico(ruta, folderId, campos, mensaje) -> Promise histórico actualizado, o null */
const Backend = (() => {
  const OWNER = 'CarlosLCAI';
  const REPO = 'Monitor-de-Licitaciones';
  const BRANCH = 'main';
  const API = `https://api.github.com/repos/${OWNER}/${REPO}`;
  const TOKEN_KEY = 'monitorLicitaciones.githubToken';
  const NOMBRE_KEY = 'monitorLicitaciones.nombreUsuario';
  const HISTORIA_KEY = 'monitorLicitaciones.historiaApi';
  const HISTORIA_TTL_MS = 12 * 3600 * 1000;
  const FLUJOS = { revisionIA: 'revision_ia.yml', carpeta: 'preparar_carpeta.yml' };
  let cuentaCache = null;

  // --- Credenciales y usuario -------------------------------------------------------------------
  function token() {
    try { return localStorage.getItem(TOKEN_KEY) || ''; } catch (e) { return ''; }
  }

  function tieneCredenciales() {
    return !!token();
  }

  function nombreLocal() {
    try { return localStorage.getItem(NOMBRE_KEY) || ''; } catch (e) { return ''; }
  }

  // Cada persona usa su propio token; GitHub dice a qué usuario pertenece.
  async function cuentaRemota() {
    const t = token();
    if (cuentaCache && cuentaCache.token === t) return cuentaCache.login;
    const resp = await fetch('https://api.github.com/user', {
      headers: { Authorization: `token ${t}`, Accept: 'application/vnd.github+json' },
      cache: 'no-store',
    });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const login = (await resp.json()).login;
    cuentaCache = { token: t, login };
    return login;
  }

  // Si varias personas comparten un token, GitHub las ve como el mismo usuario: el nombre que escribe cada
  // una en su navegador manda sobre el usuario de GitHub al firmar.
  async function usuarioActual() {
    return nombreLocal() || cuentaRemota();
  }

  function guardarCredenciales(clave, nombre) {
    try {
      if (clave) localStorage.setItem(TOKEN_KEY, clave);
      if (nombre) localStorage.setItem(NOMBRE_KEY, nombre); else localStorage.removeItem(NOMBRE_KEY);
    } catch (e) { /* sin almacenamiento */ }
  }

  function cabeceras(aceptar) {
    return { Authorization: `token ${token()}`, Accept: aceptar || 'application/vnd.github+json' };
  }

  // --- Codificación para la API de contenidos de GitHub -----------------------------------------
  function utf8ToBase64(str) {
    const bytes = new TextEncoder().encode(str);
    let binario = '';
    bytes.forEach(b => { binario += String.fromCharCode(b); });
    return btoa(binario);
  }

  function base64ToUtf8(b64) {
    const binario = atob(b64.replace(/\n/g, ''));
    const bytes = Uint8Array.from(binario, c => c.charCodeAt(0));
    return new TextDecoder('utf-8').decode(bytes);
  }

  // --- Lectura ----------------------------------------------------------------------------------
  // Archivo publicado junto al visor (filtros.json, filtros/<id>/historico.json, ...). null si no existe.
  async function leerJson(ruta) {
    try {
      const r = await fetch('./' + ruta, { cache: 'no-cache' });
      return r.ok ? await r.json() : null;
    } catch (e) {
      return null;
    }
  }

  // El histórico tal como está ya guardado, sin esperar a que se republique la web.
  async function leerHistoricoFresco(ruta) {
    const resp = await fetch(`${API}/contents/${ruta}?ref=${BRANCH}`, {
      headers: cabeceras('application/vnd.github.raw+json'),
      cache: 'no-store',
    });
    if (!resp.ok) return null;
    return JSON.parse(await resp.text());
  }

  // ZIP de carpetas ya preparadas: Map nombre del ZIP -> { actualizado (ms), url }. Vacío si aún no hay ninguna.
  async function listarCarpetas() {
    const h = { Accept: 'application/vnd.github+json' };
    if (token()) h.Authorization = 'token ' + token();
    const resp = await fetch(`${API}/releases/tags/carpetas`, { headers: h, cache: 'no-store' });
    if (!resp.ok) return new Map(); // 404: todavía no se ha preparado ninguna
    const release = await resp.json();
    return new Map((release.assets || []).map(a => [a.name, { actualizado: Date.parse(a.updated_at), url: a.browser_download_url }]));
  }

  // --- Operaciones ------------------------------------------------------------------------------
  async function lanzarFlujo(clave, entradas) {
    const flujo = FLUJOS[clave];
    const resp = await fetch(`${API}/actions/workflows/${flujo}/dispatches`, {
      method: 'POST',
      headers: cabeceras(),
      body: JSON.stringify({ ref: BRANCH, inputs: entradas || {} }),
    });
    if (resp.status === 204) return;
    const err = await resp.json().catch(() => ({}));
    let ayuda = '';
    if (resp.status === 403) ayuda = '\n\nEl token necesita el permiso "Actions: lectura y escritura" (token fine-grained) o el alcance "workflow" (token clásico).';
    if (resp.status === 404) ayuda = '\n\nNo se encuentra el flujo "' + flujo + '" en el repositorio (¿se ha subido?) o el token no tiene acceso a Actions.';
    if (resp.status === 422) ayuda = '\n\nEl flujo no admite lanzamiento manual (workflow_dispatch).';
    throw new Error(`HTTP ${resp.status}: ${err.message || 'sin detalle'}${ayuda}`);
  }

  // Criterios de un filtro (config.json).
  async function guardarConfig(ruta, nuevoCfg, intentosMax = 3) {
    if (!tieneCredenciales()) throw new Error('Configura primero un token de GitHub.');
    const apiUrl = `${API}/contents/${ruta}`;
    const h = cabeceras();
    const contenido = JSON.stringify(nuevoCfg, null, 2) + '\n';

    for (let intento = 1; intento <= intentosMax; intento++) {
      const actual = await fetch(`${apiUrl}?ref=${BRANCH}`, { headers: h, cache: 'no-store' });
      let sha; // sin valor = el archivo no existe todavía, se creará de cero
      if (actual.status !== 404) {
        if (!actual.ok) {
          throw new Error(`No se pudo leer ${ruta} en GitHub (HTTP ${actual.status}). Revisa el token y el nombre del repo.`);
        }
        sha = (await actual.json()).sha;
      }

      const cuerpoPut = { message: `Actualiza ${ruta} desde el visor`, content: utf8ToBase64(contenido), branch: BRANCH };
      if (sha) cuerpoPut.sha = sha;

      const resp = await fetch(apiUrl, { method: 'PUT', headers: h, body: JSON.stringify(cuerpoPut) });
      if (resp.ok) return resp.json();
      if (resp.status === 409 && intento < intentosMax) continue; // otro cambio llegó primero; reintenta con el sha más reciente

      const err = await resp.json().catch(() => ({}));
      throw new Error(`Error al guardar (HTTP ${resp.status}): ${err.message || 'sin detalle'}`);
    }
  }

  // Contenido y sha de un archivo del repositorio. La API de contenidos no devuelve el contenido de los
  // archivos de más de 1 MB (content vacío, pero con su sha): en ese caso se pide el blob exacto de ese sha,
  // que sirve hasta 100 MB y es coherente con el sha que luego se usa para guardar.
  async function leerArchivo(ruta) {
    const r = await fetch(`${API}/contents/${ruta}?ref=${BRANCH}`, { headers: cabeceras(), cache: 'no-store' });
    if (!r.ok) {
      throw new Error(`No se pudo leer ${ruta} en GitHub (HTTP ${r.status}). Revisa el token y el nombre del repo.`);
    }
    const d = await r.json();
    if (d.content) return { sha: d.sha, texto: base64ToUtf8(d.content) };
    const b = await fetch(`${API}/git/blobs/${d.sha}`, { headers: cabeceras('application/vnd.github.raw+json'), cache: 'no-store' });
    if (!b.ok) {
      throw new Error(`No se pudo leer ${ruta} en GitHub (archivo grande, HTTP ${b.status}). Revisa el token y el nombre del repo.`);
    }
    return { sha: d.sha, texto: await b.text() };
  }

  // Cambia campos de una licitación dentro del histórico de un filtro. Devuelve el histórico guardado, o
  // null si ese filtro ya no contiene la licitación.
  async function actualizarHistorico(ruta, folderId, campos, mensaje, intentosMax = 3) {
    const apiUrl = `${API}/contents/${ruta}`;
    const h = cabeceras();

    for (let intento = 1; intento <= intentosMax; intento++) {
      const actual = await leerArchivo(ruta);
      const historicoRemoto = JSON.parse(actual.texto);

      let tocados = 0;
      historicoRemoto.forEach(r => {
        if (r.folder_id === folderId) { Object.assign(r, campos); tocados++; }
      });
      if (!tocados) return null;

      const contenido = JSON.stringify(historicoRemoto, null, 2) + '\n';
      const resp = await fetch(apiUrl, {
        method: 'PUT',
        headers: h,
        body: JSON.stringify({ message: mensaje + ` (${ruta})`, content: utf8ToBase64(contenido), sha: actual.sha, branch: BRANCH }),
      });
      if (resp.ok) return historicoRemoto;
      if (resp.status === 409 && intento < intentosMax) continue; // otro cambio llegó primero; reintenta con la versión más reciente

      const err = await resp.json().catch(() => ({}));
      throw new Error(`Error al guardar (HTTP ${resp.status}): ${err.message || 'sin detalle'}`);
    }
    throw new Error('Hubo conflictos al guardar en GitHub; inténtalo de nuevo.');
  }

  // Cifras de la pestaña Estadísticas (versiones del script y del visor, ejecuciones, revisiones IA).
  async function historiaMonitor() {
    try {
      const guardada = JSON.parse(localStorage.getItem(HISTORIA_KEY));
      if (guardada && Date.now() - guardada.t < HISTORIA_TTL_MS) return guardada.d;
    } catch (e) { /* sin caché */ }

    const h = { Accept: 'application/vnd.github+json' };
    if (token()) h.Authorization = 'token ' + token();
    const pedir = async url => {
      const r = await fetch(url, { headers: h });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r;
    };
    // Nº de elementos = número de la última página con per_page=1; la primera subida es esa misma página.
    const contarCommits = async ruta => {
      const r = await pedir(API + '/commits?path=' + ruta + '&per_page=1');
      const m = (r.headers.get('Link') || '').match(/[?&]page=(\d+)>; rel="last"/);
      const n = m ? Number(m[1]) : (await r.json()).length;
      const primero = (await (await pedir(API + '/commits?path=' + ruta + '&per_page=1&page=' + n)).json())[0];
      return { n, primera: primero && primero.commit.committer.date };
    };
    const flujo = 'monitor_licitaciones.yml';
    const [script, visor, todas, exitosas, revisiones] = await Promise.all([
      contarCommits('fetch_licitaciones.py'),
      contarCommits('index.html'),
      pedir(API + '/actions/workflows/' + flujo + '/runs?per_page=1').then(r => r.json()),
      pedir(API + '/actions/workflows/' + flujo + '/runs?per_page=1&status=success').then(r => r.json()),
      pedir(API + '/actions/workflows/' + FLUJOS.revisionIA + '/runs?per_page=1&status=success').then(r => r.json()),
    ]);
    const primeraEjecucion = (await (await pedir(API + '/actions/workflows/' + flujo + '/runs?per_page=1&page=' + todas.total_count)).json()).workflow_runs[0];
    const d = {
      script, visor,
      ejecuciones: todas.total_count, lecturas: exitosas.total_count,
      primeraEjecucion: primeraEjecucion && primeraEjecucion.created_at,
      revisionesIA: revisiones.total_count,
    };
    try { localStorage.setItem(HISTORIA_KEY, JSON.stringify({ t: Date.now(), d })); } catch (e) { /* sin caché */ }
    return d;
  }

  return {
    propietario: OWNER,
    tieneCredenciales, nombreLocal, usuarioActual, cuentaRemota, guardarCredenciales,
    leerJson, leerHistoricoFresco, listarCarpetas, historiaMonitor,
    lanzarFlujo, guardarConfig, actualizarHistorico,
  };
})();
