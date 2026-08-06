/* =====================================================================
   APPS SCRIPT — Liga de Basquetbol
   -----------------------------------------------------------------------
   QUÉ HACE ESTE ARCHIVO (dos cosas independientes):

   1) actualizarDropdownEquipos()
      Cada vez que se inscribe un equipo nuevo (o cambia la lista en la
      hoja "Equipos"), actualiza automáticamente el menú desplegable
      "Equipo" de tu Form de Integrantes — así los jugadores eligen de
      una lista y nunca escriben el nombre del equipo a mano (cero typos).

   2) doPost(e)
      Recibe los datos que manda la página web "Captura" (marcador +
      asistencia de ambos rosters) y los escribe directo en las hojas
      "Resultados", "Rol de Juego" y "Asistencia", sin tocar las
      fórmulas que ya existen en el archivo.

   -----------------------------------------------------------------------
   CÓMO INSTALARLO (una sola vez, ~5 minutos):

   1. Abre tu Google Sheet (el que ya tiene las hojas Equipos, Resultados, etc.)
   2. Menú Extensiones > Apps Script.
   3. Borra el contenido de Code.gs y pega TODO este archivo.
   4. Arriba, donde dice "Seleccionar función", elige actualizarDropdownEquipos
      y dale clic a "Ejecutar". La primera vez te va a pedir autorización:
      acepta los permisos (son tuyos, sobre tu propio Sheet y tu propio Form).
   5. Activa el disparador automático: ícono de reloj (Activadores) en el
      menú izquierdo > "Añadir activador" >
        - Función: actualizarDropdownEquipos
        - Origen del evento: Desde la hoja de cálculo
        - Tipo de evento: Al enviar formulario
      Guarda. Desde ahora, cada vez que alguien llene CUALQUIER Form
      conectado a este Sheet (inscripción, roster, pagos), se refresca
      la lista de equipos en el Form de Integrantes.
   6. Para la página "Captura": menú Implementar > Nueva implementación >
        - Tipo: Aplicación web
        - Ejecutar como: Yo (tu cuenta)
        - Quién tiene acceso: Cualquier usuario
      Dale "Implementar", autoriza, y copia la URL que te da ("URL de la
      aplicación web"). Esa URL se pega en el archivo Captura_Liga_Basquetbol.html,
      dentro de CONFIG.appsScriptUrl.

   NOTA DE SEGURIDAD: como el Web App queda accesible para "Cualquier
   usuario" (para que la página de Captura le pueda escribir sin que el
   organizador tenga que iniciar sesión cada vez), cualquiera con esa URL
   podría en teoría mandarle datos. Para una liga amateur esto normalmente
   es aceptable, pero no compartas la URL del Web App públicamente —
   solo dásela a quien vaya a capturar resultados.
   ===================================================================== */

// ---- CONFIGURA ESTO ----
var FORM_ID_INTEGRANTES = "1gi0xneCKS1V08g15YFVBAva1aQFTMZmJW8m20SkzhGA"; // Form de Roster de Jugador
var TITULO_PREGUNTA_EQUIPO = "Equipo"; // debe coincidir con el título de esa pregunta en el Form
// Misma clave que CONFIG.adminPassword en Admin_Liga_Basquetbol.html. Protege
// la lectura de "Inscripción de Equipos" (tiene teléfonos y los códigos de
// equipo), que por eso NO está en la lista blanca pública de doGet.
var ADMIN_CLAVE_ = "liga2026";
// URL real de la Consola ya publicada (GitHub Pages + dominio propio).
var REGLAMENTO_URL_VARONIL_ = "https://www.laligadebasquet.com/?view=reglamento&cat=varonil";
var REGLAMENTO_URL_FEMENIL_ = "https://www.laligadebasquet.com/?view=reglamento&cat=femenil";
// -------------------------


/**
 * PASO 2 del flujo: copia los nombres de la hoja "Equipos" (columna B)
 * al menú desplegable "Equipo" del Form de Integrantes.
 */
function actualizarDropdownEquipos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaEquipos = ss.getSheetByName("Equipos");
  if (!hojaEquipos) {
    Logger.log("No se encontró la hoja 'Equipos'.");
    return;
  }

  var lastRow = hojaEquipos.getLastRow();
  if (lastRow < 2) return;

  var valores = hojaEquipos.getRange(2, 2, lastRow - 1, 1).getValues(); // columna B
  var equipos = valores
    .map(function (fila) { return String(fila[0]).trim(); })
    .filter(function (nombre) { return nombre !== ""; });

  if (equipos.length === 0) {
    Logger.log("Todavía no hay equipos en la hoja 'Equipos'.");
    return;
  }

  var form = FormApp.openById(FORM_ID_INTEGRANTES);
  var items = form.getItems();
  var actualizado = false;

  items.forEach(function (item) {
    var titulo = item.getTitle().trim().toLowerCase();
    if (titulo.indexOf(TITULO_PREGUNTA_EQUIPO.toLowerCase()) === -1) return;

    if (item.getType() === FormApp.ItemType.LIST) {
      item.asListItem().setChoiceValues(equipos);
      actualizado = true;
    } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
      item.asMultipleChoiceItem().setChoiceValues(equipos);
      actualizado = true;
    } else if (item.getType() === FormApp.ItemType.CHECKBOX) {
      item.asCheckboxItem().setChoiceValues(equipos);
      actualizado = true;
    }
  });

  if (!actualizado) {
    Logger.log("No se encontró una pregunta tipo lista/opción múltiple con título que contenga '" + TITULO_PREGUNTA_EQUIPO + "'. Revisa TITULO_PREGUNTA_EQUIPO arriba.");
  } else {
    Logger.log("Dropdown de equipos actualizado con: " + equipos.join(", "));
  }
}


/**
 * Recibe POST de la página "Admin" (pestañas Rol de Juego / Resultado y
 * Asistencia). El campo "action" dice qué hacer:
 *
 * action "programar" — agrega un partido nuevo a "Rol de Juego":
 * {
 *   "action": "programar",
 *   "jornada": 10, "fecha": "2026-09-19", "hora": "18:00",
 *   "sede": "Cancha Central", "categoria": "Varonil A",
 *   "equipoLocal": "Equipo 1", "equipoVisit": "Equipo 8",
 *   "arbitro": "R. Pérez"
 * }
 *
 * action "resultado" (o sin "action", por compatibilidad) — guarda el
 * marcador de un partido NUEVO y la asistencia de ambos rosters (pestaña
 * "Resultado y Asistencia" de Admin, solo partidos aún no jugados):
 * {
 *   "action": "resultado",
 *   "jornada": 3, "fecha": "2026-08-01",
 *   "equipoLocal": "Equipo 1", "ptsLocal": 78,
 *   "equipoVisit": "Equipo 8", "ptsVisit": 79,
 *   "sede": "Cancha Central",
 *   "asistencia": [
 *     {"equipo":"Equipo 1", "jugador":"Hugo Herrera", "asistio":"Sí"},
 *     {"equipo":"Equipo 8", "jugador":"...", "asistio":"No"}
 *   ]
 * }
 *
 * action "editar" — corrige el marcador y/o la asistencia de un partido que
 * YA estaba guardado (pestaña "Pasado" de Admin). A diferencia de "resultado",
 * esto NO agrega una fila nueva: busca la fila existente en "Resultados" por
 * jornada+equipos y la sobrescribe, y actualiza (o agrega si faltaba) cada
 * fila de "Asistencia" que corresponda. Mismo formato que "resultado".
 *
 * action "registrarPago" — agrega un pago/abono nuevo a "Pagos" (pestaña
 * "Pagos" de Admin, botón "Ingresar Pago"):
 * {
 *   "action": "registrarPago",
 *   "equipo": "Equipo 1", "categoria": "Varonil A",
 *   "concepto": "Inscripción Equipo", "monto": 1500,
 *   "metodoPago": "Efectivo", "quienRecibio": "Mauricio",
 *   "referencia": "", "fecha": "2026-07-14", "notas": ""
 * }
 *
 * action "inscribirEquipo" — alta de un equipo nuevo desde el formulario
 * público de la Consola. Escribe en "Inscripción de Equipos" y registra el
 * nombre oficial en "Equipos" (columna B) para que las fórmulas lo tomen.
 * Entra siempre como "Pendiente": no aparece como aprobado hasta que el
 * organizador lo apruebe desde el Admin. Genera un CÓDIGO único de 6
 * caracteres que la respuesta regresa en "codigo" — ese código es lo único
 * que necesitan los jugadores del equipo para darse de alta (así ningún
 * jugador ve la lista completa de equipos inscritos, solo el suyo).
 * Todos los campos son obligatorios, incluyendo "reglamentoAceptado" (el
 * capitán tiene que haber marcado el checkbox de "He leído el reglamento").
 * {
 *   "action": "inscribirEquipo",
 *   "nombreEquipo": "LOS HALCONES", "categoria": "Varonil A",
 *   "capitan": "Juan Pérez", "telCapitan": "8112345678",
 *   "subCapitan": "Luis Gómez", "telSubCapitan": "8187654321",
 *   "correo": "juan@mail.com", "menoresEdad": "Sí",
 *   "reglamentoAceptado": true, "notas": ""
 * }
 * Respuesta: { "ok": true, "codigo": "7K3PXA" }
 *
 * action "verificarCodigoEquipo" — resuelve un código de equipo a su nombre
 * y categoría, SIN escribir nada. Se usa en el formulario de Alta de Jugador
 * para mostrar "Bienvenido al equipo ___" antes de enviar el resto del form.
 * {
 *   "action": "verificarCodigoEquipo", "codigo": "7K3PXA"
 * }
 * Respuesta: { "ok": true, "equipo": "LOS HALCONES", "categoria": "Varonil A" }
 *
 * action "altaJugador" — alta de un jugador desde el formulario público.
 * El equipo se determina ÚNICAMENTE por el código (no por texto libre que
 * mande el navegador), así que el código es obligatorio. Escribe en
 * "Integrantes" (incluyendo el correo) y le manda un correo de bienvenida
 * al jugador con una vista previa de su credencial y un resumen del
 * reglamento (si el envío falla, no rompe el alta — solo se registra en
 * el Registro de ejecución).
 * "nombre" es SOLO el nombre de pila y "apellido" va aparte (así ya no se
 * confunde con "solo pusieron el nombre") — el servidor arma el nombre
 * completo uniendo ambos.
 * Si por "fechaNacimiento" el jugador resulta menor de edad (el servidor
 * vuelve a calcular la edad, no confía en lo que mande el navegador), son
 * obligatorios además "tutorNombre", "cartaResponsivaAceptada" (true) y
 * las fotos "ineFrente"/"ineReverso" (frente y reverso de la INE del padre,
 * madre o tutor que autoriza).
 * {
 *   "action": "altaJugador",
 *   "nombre": "Pedro", "apellido": "Fox", "fechaNacimiento": "2001-05-14",
 *   "correo": "pedro@mail.com",
 *   "tutorNombre": "María Fox", "cartaResponsivaAceptada": true,
 *   "ineFrente": "data:image/jpeg;base64,...", "ineReverso": "data:image/jpeg;base64,...",
 *   "codigo": "7K3PXA", "foto": "data:image/jpeg;base64,..."
 * }
 * Respuesta: { "ok": true, "equipo": "LOS HALCONES", "categoria": "Varonil A" }
 *
 * action "aprobarEquipo" — el organizador aprueba (o rechaza) un equipo
 * pendiente desde el Admin. Cambia la columna "Estatus" de la hoja
 * "Inscripción de Equipos".
 * {
 *   "action": "aprobarEquipo",
 *   "nombreEquipo": "LOS HALCONES", "estatus": "Aprobado"
 * }
 *
 * action "leerInscripciones" — el Admin la usa para leer TODA la hoja
 * "Inscripción de Equipos" (teléfonos y códigos incluidos). Por eso esa
 * hoja no está en la lista blanca pública de doGet: solo se entrega aquí,
 * y solo si "clave" coincide con ADMIN_CLAVE_ (la misma contraseña del Admin).
 * {
 *   "action": "leerInscripciones", "clave": "liga2026"
 * }
 * Respuesta: { "ok": true, "filas": [ {"Nombre del Equipo": "...", ...}, ... ] }
 *
 * action "reenviarBienvenida" — el organizador reenvía el correo de
 * bienvenida (con la credencial y el reglamento tal como están AHORA) a UN
 * jugador puntual que ya está dado de alta, buscándolo por correo en
 * "Integrantes". No vuelve a escribir nada, solo reenvía el correo. Sirve
 * para reenviar a alguien después de corregir el contenido del correo, sin
 * tener que darlo de alta otra vez. Protegida con "clave".
 * {
 *   "action": "reenviarBienvenida", "clave": "liga2026", "correo": "pedro@mail.com"
 * }
 * Respuesta: { "ok": true }
 *
 * action "reenviarBienvenidaCapitan" — igual que "reenviarBienvenida" pero
 * para el correo de bienvenida del CAPITÁN (nombre de equipo, categoría,
 * código de alta de jugadores), buscando por "Correo de Contacto" en
 * "Inscripción de Equipos". No genera código ni contraseña nuevos, reenvía
 * con los que ya existen. Protegida con "clave".
 * {
 *   "action": "reenviarBienvenidaCapitan", "clave": "liga2026", "correo": "pedro@mail.com"
 * }
 * Respuesta: { "ok": true }
 *
 * action "loginCapitan" — login del Portal de Capitanes (capitanes.html).
 * Valida correo+password contra "Inscripción de Equipos" (columnas Correo
 * de Contacto y Contraseña Capitán, esta última generada automáticamente
 * al inscribir el equipo y mandada por correo). Si coincide, regresa TODO
 * lo que el Portal necesita en un solo golpe: equipo, categoría, código,
 * roster (de "Integrantes") y resumen de pagos (de "Pagos" contra la cuota
 * fija CUOTA_EQUIPO_). Nunca expone datos de otros equipos.
 * {
 *   "action": "loginCapitan", "correo": "juan@mail.com", "password": "7K3PXA9M"
 * }
 * Respuesta: { "ok": true, "equipo": "...", "categoria": "...", "codigo": "...",
 *              "roster": [{"nombre":"...", "correo":"..."}, ...],
 *              "pagos": {"cuota":1500, "pagado":1500, "saldo":0, "historial":[...]} }
 *
 * action "cambiarPasswordCapitan" — el capitán cambia su propia contraseña
 * del Portal de Capitanes desde adentro del Portal. Valida la contraseña
 * actual antes de sobrescribirla.
 * {
 *   "action": "cambiarPasswordCapitan", "correo": "juan@mail.com",
 *   "passwordActual": "7K3PXA9M", "passwordNueva": "MiClaveNueva123"
 * }
 * Respuesta: { "ok": true }
 */
function doPost(e) {
  var resultado = { ok: false };
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.action === "programar") {
      programarPartido(ss, data);
    } else if (data.action === "editar") {
      editarResultado(ss, data);
      editarAsistencia(ss, data);
    } else if (data.action === "registrarPago") {
      registrarPago(ss, data);
    } else if (data.action === "inscribirEquipo") {
      var infoEq = inscribirEquipo(ss, data);
      resultado.codigo = infoEq.codigo;
    } else if (data.action === "altaJugador") {
      var infoJg = altaJugador(ss, data);
      resultado.equipo = infoJg.equipo;
      resultado.categoria = infoJg.categoria;
    } else if (data.action === "verificarCodigoEquipo") {
      var infoVer = verificarCodigoEquipo(ss, data);
      resultado.equipo = infoVer.nombreEquipo;
      resultado.categoria = infoVer.categoria;
    } else if (data.action === "aprobarEquipo") {
      aprobarEquipo(ss, data);
    } else if (data.action === "leerInscripciones") {
      resultado.filas = leerInscripciones(ss, data);
    } else if (data.action === "reenviarBienvenida") {
      var infoReenvio = reenviarCorreoBienvenidaJugador_(ss, data);
      resultado.enviados = infoReenvio.enviados;
    } else if (data.action === "reenviarBienvenidaCapitan") {
      var infoReenvioCap = reenviarCorreoBienvenidaCapitan_(ss, data);
      resultado.enviados = infoReenvioCap.enviados;
    } else if (data.action === "loginCapitan") {
      var infoCapitan = loginCapitan(ss, data);
      resultado.equipo = infoCapitan.equipo;
      resultado.categoria = infoCapitan.categoria;
      resultado.codigo = infoCapitan.codigo;
      resultado.roster = infoCapitan.roster;
      resultado.pagos = infoCapitan.pagos;
    } else if (data.action === "cambiarPasswordCapitan") {
      cambiarPasswordCapitan(ss, data);
    } else if (data.action === "leerPruebas") {
      resultado.hojas = leerPruebas(ss, data);
    } else {
      escribirResultado(ss, data);
      marcarRolComoJugado(ss, data);
      escribirAsistencia(ss, data);
    }

    resultado.ok = true;
  } catch (err) {
    resultado.error = String(err);
  }
  return ContentService.createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
   INSCRIPCIONES PÚBLICAS (formularios de la Consola)
   ===================================================================== */

/** Normaliza para comparar sin acentos ni mayúsculas. */
function normalizarTexto_(txt) {
  return String(txt == null ? "" : txt)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

/**
 * Detecta equipos de PRUEBA: cualquier equipo cuyo nombre normalizado EMPIECE
 * con "prueba" (no importan mayusculas ni acentos). Asi entran "Prueba",
 * "PRUEBA", "Prueba 2", "PRUEBA 3", "prueba final", etc., sin tener que tocar
 * el codigo cada vez que el organizador hace otra prueba.
 *
 * Se eligio "empieza con" en vez de coincidencia exacta a proposito: el
 * organizador numera sus pruebas ("Prueba 2") y con coincidencia exacta esos
 * equipos se colaban a las vistas publicas. El riesgo de que un equipo REAL
 * se llame "Prueba algo" es practicamente nulo, y si pasara no se pierde
 * nada: el equipo sigue existiendo y aparece en la pestaña "Pruebas" del
 * Admin, solo habria que renombrarlo.
 *
 * Estos equipos existen para que el organizador pueda probar el flujo
 * completo (inscripcion, alta de jugadores, pagos) sin que esos datos se
 * mezclen con los equipos reales en ninguna categoria ni vista publica. Se
 * usa para esconderlos de doGet() y de leerInscripciones(), y para juntarlos
 * aparte en leerPruebas().
 */
function esEquipoPrueba_(nombre) {
  return normalizarTexto_(nombre).indexOf("prueba") === 0;
}

// Caracteres para el código de equipo: sin 0/O, 1/I/L ni vocales que se
// puedan confundir al dictarlo por teléfono.
var CODIGO_EQUIPO_CHARS_ = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
var CODIGO_EQUIPO_LARGO_ = 6;
var CODIGO_EQUIPO_COL_ = 9; // columna I de "Inscripción de Equipos"
var MENORES_EDAD_COL_ = 10; // columna J de "Inscripción de Equipos"
var REGLAMENTO_ACEPTADO_COL_ = 11; // columna K de "Inscripción de Equipos"
var CORREO_EQUIPO_COL_ = 12; // columna L de "Inscripción de Equipos"
var CAPITAN_PASSWORD_COL_ = 13; // columna M de "Inscripción de Equipos"

// Caracteres para la contraseña del Portal de Capitanes: mayúsculas y
// números, sin 0/O/1/I para que no se confundan al leerla desde el correo.
var CAPITAN_PASSWORD_CHARS_ = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
var CAPITAN_PASSWORD_LARGO_ = 8;

/** Genera una contraseña aleatoria para el Portal de Capitanes. */
function generarPasswordCapitan_(){
  var pass = "";
  for (var i = 0; i < CAPITAN_PASSWORD_LARGO_; i++){
    pass += CAPITAN_PASSWORD_CHARS_.charAt(Math.floor(Math.random() * CAPITAN_PASSWORD_CHARS_.length));
  }
  return pass;
}

/** Genera un código de equipo que no choque con los que ya existen. */
function generarCodigoEquipo_(hojaInsc) {
  var existentes = {};
  var ultima = hojaInsc.getLastRow();
  if (ultima >= 2) {
    var col = hojaInsc.getRange(2, CODIGO_EQUIPO_COL_, ultima - 1, 1).getValues();
    col.forEach(function (f) {
      var v = String(f[0] || "").trim().toUpperCase();
      if (v) existentes[v] = true;
    });
  }
  for (var intento = 0; intento < 30; intento++) {
    var codigo = "";
    for (var k = 0; k < CODIGO_EQUIPO_LARGO_; k++) {
      codigo += CODIGO_EQUIPO_CHARS_.charAt(Math.floor(Math.random() * CODIGO_EQUIPO_CHARS_.length));
    }
    if (!existentes[codigo]) return codigo;
  }
  throw new Error("No se pudo generar un código único, intenta de nuevo.");
}

/**
 * Busca un equipo por su código en "Inscripción de Equipos".
 * Devuelve {nombreEquipo, categoria} o null si el código no existe.
 * No revela nada más (ni el resto de equipos, ni datos del capitán).
 */
function buscarEquipoPorCodigo_(ss, codigoCrudo) {
  var codigo = String(codigoCrudo || "").trim().toUpperCase();
  if (!codigo) return null;
  var hoja = ss.getSheetByName("Inscripción de Equipos");
  if (!hoja) return null;
  var ultima = hoja.getLastRow();
  if (ultima < 2) return null;
  var valores = hoja.getRange(2, 1, ultima - 1, CODIGO_EQUIPO_COL_).getValues(); // A:I
  for (var i = 0; i < valores.length; i++) {
    var fila = valores[i];
    var codigoFila = String(fila[CODIGO_EQUIPO_COL_ - 1] || "").trim().toUpperCase();
    if (codigoFila && codigoFila === codigo) {
      return { nombreEquipo: String(fila[1]).trim(), categoria: String(fila[3]).trim() };
    }
  }
  return null;
}

/** Resuelve un código a su equipo/categoría. Lanza error si no es válido. */
function verificarCodigoEquipo(ss, data) {
  var info = buscarEquipoPorCodigo_(ss, data.codigo);
  if (!info) throw new Error("Código no válido. Verifica con tu capitán.");
  return info;
}

/**
 * Da de alta un equipo nuevo.
 *  1. Agrega la fila de respuesta a "Inscripción de Equipos" (estatus Pendiente)
 *     con un código único de 6 caracteres en la columna "Código".
 *  2. Escribe el nombre oficial en la primera celda vacía de "Equipos" col. B,
 *     que es lo que dispara las fórmulas de esa hoja.
 * Rechaza nombres duplicados (comparando sin acentos ni mayúsculas).
 * Devuelve {codigo} para que la Consola se lo muestre al capitán.
 */
function inscribirEquipo(ss, data) {
  var nombre = String(data.nombreEquipo || "").trim();
  if (!nombre) throw new Error("Falta el nombre del equipo.");
  if (!data.categoria) throw new Error("Falta la categoría.");
  if (!data.capitan) throw new Error("Falta el capitán.");
  if (!data.telCapitan) throw new Error("Falta el teléfono del capitán.");
  if (!data.subCapitan) throw new Error("Falta el sub-capitán.");
  if (!data.telSubCapitan) throw new Error("Falta el teléfono del sub-capitán.");
  var correoEq = String(data.correo || "").trim();
  if (!correoEq) throw new Error("Falta el correo de contacto.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correoEq)) throw new Error("El correo de contacto no es válido.");
  if (!data.reglamentoAceptado) throw new Error("Falta aceptar el reglamento de la liga.");

  var hojaInsc = ss.getSheetByName("Inscripción de Equipos");
  if (!hojaInsc) throw new Error("No existe la hoja 'Inscripción de Equipos'.");
  var hojaEquipos = ss.getSheetByName("Equipos");
  if (!hojaEquipos) throw new Error("No existe la hoja 'Equipos'.");

  // --- No permitir nombres repetidos ---
  var objetivo = normalizarTexto_(nombre);
  var ultimaEq = hojaEquipos.getLastRow();
  if (ultimaEq >= 2) {
    var existentes = hojaEquipos.getRange(2, 2, ultimaEq - 1, 1).getValues();
    for (var i = 0; i < existentes.length; i++) {
      if (normalizarTexto_(existentes[i][0]) === objetivo) {
        throw new Error("Ya hay un equipo inscrito con el nombre '" + nombre + "'.");
      }
    }
  }

  // El encabezado de la columna Código se pone solo la primera vez que hace falta.
  if (String(hojaInsc.getRange(1, CODIGO_EQUIPO_COL_).getValue()).trim() === "") {
    hojaInsc.getRange(1, CODIGO_EQUIPO_COL_).setValue("Código");
  }
  if (String(hojaInsc.getRange(1, MENORES_EDAD_COL_).getValue()).trim() === "") {
    hojaInsc.getRange(1, MENORES_EDAD_COL_).setValue("Menores de Edad");
  }
  if (String(hojaInsc.getRange(1, REGLAMENTO_ACEPTADO_COL_).getValue()).trim() === "") {
    hojaInsc.getRange(1, REGLAMENTO_ACEPTADO_COL_).setValue("Reglamento Aceptado");
  }
  if (String(hojaInsc.getRange(1, CORREO_EQUIPO_COL_).getValue()).trim() === "") {
    hojaInsc.getRange(1, CORREO_EQUIPO_COL_).setValue("Correo de Contacto");
  }
  if (String(hojaInsc.getRange(1, CAPITAN_PASSWORD_COL_).getValue()).trim() === "") {
    hojaInsc.getRange(1, CAPITAN_PASSWORD_COL_).setValue("Contraseña Capitán");
  }

  var codigo = generarCodigoEquipo_(hojaInsc);
  var passwordCapitan = generarPasswordCapitan_();
  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  var menoresEdad = String(data.menoresEdad || "No").trim() === "Sí" ? "Sí" : "No";

  // Capitán / Sub-capitán van juntos en una sola celda, igual que el Form viejo.
  var capSub = data.capitan + (data.telCapitan ? " " + data.telCapitan : "");
  if (data.subCapitan) {
    capSub += " / " + data.subCapitan + (data.telSubCapitan ? " " + data.telSubCapitan : "");
  }

  // Columnas: Timestamp, Nombre del Equipo, Capitán / Sub-capitán, Categoría,
  //           Integrantes, Horario Confirmado, Pago Confirmado, Estatus, Código,
  //           Menores de Edad, Reglamento Aceptado, Correo de Contacto,
  //           Contraseña Capitán
  hojaInsc.appendRow([
    ahora,
    nombre,
    capSub,
    data.categoria,
    "",              // los jugadores se dan de alta uno por uno en su propio formulario
    "Pendiente",
    "Pendiente",
    "Pendiente",
    codigo,
    menoresEdad,
    "Sí",            // Reglamento Aceptado — ya se validó arriba que sea obligatorio
    correoEq,
    passwordCapitan
  ]);

  // --- Registrar el nombre oficial en "Equipos" (columna B) ---
  // La hoja viene con filas pre-cargadas de fórmulas, así que buscamos la
  // primera celda vacía de la columna B en vez de usar appendRow.
  var filaDestino = 0;
  var maxFilas = hojaEquipos.getMaxRows();
  var colB = hojaEquipos.getRange(2, 2, maxFilas - 1, 1).getValues();
  for (var j = 0; j < colB.length; j++) {
    if (String(colB[j][0]).trim() === "") { filaDestino = j + 2; break; }
  }
  if (!filaDestino) filaDestino = maxFilas + 1; // por si ya se llenaron todas
  hojaEquipos.getRange(filaDestino, 2).setValue(nombre);

  // El correo de bienvenida al capitán NO debe romper el alta si falla,
  // igual que el de bienvenida al jugador.
  try {
    enviarCorreoBienvenidaCapitan_(correoEq, nombre, data.categoria, codigo, passwordCapitan);
  } catch (errCorreoCap) {
    Logger.log("No se pudo enviar el correo de bienvenida al capitán de " + nombre + ": " + errCorreoCap);
  }

  return { codigo: codigo };
}

/**
 * Calcula la edad en años cumplidos a partir de una fecha "yyyy-MM-dd"
 * (el formato que manda <input type="date">). Devuelve null si la fecha
 * no se puede interpretar.
 */
function calcularEdad_(fechaISO) {
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(fechaISO || "").trim());
  if (!m) return null;
  var nacimiento = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (isNaN(nacimiento.getTime())) return null;
  var hoy = new Date();
  var edad = hoy.getFullYear() - nacimiento.getFullYear();
  var aunNoCumple = (hoy.getMonth() < nacimiento.getMonth()) ||
    (hoy.getMonth() === nacimiento.getMonth() && hoy.getDate() < nacimiento.getDate());
  if (aunNoCumple) edad--;
  return edad;
}

/**
 * Da de alta un jugador en "Integrantes". El equipo NO lo manda el
 * navegador como texto libre: se resuelve a partir del código del equipo,
 * así ningún jugador ve la lista de equipos inscritos, solo confirma el
 * suyo con "Bienvenido al equipo ___" antes de enviar el formulario.
 *
 * Si por la fecha de nacimiento el jugador es menor de edad, son
 * obligatorios además: el nombre del padre/madre/tutor, que haya aceptado
 * la carta responsiva de menores, y las fotos (frente y reverso) de la
 * INE de quien autoriza. Esto se vuelve a validar aquí (no solo en el
 * navegador) para que nadie pueda saltárselo editando el JavaScript de la
 * página.
 *
 * Devuelve {equipo, categoria} (el nombre real, resuelto por el servidor).
 */
function altaJugador(ss, data) {
  var nombrePila = String(data.nombre || "").trim();
  if (!nombrePila) throw new Error("Falta el nombre del jugador.");
  var apellido = String(data.apellido || "").trim();
  if (!apellido) throw new Error("Falta el apellido del jugador.");
  var fechaNacimiento = String(data.fechaNacimiento || "").trim();
  if (!fechaNacimiento) throw new Error("Falta la fecha de nacimiento del jugador.");
  var nombre = (nombrePila + " " + apellido).trim();
  var correo = String(data.correo || "").trim();
  if (!correo) throw new Error("Falta el correo del jugador.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) throw new Error("El correo no es válido.");

  var edad = calcularEdad_(fechaNacimiento);
  var esMenor = edad !== null && edad >= 0 && edad < 18;
  var tutorNombre = String(data.tutorNombre || "").trim();
  var ineFrente = String(data.ineFrente || "");
  var ineReverso = String(data.ineReverso || "");
  var cartaAceptada = !!data.cartaResponsivaAceptada;
  if (esMenor) {
    if (!tutorNombre) throw new Error("Falta el nombre del padre, madre o tutor.");
    if (!cartaAceptada) throw new Error("Falta aceptar la carta responsiva para menores de edad.");
    if (!ineFrente) throw new Error("Falta la foto del frente de la INE del padre, madre o tutor.");
    if (!ineReverso) throw new Error("Falta la foto del reverso de la INE del padre, madre o tutor.");
  }

  var info = verificarCodigoEquipo(ss, data); // lanza error si el código no es válido
  var equipo = info.nombreEquipo;
  var categoria = info.categoria;

  var hojaEquipos = ss.getSheetByName("Equipos");
  if (!hojaEquipos) throw new Error("No existe la hoja 'Equipos'.");
  var hojaInt = ss.getSheetByName("Integrantes");
  if (!hojaInt) throw new Error("No existe la hoja 'Integrantes'.");

  // --- El equipo debe seguir existiendo en "Equipos" ---
  var existe = false;
  var ultimaEq = hojaEquipos.getLastRow();
  if (ultimaEq >= 2) {
    var equipos = hojaEquipos.getRange(2, 2, ultimaEq - 1, 1).getValues();
    for (var i = 0; i < equipos.length; i++) {
      if (String(equipos[i][0]).trim() === equipo) { existe = true; break; }
    }
  }
  if (!existe) {
    throw new Error("Ese equipo ya no aparece en 'Equipos'. Avísale al organizador.");
  }

  // --- No repetir al mismo jugador en el mismo equipo ---
  var objetivo = normalizarTexto_(nombre);
  var ultimaInt = hojaInt.getLastRow();
  if (ultimaInt >= 2) {
    var filas = hojaInt.getRange(2, 2, ultimaInt - 1, 2).getValues(); // Nombre, Equipo
    for (var k = 0; k < filas.length; k++) {
      if (normalizarTexto_(filas[k][0]) === objetivo && String(filas[k][1]).trim() === equipo) {
        throw new Error("'" + nombre + "' ya está dado de alta en " + equipo + ".");
      }
    }
  }

  // Los encabezados de columnas nuevas se ponen solo la primera vez que hacen falta.
  if (String(hojaInt.getRange(1, 5).getValue()).trim() === "") {
    hojaInt.getRange(1, 5).setValue("Correo");
  }
  if (String(hojaInt.getRange(1, 6).getValue()).trim() === "") {
    hojaInt.getRange(1, 6).setValue("Apellido");
  }
  if (String(hojaInt.getRange(1, 7).getValue()).trim() === "") {
    hojaInt.getRange(1, 7).setValue("Fecha de Nacimiento");
  }
  if (String(hojaInt.getRange(1, 8).getValue()).trim() === "") {
    hojaInt.getRange(1, 8).setValue("Aviso de Imagen");
  }
  if (String(hojaInt.getRange(1, 9).getValue()).trim() === "") {
    hojaInt.getRange(1, 9).setValue("Es Menor de Edad");
  }
  if (String(hojaInt.getRange(1, 10).getValue()).trim() === "") {
    hojaInt.getRange(1, 10).setValue("Nombre del Tutor");
  }
  if (String(hojaInt.getRange(1, 11).getValue()).trim() === "") {
    hojaInt.getRange(1, 11).setValue("Carta Responsiva Aceptada");
  }
  if (String(hojaInt.getRange(1, 12).getValue()).trim() === "") {
    hojaInt.getRange(1, 12).setValue("INE Frente");
  }
  if (String(hojaInt.getRange(1, 13).getValue()).trim() === "") {
    hojaInt.getRange(1, 13).setValue("INE Reverso");
  }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  // Columnas: Timestamp, Nombre Completo, Equipo, Foto, Correo, Apellido,
  //           Fecha de Nacimiento, Aviso de Imagen, Es Menor de Edad,
  //           Nombre del Tutor, Carta Responsiva Aceptada, INE Frente, INE Reverso
  hojaInt.appendRow([
    ahora, nombre, equipo, data.foto || "", correo, apellido, fechaNacimiento, "Sí",
    esMenor ? "Sí" : "No",
    esMenor ? tutorNombre : "",
    esMenor ? "Sí" : "",
    esMenor ? ineFrente : "",
    esMenor ? ineReverso : ""
  ]);

  // El correo de bienvenida NO debe romper el alta si falla (cuota de
  // MailApp agotada, correo raro, etc.) — por eso va en su propio try/catch.
  try {
    enviarCorreoBienvenidaJugador_(correo, nombre, equipo, categoria, data.foto || "");
  } catch (errCorreo) {
    Logger.log("No se pudo enviar el correo de bienvenida a " + correo + ": " + errCorreo);
  }

  return { equipo: equipo, categoria: categoria };
}

/**
 * Manda el correo de bienvenida a un jugador recién dado de alta: mensaje
 * de bienvenida, una vista previa de cómo se ve su credencial de jugador
 * (mismo diseño que la Consola, armada con tablas y estilos en línea para
 * que se vea bien en Gmail/Outlook/celular), confirmación de equipo/categoría
 * y un resumen de las reglas más importantes (no el reglamento completo,
 * para que sea fácil de leer desde el celular). Usa MailApp, que es gratis
 * con la cuota normal de Gmail (no requiere ningún servicio de pago).
 * "foto" es el data URI (base64) que subió el jugador, o "" si no subió
 * ninguna — en ese caso la credencial muestra sus iniciales. La foto va
 * como imagen incrustada (inlineImages + cid:) en vez de src="data:...",
 * porque Gmail y otros clientes bloquean las imágenes con src="data:..."
 * en correos recibidos y se vería en blanco. El link "click aquí" del
 * resumen del reglamento apunta a REGLAMENTO_URL_VARONIL_ o
 * REGLAMENTO_URL_FEMENIL_ según si "categoria" contiene "femenil" (ambos
 * son placeholders — hay que reemplazarlos por la URL real ya publicada).
 */
function enviarCorreoBienvenidaJugador_(correo, nombre, equipo, categoria, foto) {
  var asunto = "¡Bienvenido a La Liga de Basquet, " + nombre + "!";
  var iniciales = nombre.split(" ").filter(function (p) { return p; })
    .map(function (p) { return p.charAt(0); }).slice(0, 2).join("").toUpperCase();
  var urlReglamento = /femenil/i.test(categoria || "") ? REGLAMENTO_URL_FEMENIL_ : REGLAMENTO_URL_VARONIL_;

  var fotoBlob = null;
  if (foto && foto.indexOf("data:") === 0 && foto.indexOf(",") > -1) {
    try {
      var comaIdx = foto.indexOf(",");
      var contentType = foto.substring(5, comaIdx).split(";")[0] || "image/jpeg";
      var base64Data = foto.substring(comaIdx + 1);
      fotoBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, "foto.jpg");
    } catch (errFoto) {
      fotoBlob = null;
    }
  }
  var fotoCeldaHtml = fotoBlob
    ? '<img src="cid:fotoJugador" width="72" height="80" style="display:block;width:72px;height:80px;object-fit:cover;">'
    : '<span style="font-family:Arial,Helvetica,sans-serif;font-size:20px;font-weight:800;color:#6B6B6B;">' + iniciales + '</span>';

  var credencialHtml =
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="280" align="center" ' +
      'style="margin:6px auto 4px auto;border:2px solid #000000;border-radius:10px;background:#ffffff;font-family:Arial,Helvetica,sans-serif;">' +
      '<tr><td style="padding:10px 14px 8px 14px;text-align:center;border-bottom:2px solid #000000;">' +
        '<div style="font-weight:900;font-size:14px;color:#000000;text-transform:uppercase;letter-spacing:-.2px;">LA LIGA DE BASQUET</div>' +
        '<div style="font-weight:400;font-size:8px;color:#000000;text-transform:uppercase;letter-spacing:2.2px;margin-top:2px;">GANTE SAN PEDRO</div>' +
        '<div style="font-weight:700;font-size:8px;color:#6B6B6B;text-transform:uppercase;letter-spacing:1.6px;margin-top:2px;">TEMPORADA AGOSTO 2026</div>' +
      '</td></tr>' +
      '<tr><td style="padding:12px 14px 14px 14px;">' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>' +
          '<td width="72" valign="top" style="padding-right:10px;">' +
            '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="72" ' +
              'style="border:1px solid #111111;background:#F5F5F5;">' +
              '<tr><td align="center" valign="middle" width="72" height="80">' + fotoCeldaHtml + '</td></tr>' +
            '</table>' +
          '</td>' +
          '<td valign="top" style="padding-top:4px;">' +
            '<div style="font-weight:800;font-size:15px;color:#111111;text-transform:uppercase;line-height:1.3;">' + nombre + '</div>' +
            '<div style="font-size:10px;font-weight:700;color:#6B6B6B;margin-top:4px;letter-spacing:.5px;">' + categoria + '</div>' +
            '<div style="font-weight:700;font-size:12px;color:#6B6B6B;text-transform:uppercase;letter-spacing:.3px;margin-top:6px;">' + equipo + '</div>' +
          '</td>' +
        '</tr></table>' +
      '</td></tr>' +
      '<tr><td style="padding:0 14px 12px 14px;text-align:center;">' +
        '<div style="font-size:18px;letter-spacing:2px;">🏀🏀🏀</div>' +
        '<div style="font-size:9px;color:#6B6B6B;margin-top:2px;">Cada 🏀 representa un partido jugado</div>' +
      '</td></tr>' +
    '</table>' +
    '<p style="text-align:center;font-size:11px;color:#6B6B6B;font-style:italic;margin:0 0 20px 0;">' +
      'Así se ve tu credencial. Consúltala completa (con tus puntos anotados) en la sección "Credencial" de la Consola.</p>';
  var htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto;color:#111111;">' +
      '<div style="background:#000000;padding:20px;text-align:center;border-bottom:4px solid #F37228;">' +
        '<div style="color:#ffffff;font-weight:900;font-size:18px;letter-spacing:-.3px;">LA LIGA DE BASQUET</div>' +
        '<div style="color:#ffffff;font-weight:400;font-size:10px;letter-spacing:3px;margin-top:4px;">GANTE SAN PEDRO</div>' +
      '</div>' +
      '<div style="padding:24px 20px;">' +
        '<p style="font-size:16px;">¡Bienvenido, <strong>' + nombre + '</strong>! 🏀</p>' +
        '<p>Ya quedaste dado de alta en <strong>' + equipo + '</strong>, categoría <strong>' + categoria + '</strong>.</p>' +
        credencialHtml +
        '<div style="background:#F5F5F5;border-radius:10px;padding:16px 18px;margin:20px 0;">' +
          '<p style="font-weight:800;font-size:13px;text-transform:uppercase;letter-spacing:.5px;margin:0 0 10px;">Lo esencial del reglamento</p>' +
          '<ul style="padding-left:18px;margin:0;font-size:14px;line-height:1.7;">' +
            '<li>Para playoffs necesitas haber jugado al menos 2 partidos de temporada regular.</li>' +
            '<li>Uniforme: mismo color de jersey que tu equipo, con número visible, NO se puede repetir número.</li>' +
            '<li>Para ganar el MVP de temporada regular (el jugador con más puntos anotados durante la temporada) hay que disputar mínimo 4 partidos.</li>' +
            '<li>Se toman en cuenta los puntos anotados en cada juego.</li>' +
            '<li>Todos los partidos son transmitidos por nuestro canal de YouTube.</li>' +
          '</ul>' +
        '</div>' +
        '<p style="font-size:13px;color:#6B6B6B;">Este es un resumen. Si quieres verlo completo, ' +
          '<a href="' + urlReglamento + '" style="color:#F37228;font-weight:700;">haz click aquí</a>.</p>' +
        '<p style="font-size:14px;">Gracias por tu preferencia.</p>' +
        '<div style="border-top:1px solid #E2E2E2;margin-top:20px;padding-top:16px;font-size:13px;color:#6B6B6B;">' +
          '<p style="margin:0 0 4px;"><strong>Contacto</strong></p>' +
          '<p style="margin:0;">81 1781 7451 · control@laligadebasquet.com</p>' +
          '<p style="margin:12px 0 0;">' +
            '<a href="https://www.facebook.com/laligadebasquet/" style="color:#F37228;font-weight:700;text-decoration:none;">Facebook</a>' +
            '&nbsp;·&nbsp;' +
            '<a href="https://www.instagram.com/laligadebasquet/" style="color:#F37228;font-weight:700;text-decoration:none;">Instagram</a>' +
            '&nbsp;·&nbsp;' +
            '<a href="https://www.tiktok.com/@laligadebasquet?lang=es-419" style="color:#F37228;font-weight:700;text-decoration:none;">TikTok</a>' +
            '&nbsp;·&nbsp;' +
            '<a href="https://www.youtube.com/channel/UCHH7p6lP-rhy6Xa7LCuZFKA" style="color:#F37228;font-weight:700;text-decoration:none;">YouTube</a>' +
          '</p>' +
        '</div>' +
      '</div>' +
    '</div>';
  // Envío vía Resend (API transaccional) para que el remitente real sea
  // control@laligadebasquet.com. Requiere la propiedad de script
  // RESEND_API_KEY (Configuración del proyecto > Propiedades del script).
  var apiKeyResend = PropertiesService.getScriptProperties().getProperty("RESEND_API_KEY");
  if (!apiKeyResend) {
    throw new Error("Falta configurar RESEND_API_KEY en las Propiedades del script.");
  }
  var payloadResend = {
    from: "La Liga de Basquet · Gante San Pedro <control@laligadebasquet.com>",
    to: correo,
    subject: asunto,
    html: htmlBody
  };
  if (fotoBlob) {
    payloadResend.attachments = [{
      filename: "foto.jpg",
      content: base64Data,
      content_id: "fotoJugador"
    }];
  }
  var respuestaResend = UrlFetchApp.fetch("https://api.resend.com/emails", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKeyResend },
    payload: JSON.stringify(payloadResend),
    muteHttpExceptions: true
  });
  var codigoResend = respuestaResend.getResponseCode();
  if (codigoResend < 200 || codigoResend >= 300) {
    throw new Error("Resend respondió " + codigoResend + ": " + respuestaResend.getContentText());
  }
}

/**
 * Manda el correo de bienvenida al CAPITÁN justo después de inscribir su
 * equipo (acción "inscribirEquipo"). Incluye: nombre del equipo, categoría,
 * un link a la página de Reglamento (que ya incluye la sección de formas
 * de pago) para la categoría correspondiente, y el código de 6 caracteres
 * que sus jugadores van a necesitar para darse de alta uno por uno (con un
 * cuadro fácil de copiar con el link directo a Alta de Jugador). También le
 * recuerda que él/ella también se tiene que dar de alta como jugador. El
 * Portal de Capitanes está temporalmente oculto (ver nota en index.html),
 * así que este correo NO lo menciona. Usa Resend, igual que el correo de
 * bienvenida de jugador — si RESEND_API_KEY no está configurada, revienta
 * (el llamador de esta función ya la envuelve en try/catch).
 */
function enviarCorreoBienvenidaCapitan_(correo, nombreEquipo, categoria, codigo, passwordCapitan) {
  var asunto = "¡Bienvenido a La Liga de Basquet! Equipo " + nombreEquipo;
  var urlReglamento = /femenil/i.test(categoria || "") ? REGLAMENTO_URL_FEMENIL_ : REGLAMENTO_URL_VARONIL_;
  // Debe ser EXACTAMENTE esta URL: lleva directo a la pestaña "Alta de
  // Jugador" dentro de Inscripción en la Consola (ver aplicarDeepLink() en
  // index.html, que lee ?view=inscripcion&tab=jugador).
  var urlAltaJugador = "https://www.laligadebasquet.com/?view=inscripcion&tab=jugador";

  var htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto;color:#111111;">' +
      '<div style="background:#000000;padding:20px;text-align:center;border-bottom:4px solid #F37228;">' +
        '<div style="color:#ffffff;font-weight:900;font-size:18px;letter-spacing:-.3px;">LA LIGA DE BASQUET</div>' +
        '<div style="color:#ffffff;font-weight:400;font-size:10px;letter-spacing:3px;margin-top:4px;">GANTE SAN PEDRO</div>' +
      '</div>' +
      '<div style="padding:24px 20px;">' +
        '<p style="font-size:16px;">¡Bienvenido a La Liga de Basquet! 🏀</p>' +
        '<p>Tu equipo quedó registrado con estos datos:</p>' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" ' +
          'style="border:1px solid #E2E2E2;border-radius:8px;margin:14px 0 20px 0;">' +
          '<tr><td style="padding:10px 14px;border-bottom:1px solid #E2E2E2;font-size:11px;font-weight:700;color:#6B6B6B;text-transform:uppercase;">Nombre del equipo</td>' +
            '<td style="padding:10px 14px;border-bottom:1px solid #E2E2E2;font-size:14px;font-weight:800;text-align:right;">' + nombreEquipo + '</td></tr>' +
          '<tr><td style="padding:10px 14px;font-size:11px;font-weight:700;color:#6B6B6B;text-transform:uppercase;">Categoría</td>' +
            '<td style="padding:10px 14px;font-size:14px;font-weight:800;text-align:right;">' + categoria + '</td></tr>' +
        '</table>' +
        '<p style="font-size:13px;">Ahí también vienen las formas de pago aceptadas. ' +
          '<a href="' + urlReglamento + '" style="color:#F37228;font-weight:700;">Haz click aquí para ver el reglamento completo y las formas de pago</a>.</p>' +
        '<div style="background:#F5F5F5;border-radius:10px;padding:16px 18px;margin:20px 0;text-align:center;">' +
          '<p style="font-weight:800;font-size:12px;text-transform:uppercase;letter-spacing:.5px;margin:0 0 8px;color:#6B6B6B;">Código para alta de jugadores</p>' +
          '<p style="font-weight:900;font-size:28px;letter-spacing:4px;color:#111111;margin:0;">' + codigo + '</p>' +
          '<p style="font-size:12px;color:#6B6B6B;margin:8px 0 0;">Compártelo con tu equipo. Cada jugador lo necesita para darse de alta en ' +
            '<a href="' + urlAltaJugador + '" style="color:#F37228;font-weight:700;">Alta de Jugador</a>.</p>' +
        '</div>' +
        '<p style="font-size:13px;background:rgba(243,114,40,.1);border:1px solid #F37228;border-radius:8px;padding:12px 14px;">' +
          '<strong>No se te olvide:</strong> tú también te tienes que dar de alta como jugador con el código de arriba, en ' +
          '<a href="' + urlAltaJugador + '" style="color:#F37228;font-weight:700;">Alta de Jugador</a>.</p>' +
        '<p style="font-size:14px;">Gracias por tu preferencia.</p>' +
        '<div style="border-top:1px solid #E2E2E2;margin-top:20px;padding-top:16px;font-size:13px;color:#6B6B6B;">' +
          '<p style="margin:0 0 4px;"><strong>Contacto</strong></p>' +
          '<p style="margin:0;">81 1781 7451 · control@laligadebasquet.com</p>' +
          '<p style="margin:12px 0 0;">' +
            '<a href="https://www.facebook.com/laligadebasquet/" style="color:#F37228;font-weight:700;text-decoration:none;">Facebook</a>' +
            '&nbsp;·&nbsp;' +
            '<a href="https://www.instagram.com/laligadebasquet/" style="color:#F37228;font-weight:700;text-decoration:none;">Instagram</a>' +
            '&nbsp;·&nbsp;' +
            '<a href="https://www.tiktok.com/@laligadebasquet?lang=es-419" style="color:#F37228;font-weight:700;text-decoration:none;">TikTok</a>' +
            '&nbsp;·&nbsp;' +
            '<a href="https://www.youtube.com/channel/UCHH7p6lP-rhy6Xa7LCuZFKA" style="color:#F37228;font-weight:700;text-decoration:none;">YouTube</a>' +
          '</p>' +
        '</div>' +
      '</div>' +
    '</div>';

  var apiKeyResend = PropertiesService.getScriptProperties().getProperty("RESEND_API_KEY");
  if (!apiKeyResend) {
    throw new Error("Falta configurar RESEND_API_KEY en las Propiedades del script.");
  }
  var respuestaResend = UrlFetchApp.fetch("https://api.resend.com/emails", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKeyResend },
    payload: JSON.stringify({
      from: "La Liga de Basquet · Gante San Pedro <control@laligadebasquet.com>",
      to: correo,
      subject: asunto,
      html: htmlBody
    }),
    muteHttpExceptions: true
  });
  var codigoResend = respuestaResend.getResponseCode();
  if (codigoResend < 200 || codigoResend >= 300) {
    throw new Error("Resend respondió " + codigoResend + ": " + respuestaResend.getContentText());
  }
}

/**
 * Cambia el estatus de un equipo en "Inscripción de Equipos".
 * Se usa desde el Admin para aprobar o rechazar una inscripción pendiente.
 */
function aprobarEquipo(ss, data) {
  var nombre = String(data.nombreEquipo || "").trim();
  var estatus = String(data.estatus || "Aprobado").trim();
  if (!nombre) throw new Error("Falta el nombre del equipo.");

  var hoja = ss.getSheetByName("Inscripción de Equipos");
  if (!hoja) throw new Error("No existe la hoja 'Inscripción de Equipos'.");

  var ultima = hoja.getLastRow();
  if (ultima < 2) throw new Error("No hay inscripciones registradas.");

  var nombres = hoja.getRange(2, 2, ultima - 1, 1).getValues();
  var objetivo = normalizarTexto_(nombre);
  for (var i = 0; i < nombres.length; i++) {
    if (normalizarTexto_(nombres[i][0]) === objetivo) {
      hoja.getRange(i + 2, 8).setValue(estatus); // columna H = Estatus
      return;
    }
  }
  throw new Error("No encontré la inscripción del equipo '" + nombre + "'.");
}

/**
 * Lectura protegida de "Inscripción de Equipos" (teléfonos + códigos de
 * equipo). Solo la usa el Admin, mandando la misma contraseña que ya pide
 * su candado. Esta hoja se sacó de HOJAS_PERMITIDAS (doGet) precisamente
 * para que nadie pueda leerla anónimamente con solo la URL del Web App.
 */
function leerInscripciones(ss, data) {
  if (String(data.clave || "") !== ADMIN_CLAVE_) {
    throw new Error("No autorizado.");
  }
  var filas = leerHojaComoObjetos_(ss, "Inscripción de Equipos");
  // Los equipos de PRUEBA no se mezclan aquí: van aparte, ver leerPruebas().
  return filas.filter(function (f) { return !esEquipoPrueba_(f["Nombre del Equipo"]); });
}

/**
 * Reenvía el correo de bienvenida a TODOS los jugadores dados de alta con
 * ese correo (un mismo correo puede tener más de un jugador dado de alta,
 * por ejemplo un papá que registra a varios hijos, o pruebas), buscándolos
 * en "Integrantes" (columna E). Resuelve la categoría de cada uno buscando
 * su equipo en "Inscripción de Equipos" (columna B -> categoría en columna D).
 * No modifica ninguna hoja, solo vuelve a mandar el correo con el contenido
 * actual de enviarCorreoBienvenidaJugador_, uno por cada jugador encontrado.
 * Protegida con ADMIN_CLAVE_ porque, a diferencia de "altaJugador", cualquiera
 * con la URL podría spamear a un jugador si no se protegiera.
 * Devuelve { enviados: N } con cuántos correos se reenviaron.
 */
function reenviarCorreoBienvenidaJugador_(ss, data) {
  if (String(data.clave || "") !== ADMIN_CLAVE_) {
    throw new Error("No autorizado.");
  }
  var correoObjetivo = normalizarTexto_(data.correo);
  if (!correoObjetivo) throw new Error("Falta el correo.");

  var hojaInt = ss.getSheetByName("Integrantes");
  if (!hojaInt) throw new Error("No existe la hoja 'Integrantes'.");
  var ultimaInt = hojaInt.getLastRow();
  if (ultimaInt < 2) throw new Error("Ese correo no está dado de alta.");

  // Columnas: Timestamp, Nombre Completo, Equipo, Foto, Correo, ...
  var filas = hojaInt.getRange(2, 1, ultimaInt - 1, 5).getValues();
  var encontrados = [];
  for (var i = 0; i < filas.length; i++) {
    if (normalizarTexto_(filas[i][4]) === correoObjetivo) {
      encontrados.push({ nombre: String(filas[i][1]).trim(), equipo: String(filas[i][2]).trim(), foto: String(filas[i][3] || "") });
    }
  }
  if (!encontrados.length) throw new Error("Ese correo no está dado de alta.");

  // La categoría del equipo vive en "Inscripción de Equipos" (columna D),
  // no en "Equipos" (esa hoja solo guarda el nombre oficial, columna B).
  var hojaInsc = ss.getSheetByName("Inscripción de Equipos");
  if (!hojaInsc) throw new Error("No existe la hoja 'Inscripción de Equipos'.");
  var mapaCategorias = {};
  var ultimaInsc = hojaInsc.getLastRow();
  if (ultimaInsc >= 2) {
    var inscFilas = hojaInsc.getRange(2, 2, ultimaInsc - 1, 3).getValues(); // Nombre Equipo, Cap/Sub, Categoría
    for (var j = 0; j < inscFilas.length; j++) {
      mapaCategorias[String(inscFilas[j][0]).trim()] = String(inscFilas[j][2]).trim();
    }
  }

  encontrados.forEach(function (jg) {
    var categoria = mapaCategorias[jg.equipo] || "";
    enviarCorreoBienvenidaJugador_(data.correo, jg.nombre, jg.equipo, categoria, jg.foto);
  });

  return { enviados: encontrados.length };
}

/**
 * Reenvía el correo de bienvenida al CAPITÁN (con el contenido tal como
 * está AHORA, por ejemplo después de corregirlo) a TODOS los equipos cuya
 * "Correo de Contacto" coincida con el correo dado, buscándolos en
 * "Inscripción de Equipos". No vuelve a generar código ni contraseña, usa
 * los que ya están guardados. Protegida con "clave", igual que
 * reenviarBienvenida.
 */
function reenviarCorreoBienvenidaCapitan_(ss, data) {
  if (String(data.clave || "") !== ADMIN_CLAVE_) {
    throw new Error("No autorizado.");
  }
  var correoObjetivo = normalizarTexto_(data.correo);
  if (!correoObjetivo) throw new Error("Falta el correo.");

  var hojaInsc = ss.getSheetByName("Inscripción de Equipos");
  if (!hojaInsc) throw new Error("No existe la hoja 'Inscripción de Equipos'.");
  var ultimaInsc = hojaInsc.getLastRow();
  if (ultimaInsc < 2) throw new Error("Ese correo no está dado de alta.");

  var filas = hojaInsc.getRange(2, 1, ultimaInsc - 1, CAPITAN_PASSWORD_COL_).getValues(); // A:M
  var encontrados = [];
  for (var i = 0; i < filas.length; i++) {
    if (normalizarTexto_(filas[i][CORREO_EQUIPO_COL_ - 1]) === correoObjetivo) {
      encontrados.push({
        nombreEquipo: String(filas[i][1]).trim(),
        categoria: String(filas[i][3]).trim(),
        codigo: String(filas[i][CODIGO_EQUIPO_COL_ - 1]).trim(),
        passwordCapitan: String(filas[i][CAPITAN_PASSWORD_COL_ - 1]).trim()
      });
    }
  }
  if (!encontrados.length) throw new Error("Ese correo no está dado de alta.");

  encontrados.forEach(function (eq) {
    enviarCorreoBienvenidaCapitan_(data.correo, eq.nombreEquipo, eq.categoria, eq.codigo, eq.passwordCapitan);
  });

  return { enviados: encontrados.length };
}

/* =====================================================================
   PORTAL DE CAPITANES (capitanes.html)
   ===================================================================== */

// Cuota fija de inscripción por equipo. Debe coincidir con
// CONFIG.cuotaEquipo en Admin_Liga_Basquetbol.html.
var CUOTA_EQUIPO_ = 1500;

/**
 * Busca la fila de "Inscripción de Equipos" cuyo correo+password coincidan
 * (comparación de correo sin acentos/mayúsculas, password sin distinguir
 * mayúsculas/minúsculas). Devuelve la fila completa (array) o null.
 */
function buscarFilaCapitan_(hojaInsc, correo, password) {
  var ultima = hojaInsc.getLastRow();
  if (ultima < 2) return null;
  var valores = hojaInsc.getRange(2, 1, ultima - 1, CAPITAN_PASSWORD_COL_).getValues(); // A:M
  var correoObjetivo = normalizarTexto_(correo);
  var passwordObjetivo = String(password || "").trim().toUpperCase();
  for (var i = 0; i < valores.length; i++) {
    var fila = valores[i];
    var correoFila = normalizarTexto_(fila[CORREO_EQUIPO_COL_ - 1]);
    var passFila = String(fila[CAPITAN_PASSWORD_COL_ - 1] || "").trim().toUpperCase();
    if (correoFila && correoFila === correoObjetivo && passFila && passFila === passwordObjetivo) {
      return { fila: fila, indiceFila: i + 2 }; // indiceFila = número de fila real en la hoja
    }
  }
  return null;
}

/**
 * Login del Portal de Capitanes. Valida correo+contraseña y, si son
 * correctos, arma en un solo golpe todo lo que el Portal necesita mostrar:
 * datos del equipo, su roster completo (de "Integrantes") y su resumen de
 * pagos (de "Pagos" contra CUOTA_EQUIPO_). Nunca expone datos de otros
 * equipos ni columnas sensibles de otras filas.
 */
function loginCapitan(ss, data) {
  var hojaInsc = ss.getSheetByName("Inscripción de Equipos");
  if (!hojaInsc) throw new Error("No existe la hoja 'Inscripción de Equipos'.");

  var encontrado = buscarFilaCapitan_(hojaInsc, data.correo, data.password);
  if (!encontrado) throw new Error("Correo o contraseña incorrectos.");
  var fila = encontrado.fila;

  var equipo = String(fila[1]).trim();
  var categoria = String(fila[3]).trim();
  var codigo = String(fila[CODIGO_EQUIPO_COL_ - 1]).trim();

  // --- Roster: de "Integrantes", filtrado por este equipo ---
  var roster = [];
  var hojaInt = ss.getSheetByName("Integrantes");
  if (hojaInt) {
    var ultimaInt = hojaInt.getLastRow();
    if (ultimaInt >= 2) {
      var filasInt = hojaInt.getRange(2, 1, ultimaInt - 1, 13).getValues(); // A:M
      for (var j = 0; j < filasInt.length; j++) {
        var f = filasInt[j];
        if (String(f[2]).trim() === equipo) { // columna C = Equipo
          roster.push({
            nombre: String(f[1]).trim(), // columna B = Nombre Completo
            correo: String(f[4] || "").trim() // columna E = Correo
          });
        }
      }
    }
  }

  // --- Pagos: de "Pagos", filtrado por este equipo ---
  var pagado = 0;
  var historial = [];
  var hojaPagos = ss.getSheetByName("Pagos");
  if (hojaPagos) {
    var ultimaPagos = hojaPagos.getLastRow();
    if (ultimaPagos >= 2) {
      var filasPagos = hojaPagos.getRange(2, 1, ultimaPagos - 1, 10).getValues(); // A:J
      for (var k = 0; k < filasPagos.length; k++) {
        var p = filasPagos[k];
        if (String(p[1]).trim() === equipo) { // columna B = Equipo
          var monto = Number(p[5]) || 0; // columna F = Monto
          pagado += monto;
          historial.push({
            fecha: String(p[0] || ""),
            concepto: String(p[4] || ""),
            monto: monto,
            metodoPago: String(p[6] || ""),
            referencia: String(p[7] || "")
          });
        }
      }
    }
  }

  return {
    equipo: equipo,
    categoria: categoria,
    codigo: codigo,
    roster: roster,
    pagos: {
      cuota: CUOTA_EQUIPO_,
      pagado: pagado,
      saldo: CUOTA_EQUIPO_ - pagado,
      historial: historial
    }
  };
}

/**
 * El capitán cambia su propia contraseña desde adentro del Portal. Valida
 * la contraseña actual con el mismo criterio que loginCapitan antes de
 * sobrescribir la columna M con la nueva.
 */
function cambiarPasswordCapitan(ss, data) {
  var passwordNueva = String(data.passwordNueva || "").trim();
  if (!passwordNueva) throw new Error("Falta la contraseña nueva.");
  if (passwordNueva.length < 4) throw new Error("La contraseña nueva debe tener al menos 4 caracteres.");

  var hojaInsc = ss.getSheetByName("Inscripción de Equipos");
  if (!hojaInsc) throw new Error("No existe la hoja 'Inscripción de Equipos'.");

  var encontrado = buscarFilaCapitan_(hojaInsc, data.correo, data.passwordActual);
  if (!encontrado) throw new Error("Correo o contraseña actual incorrectos.");

  hojaInsc.getRange(encontrado.indiceFila, CAPITAN_PASSWORD_COL_).setValue(passwordNueva.toUpperCase());
}

function programarPartido(ss, data) {
  var hoja = ss.getSheetByName("Rol de Juego");
  if (!hoja) throw new Error("No existe la hoja 'Rol de Juego'.");
  // Esta hoja no tiene fórmulas precargadas, así que appendRow es seguro.
  hoja.appendRow([
    data.jornada, data.fecha, data.hora, data.sede, data.categoria,
    data.equipoLocal, data.equipoVisit, data.arbitro, "Programado"
  ]);
}

/**
 * PASO 5c — agrega una fila nueva a "Pagos". Columnas: Timestamp, Equipo,
 * Categoría, Jugador (opcional), Concepto, Monto, Método de Pago,
 * Referencia/Comprobante, Registrado por / Quién recibió, Notas.
 */
function registrarPago(ss, data) {
  var hoja = ss.getSheetByName("Pagos");
  if (!hoja) throw new Error("No existe la hoja 'Pagos'.");
  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  hoja.appendRow([
    data.fecha ? (data.fecha + " " + ahora.split(" ")[1]) : ahora,
    data.equipo, data.categoria || "", "", data.concepto || "Inscripción Equipo",
    data.monto, data.metodoPago, data.referencia || "", data.quienRecibio || "", data.notas || ""
  ]);
}

/**
 * PASO 5b — corrige el marcador de un partido que ya estaba en "Resultados"
 * (lo encuentra por jornada + equipo local + equipo visitante y sobrescribe
 * columnas D y F). No toca las fórmulas de Ganador/Diferencia.
 */
function editarResultado(ss, data) {
  var hoja = ss.getSheetByName("Resultados");
  if (!hoja) throw new Error("No existe la hoja 'Resultados'.");
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) throw new Error("La hoja 'Resultados' está vacía, no hay nada que corregir.");

  var rango = hoja.getRange(2, 1, lastRow - 1, 6).getValues(); // A:F
  for (var i = 0; i < rango.length; i++) {
    var fila = rango[i];
    if (String(fila[0]) === String(data.jornada) && fila[2] === data.equipoLocal && fila[4] === data.equipoVisit) {
      hoja.getRange(i + 2, 4).setValue(data.ptsLocal); // columna D = Puntos Local
      hoja.getRange(i + 2, 6).setValue(data.ptsVisit); // columna F = Puntos Visitante
      return;
    }
  }
  throw new Error("No se encontró ese partido en 'Resultados' para corregirlo (¿ya lo habías guardado antes?).");
}

/**
 * PASO 5b — corrige la asistencia de un partido ya guardado. Para cada
 * jugador del payload: si ya existe una fila de Asistencia con la misma
 * jornada+equipo+jugador, la actualiza en su lugar; si no existía (por
 * ejemplo, un jugador que se agregó al roster después), la agrega.
 */
function editarAsistencia(ss, data) {
  var hoja = ss.getSheetByName("Asistencia");
  if (!hoja) throw new Error("No existe la hoja 'Asistencia'.");
  if (!data.asistencia || !data.asistencia.length) return;

  var lastRow = hoja.getLastRow();
  var index = {}; // "jornada|equipo|jugador" -> número de fila real en la hoja
  if (lastRow >= 2) {
    var rango = hoja.getRange(2, 1, lastRow - 1, 6).getValues(); // A:F
    for (var i = 0; i < rango.length; i++) {
      var fila = rango[i];
      var key = fila[1] + "|" + fila[3] + "|" + fila[4]; // Jornada|Equipo|Jugador
      index[key] = i + 2;
    }
  }

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  var filasNuevas = [];
  data.asistencia.forEach(function (a) {
    var key = data.jornada + "|" + a.equipo + "|" + a.jugador;
    if (index[key]) {
      hoja.getRange(index[key], 1, 1, 6).setValues([[ahora, data.jornada, data.fecha, a.equipo, a.jugador, a.asistio]]);
    } else {
      filasNuevas.push([ahora, data.jornada, data.fecha, a.equipo, a.jugador, a.asistio]);
    }
  });
  if (filasNuevas.length) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filasNuevas.length, 6).setValues(filasNuevas);
  }
}

/* =====================================================================
   LECTURA DE DATOS (doGet)
   ---------------------------------------------------------------------
   La Consola y el Admin leen los datos desde aquí en vez de publicar
   las hojas en la web. Ventaja: el Web App corre con TU cuenta, así que
   nadie necesita permisos sobre el Sheet, y solo se entrega lo que la
   lista blanca de abajo autoriza — los teléfonos y los montos nunca
   quedan en una URL pública indexable.

   Uso:
     .../exec?hoja=Equipos          -> filas de la hoja "Equipos"
     .../exec?hojas=Equipos,Pagos   -> varias hojas de un solo golpe
     .../exec                       -> mensaje de "activo" (prueba)

   Respuesta:
     {"ok":true,"hojas":{"Equipos":[{"Nombre Equipo":"...", ...}, ...]}}

   El parámetro callback= habilita JSONP, que es como la página esquiva
   las restricciones de CORS de Apps Script al leer.
   ===================================================================== */

// Solo estas hojas se pueden leer. Si no está en la lista, se rechaza.
// "Inscripción de Equipos" NO está aquí a propósito: tiene teléfonos de
// capitanes y los códigos secretos de cada equipo. Esa hoja solo se lee
// por POST con la acción "leerInscripciones" + la contraseña del Admin.
var HOJAS_PERMITIDAS = [
  "Equipos",
  "Integrantes",
  "Resultados",
  "Rol de Juego",
  "Tabla General",
  "Asistencia",
  "Pagos"
];

// En qué columna(s) aparece el nombre del equipo en cada hoja pública. Se usa
// para esconder de doGet() cualquier fila de un equipo de PRUEBA (ver
// esEquipoPrueba_): así ningún dato de prueba se cuela en la Consola pública
// ni en las pestañas normales del Admin (que leen estas mismas hojas).
var CAMPOS_EQUIPO_POR_HOJA_ = {
  "Equipos": ["Nombre Equipo"],
  "Integrantes": ["Equipo"],
  "Resultados": ["Equipo Local", "Equipo Visitante"],
  "Rol de Juego": ["Equipo Local", "Equipo Visitante"],
  "Tabla General": ["Equipo"],
  "Asistencia": ["Equipo"],
  "Pagos": ["Equipo"]
};

/** Quita las filas donde cualquiera de "campos" sea un equipo de PRUEBA. */
function quitarFilasDePrueba_(filas, campos) {
  if (!campos || !campos.length) return filas;
  return filas.filter(function (fila) {
    for (var i = 0; i < campos.length; i++) {
      if (esEquipoPrueba_(fila[campos[i]])) return false;
    }
    return true;
  });
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var callback = params.callback;

  // Sin parámetros: prueba de vida, para poder abrir la URL en el navegador.
  if (!params.hoja && !params.hojas) {
    return ContentService.createTextOutput("Apps Script de la Liga de Basquetbol activo.");
  }

  var salida = { ok: false };
  try {
    var pedidas = String(params.hojas || params.hoja)
      .split(",")
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s !== ""; });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resultado = {};

    pedidas.forEach(function (nombre) {
      if (HOJAS_PERMITIDAS.indexOf(nombre) === -1) {
        throw new Error("La hoja '" + nombre + "' no está autorizada para lectura.");
      }
      var filas = leerHojaComoObjetos_(ss, nombre);
      resultado[nombre] = quitarFilasDePrueba_(filas, CAMPOS_EQUIPO_POR_HOJA_[nombre]);
    });

    salida.ok = true;
    salida.hojas = resultado;
  } catch (err) {
    salida.error = String(err);
  }

  var texto = JSON.stringify(salida);
  if (callback) {
    // JSONP: la página lo carga con <script>, así no hay problema de CORS.
    return ContentService
      .createTextOutput(callback + "(" + texto + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(texto)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Junta TODO lo relacionado con equipos de PRUEBA (nombre "Prueba", ver
 * esEquipoPrueba_), de todas las hojas donde puedan aparecer -- justo lo
 * opuesto al filtro que aplican doGet() y leerInscripciones(). Es para que
 * el Admin tenga dónde revisar una prueba completa (inscripción, roster,
 * pagos, partidos) sin que se mezcle con los equipos reales en ningún otro
 * lado. Protegida con ADMIN_CLAVE_ igual que leerInscripciones, porque
 * "Inscripción de Equipos" trae teléfonos y códigos de equipo.
 *
 * Acción: "leerPruebas"
 * Payload: { action: "leerPruebas", clave: "..." }
 * Respuesta: { ok:true, hojas: { "Inscripción de Equipos":[...], "Equipos":[...],
 *   "Integrantes":[...], "Resultados":[...], "Rol de Juego":[...],
 *   "Asistencia":[...], "Pagos":[...], "Tabla General":[...] } }
 */
function leerPruebas(ss, data) {
  if (String(data.clave || "") !== ADMIN_CLAVE_) {
    throw new Error("No autorizado.");
  }

  var salida = {};

  var inscripciones = leerHojaComoObjetos_(ss, "Inscripción de Equipos");
  salida["Inscripción de Equipos"] = inscripciones.filter(function (f) {
    return esEquipoPrueba_(f["Nombre del Equipo"]);
  });

  Object.keys(CAMPOS_EQUIPO_POR_HOJA_).forEach(function (nombreHoja) {
    var campos = CAMPOS_EQUIPO_POR_HOJA_[nombreHoja];
    var filas = leerHojaComoObjetos_(ss, nombreHoja);
    salida[nombreHoja] = filas.filter(function (fila) {
      return campos.some(function (campo) { return esEquipoPrueba_(fila[campo]); });
    });
  });

  return salida;
}

/**
 * Convierte una hoja en un arreglo de objetos usando la fila 1 como
 * encabezados. Se saltan las filas totalmente vacías (las hojas traen
 * cientos de filas precargadas con fórmulas que devuelven "").
 */
function leerHojaComoObjetos_(ss, nombre) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) throw new Error("No existe la hoja '" + nombre + "'.");

  var ultimaFila = hoja.getLastRow();
  var ultimaCol = hoja.getLastColumn();
  if (ultimaFila < 2 || ultimaCol < 1) return [];

  var valores = hoja.getRange(1, 1, ultimaFila, ultimaCol).getDisplayValues();
  var encabezados = valores[0].map(function (h) { return String(h).trim(); });

  var filas = [];
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    var vacia = fila.every(function (c) { return String(c).trim() === ""; });
    if (vacia) continue;

    var obj = {};
    for (var j = 0; j < encabezados.length; j++) {
      if (encabezados[j] === "") continue;
      obj[encabezados[j]] = fila[j];
    }
    filas.push(obj);
  }
  return filas;
}

function escribirResultado(ss, data) {
  var hoja = ss.getSheetByName("Resultados");
  if (!hoja) throw new Error("No existe la hoja 'Resultados'.");

  // La hoja ya trae fórmulas de Ganador/Diferencia precargadas en todas las
  // filas (columnas G y H), así que buscamos la primera fila VACÍA en la
  // columna A (Jornada) en vez de usar appendRow (que se iría hasta el
  // final de las fórmulas y rompería el diseño).
  var lastRow = hoja.getMaxRows();
  var colA = hoja.getRange(2, 1, lastRow - 1, 1).getValues();
  var filaLibre = -1;
  for (var i = 0; i < colA.length; i++) {
    if (colA[i][0] === "" || colA[i][0] === null) {
      filaLibre = i + 2;
      break;
    }
  }
  if (filaLibre === -1) throw new Error("No hay filas libres en 'Resultados'. Avísale a quien armó el archivo para ampliar el rango.");

  hoja.getRange(filaLibre, 1, 1, 6).setValues([[
    data.jornada, data.fecha, data.equipoLocal, data.ptsLocal, data.equipoVisit, data.ptsVisit
  ]]);
  hoja.getRange(filaLibre, 9).setValue(data.sede || ""); // columna I = Observaciones
}

function marcarRolComoJugado(ss, data) {
  var hoja = ss.getSheetByName("Rol de Juego");
  if (!hoja) return;
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return;
  var rango = hoja.getRange(2, 1, lastRow - 1, 9).getValues(); // A:I
  for (var i = 0; i < rango.length; i++) {
    var fila = rango[i];
    if (String(fila[0]) === String(data.jornada) && fila[5] === data.equipoLocal && fila[6] === data.equipoVisit) {
      hoja.getRange(i + 2, 9).setValue("Jugado"); // columna I = Estatus
      break;
    }
  }
}

function escribirAsistencia(ss, data) {
  var hoja = ss.getSheetByName("Asistencia");
  if (!hoja) throw new Error("No existe la hoja 'Asistencia'.");
  if (!data.asistencia || !data.asistencia.length) return;

  var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  var filas = data.asistencia.map(function (a) {
    return [ahora, data.jornada, data.fecha, a.equipo, a.jugador, a.asistio];
  });
  // Asistencia no tiene fórmulas precargadas, así que appendRow normal está bien.
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
}
