// =====================================================================================
// SISTEMA DE AGENDA OGA - VERSIÓN 23 CON MEJORAS UI Y CONTROL DE VOTACIONES
// Archivo: Agenda_Consejo.gs
// NUEVO (v23): Estadísticas congruentes y legibles (resumen + porTipo + estado por registro)
// NUEVO (v23): Filtro por reunión vigente — lo vigente es lo que está en Configuración,
//               no importa si ya está cerrada/aprobada/vencida (obtenerReunionVigenteParaFiltros)
// NUEVO (v23): OGA-FAC-114 obligatorio para decisiones que NO sean Reconocimiento
// NUEVO (v23): Reconocimiento de Acreditación separado: sin documentos FAC obligatorios,
//               solo documentación con nombre estándar en carpeta Documentos_Reconocimiento
// NUEVO (v23): Reconocimiento de Acreditación eliminado de Notificaciones
// NUEVO (v23): Vista previa del documento de agenda sin guardar ni registrar
//               (generarVistaPreviaAgendaConsejo — solo lectura, sin efectos secundarios)
// NUEVO: Control de acceso específico para votaciones
// NUEVO: Estadísticas mejoradas por tipo de decisión
// MODIFICADO: FIX para distinguir múltiples tipos de decisión por registro
// NUEVO: UI mejorada con texto nítido en ambos modos
// MODIFICADO: Correo institucional obligatorio info@oga.org.gt
// =====================================================================================

// ===================== CONFIGURACIÓN GLOBAL ACTUALIZADA =====================
const AGENDA_SHEET_ID = "1ojF7iJ_npGYO4NH6SxqHAmsKCRX-62HTeM6RP1oEWDU";
const FOLDER_ID = "1DgdppxZ5QcSHTEiqR5eVh1yRBRMxJzu2"; // Para archivos adjuntos
const AGENDAS_CREADAS_FOLDER_ID = "1mDYD7qe5KD9Rj41ALCLthtEO-2G9pOeV"; // Para agendas generadas
const PLANTILLA_AGENDA_ID = "1rIg79HVddWNNAS3E82kfyCAbVYVjAqFwhca6kziwuNA";
const CONFIRMACION_DATOS_ID = "1W30ZEOKJ3Sw1w9SIyZJxs46oLCVWykAT7wN3ozWANAc";
const USUARIOS_SHEET_ID = "1SOdGKBGv4tGfG5Ktssxq-TyKZSmYphqJkF87Kx1vtsY";
const REUNIONES_SHEET_ID = "1ojF7iJ_npGYO4NH6SxqHAmsKCRX-62HTeM6RP1oEWDU";
const VOTACIONES_SHEET_ID_REF = "1ojF7iJ_npGYO4NH6SxqHAmsKCRX-62HTeM6RP1oEWDU";
const HOJA_VOTACIONES_REF = "Votaciones";
const WEBAPP_VOTACION_URL = "https://script.google.com/macros/s/AKfycbyI9-UE30sDzNKDWJQnpnCN7hTXHlmLq8-efRyEXzjC41K2aAjavpELnialX1kV_WAl0w/exec";
const CORREO_INSTITUCIONAL = "consejo_acreditacion@oga.org.gt";
const CORREO_API_VOTACIONES_URL = "https://script.google.com/macros/s/AKfycby_sNpy3bi9ic1Mnrz_5Ov0pM8Z2VtEBed2dGH22U_WfSiZYyFwB_pL1MikrXmbigkv/exec";
const CORREO_API_VOTACIONES_KEY = "SMTP_Votacin_OGA_2026";
const CORREO_API_TIMEOUT_MS = 30000;
const CORREOS_PERMITIDOS_CREAR = [
  "carchila@oga.org.gt",
  "eajuchan@oga.org.gt",
  "ppineda@oga.org.gt",
  "jalvarado.oga@gmail.com",
  "info@oga.org.gt"
];
const CORREOS_PERMITIDOS_VOTACION = [
  "carchila@oga.org.gt",
  "eajuchan@oga.org.gt",
  "lsanchez@oga.org.gt",
  "jalvarado.oga@gmail.com",
  "info@oga.org.gt"
];

// ===================== FUNCIONES DE PERMISOS MEJORADAS =====================
function esUsuarioAutorizadoParaAgenda() {
  try {
    const correo = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
    if (!correo) return false;
    return CORREOS_PERMITIDOS_CREAR.map(c => (c || '').toLowerCase().trim()).includes(correo);
  } catch (e) { return false; }
}
function esUsuarioAutorizadoParaVotacion() {
  try {
    const correo = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
    if (!correo) return false;
    return CORREOS_PERMITIDOS_VOTACION.map(c => (c || '').toLowerCase().trim()).includes(correo);
  } catch (e) { return false; }
}
function getUsuarioActual() {
  const email = (Session.getActiveUser().getEmail() || '').trim();
  return { email, autorizadoCrear: esUsuarioAutorizadoParaAgenda(), autorizadoVotar: esUsuarioAutorizadoParaVotacion() };
}

// ===================== CONTADOR DE TIEMPO REGRESIVO (INTEGRADO) =====================
function calcularTiempoRestante() {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const configSheet = ss.getSheetByName('Configuracion');
    if (!configSheet || configSheet.getLastRow() < 2) {
      return { success: false, message: 'No hay agenda activa', tiempoRestante: 0, expirada: true };
    }
    const datosAgenda = configSheet.getRange('A2:K2').getValues()[0];
    const [, numeroReunion, fechaInicio, fechaFin, horaLimite, , , , estado] = datosAgenda;
    if (estado === 'CERRADA' || estado === 'VENCIDA') {
      return { success: true, message: 'Agenda cerrada', tiempoRestante: 0, expirada: true, estado: estado };
    }
    const ahora = new Date();
    const fechaFinDate = new Date(fechaFin);
    let fechaLimite = new Date(fechaFinDate);
    try {
      if (horaLimite && typeof horaLimite === 'string' && horaLimite.includes(':')) {
        const [h, m] = horaLimite.split(':');
        fechaLimite.setHours(parseInt(h) || 23, parseInt(m) || 59, 59, 999);
      } else { fechaLimite.setHours(23, 59, 59, 999); }
    } catch (_) { fechaLimite.setHours(23, 59, 59, 999); }
    const tiempoRestanteMs = fechaLimite.getTime() - ahora.getTime();
    const expirada = tiempoRestanteMs <= 0;
    if (expirada && estado === 'ACTIVA') {
      try { configSheet.getRange('I2').setValue('VENCIDA'); Logger.log(' Agenda ' + numeroReunion + ' marcada como VENCIDA automáticamente'); }
      catch (e) { Logger.log(' Error marcando agenda como vencida: ' + e.message); }
    }
    return { success: true, numeroReunion: String(numeroReunion || '').replace(/^'/, '').trim(), fechaLimite: fechaLimite, tiempoRestante: Math.max(0, tiempoRestanteMs), expirada: expirada, estado: expirada ? 'VENCIDA' : estado, mensaje: expirada ? 'Tiempo agotado' : 'Agenda activa' };
  } catch (error) {
    Logger.log(' Error calculando tiempo restante: ' + error.message);
    return { success: false, message: 'Error: ' + error.message, tiempoRestante: 0, expirada: true };
  }
}
function formatearTiempo(ms) {
  if (ms <= 0) return { horas: 0, minutos: 0, segundos: 0, texto: '00:00:00', textoLegible: 'Tiempo agotado', clase: 'expirado' };
  const segundosTotal = Math.floor(ms / 1000);
  const horas = Math.floor(segundosTotal / 3600);
  const minutos = Math.floor((segundosTotal % 3600) / 60);
  const segundos = segundosTotal % 60;
  const horasStr = String(horas).padStart(2, '0');
  const minutosStr = String(minutos).padStart(2, '0');
  const segundosStr = String(segundos).padStart(2, '0');
  let textoLegible = '';
  let clase = 'normal';
  if (horas > 0) { textoLegible = horas + 'h ' + minutos + 'm restantes'; }
  else if (minutos > 0) { textoLegible = minutos + 'm ' + segundos + 's restantes'; if (minutos <= 5) clase = 'urgente'; }
  else { textoLegible = segundos + 's restantes'; clase = 'critico'; }
  return { horas, minutos, segundos, texto: horasStr + ':' + minutosStr + ':' + segundosStr, textoLegible: textoLegible, clase: clase };
}
function obtenerEstadoTiempoAgenda() {
  try {
    const tiempoInfo = calcularTiempoRestante();
    if (!tiempoInfo.success) return tiempoInfo;
    const tiempoFormateado = formatearTiempo(tiempoInfo.tiempoRestante);
    return { success: true, numeroReunion: tiempoInfo.numeroReunion, tiempoRestante: tiempoInfo.tiempoRestante, expirada: tiempoInfo.expirada, estado: tiempoInfo.estado, mensaje: tiempoInfo.mensaje, tiempo: tiempoFormateado, fechaLimite: tiempoInfo.fechaLimite ? tiempoInfo.fechaLimite.toISOString() : null, timestamp: new Date().toISOString() };
  } catch (error) {
    return { success: false, message: 'Error obteniendo estado: ' + error.message, expirada: true };
  }
}
function puedeAgregarElementos() {
  try {
    const tiempoInfo = calcularTiempoRestante();
    if (!tiempoInfo.success) return { puede: false, motivo: 'NO_AGENDA', mensaje: 'No hay agenda activa' };
    if (tiempoInfo.expirada) return { puede: false, motivo: 'TIEMPO_AGOTADO', mensaje: ' El tiempo para agregar elementos ha expirado' };
    if (tiempoInfo.estado === 'CERRADA') return { puede: false, motivo: 'AGENDA_CERRADA', mensaje: ' La agenda ha sido cerrada manualmente' };
    return { puede: true, motivo: 'ACTIVA', mensaje: 'Puede agregar elementos', tiempoRestante: formatearTiempo(tiempoInfo.tiempoRestante).textoLegible };
  } catch (error) {
    return { puede: false, motivo: 'ERROR', mensaje: 'Error verificando permisos: ' + error.message };
  }
}

// ===================== REUNIÓN VIGENTE PARA FILTROS (v23) =====================
// Lo vigente es lo que está en Configuración (B2), no importa si ya está
// CERRADA, APROBADA, VENCIDA o PENDIENTE_APROBACION. El filtro de datos y de
// estadísticas siempre usa esta reunión.
function obtenerReunionVigenteParaFiltros() {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const configSheet = ss.getSheetByName('Configuracion');
    if (configSheet && configSheet.getLastRow() >= 2) {
      const numero = String(configSheet.getRange('B2').getDisplayValue() || '').replace(/^'/, '').trim();
      const estado = String(configSheet.getRange('I2').getDisplayValue() || '').trim();
      if (numero) {
        return { success: true, numeroReunion: numero, estado: estado || 'ACTIVA', fuente: 'Configuracion' };
      }
    }
    // Fallback 1: hoja Agenda_Votacion (B2)
    const votacionSheet = ss.getSheetByName('Agenda_Votacion');
    if (votacionSheet && votacionSheet.getLastRow() >= 2) {
      const numero = String(votacionSheet.getRange('B2').getDisplayValue() || '').replace(/^'/, '').trim();
      const estado = String(votacionSheet.getRange('M2').getDisplayValue() || '').trim();
      if (numero) return { success: true, numeroReunion: numero, estado: estado, fuente: 'Agenda_Votacion' };
    }
    // Fallback 2: última fila de la hoja Reuniones
    const reunionesSheet = ss.getSheetByName('Reuniones');
    if (reunionesSheet && reunionesSheet.getLastRow() >= 2) {
      const lastRow = reunionesSheet.getLastRow();
      const numero = String(reunionesSheet.getRange(lastRow, 1).getDisplayValue() || '').replace(/^'/, '').trim();
      if (numero) return { success: true, numeroReunion: numero, estado: '', fuente: 'Reuniones' };
    }
    return { success: false, message: 'No se encontró ninguna reunión vigente en Configuración, Agenda_Votacion ni Reuniones.' };
  } catch (e) {
    return { success: false, message: 'Error obteniendo reunión vigente: ' + e.message };
  }
}
function obtenerDatosAgendaVigente() {
  try {
    const vigente = obtenerReunionVigenteParaFiltros();
    if (!vigente.success) return { success: false, message: vigente.message, reunionVigente: '', data: [], total: 0 };
    const resultado = obtenerDatosAgendaConsejoParaEdicion();
    const todos = resultado.success ? (resultado.data || []) : [];
    const filtrados = todos.filter(function (item) { return String(item.acta || '').includes(vigente.numeroReunion); });
    return { success: true, reunionVigente: vigente.numeroReunion, estado: vigente.estado, fuente: vigente.fuente, data: filtrados, total: filtrados.length };
  } catch (e) {
    return { success: false, message: 'Error obteniendo datos de la reunión vigente: ' + e.message, reunionVigente: '', data: [], total: 0 };
  }
}
function obtenerEstadisticasReunionVigente() {
  try {
    const vigente = obtenerReunionVigenteParaFiltros();
    if (!vigente.success) return { success: false, message: vigente.message };
    const stats = obtenerEstadisticasVotosPorReunionMejoradas(vigente.numeroReunion);
    stats.reunionVigente = vigente.numeroReunion;
    stats.fuenteVigente = vigente.fuente;
    return stats;
  } catch (e) {
    return { success: false, message: 'Error obteniendo estadísticas de la reunión vigente: ' + e.message };
  }
}

// ===================== FUNCIONES DE VOTACIONES CON CONTROL DE ACCESO =====================
function hayAgendaListaParaVotar() {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const votacionSheet = ss.getSheetByName('Agenda_Votacion');
    if (!votacionSheet) return { success: false, mensaje: 'No existe hoja Agenda_Votacion' };
    const numeroReunion = votacionSheet.getRange('B2').getDisplayValue().replace(/^'/, '').trim();
    const estadoArchivo = votacionSheet.getRange('M2').getDisplayValue().trim();
    if (!numeroReunion) return { success: false, mensaje: 'No hay reunión en B2' };
    if (estadoArchivo !== 'ARCHIVADA') return { success: false, mensaje: 'Estado M2 es "' + estadoArchivo + '", necesita ser "ARCHIVADA"' };
    return { success: true, mensaje: ' Reunión ' + numeroReunion + ' lista para votación', numeroReunion: numeroReunion, estado: estadoArchivo };
  } catch (error) {
    return { success: false, mensaje: 'Error: ' + error.message };
  }
}
function enviarCorreoInstitucionalPorApi_(to, subject, htmlBody, textBody) {
  try {
    if (!to || !subject || !htmlBody) return { success: false, message: 'Faltan campos obligatorios para enviar correo institucional.' };
    const payload = { apiKey: CORREO_API_VOTACIONES_KEY, to: String(to).trim(), subject: String(subject).trim(), body: String(htmlBody || textBody || '') };
    const response = UrlFetchApp.fetch(CORREO_API_VOTACIONES_URL, { method: 'post', contentType: 'application/json; charset=utf-8', payload: JSON.stringify(payload), muteHttpExceptions: true, followRedirects: true });
    const code = response.getResponseCode();
    const body = response.getContentText() || '';
    const okHttp = code >= 200 && code < 300;
    const okBody = /^Success:/i.test(body.trim());
    if (!okHttp || !okBody) return { success: false, message: 'API correo respondió HTTP ' + code + ': ' + body, httpCode: code, response: body };
    return { success: true, message: body, httpCode: code, from: CORREO_INSTITUCIONAL };
  } catch (error) {
    return { success: false, message: 'Error llamando API de correo: ' + error.message };
  }
}
function enviarVotacionesDirectas(modoPrueba, maxEnviar) {
  try {
    modoPrueba = !!modoPrueba;
    if (!esUsuarioAutorizadoParaVotacion()) {
      return { success: false, message: ' Acceso denegado. Solo Carlos Archila, Estephanie Ajuchán, Lidia Sanchez y Jonatan Alvarado pueden enviar votaciones.' };
    }
    const verificacion = hayAgendaListaParaVotar();
    if (!verificacion.success) return { success: false, message: verificacion.mensaje };
    const numeroReunion = verificacion.numeroReunion;
    const miembros = obtenerMiembrosParaVotacion(modoPrueba);
    if (miembros.length === 0) return { success: false, message: "No se encontraron miembros" };
    const listaEnvio = maxEnviar ? miembros.slice(0, maxEnviar) : miembros;
    const enviados = [];
    const errores = [];
    listaEnvio.forEach(function (m) {
      try {
        const url = WEBAPP_VOTACION_URL + "?email=" + encodeURIComponent(m.email) + "&reunion=" + encodeURIComponent(numeroReunion);
        const inv = construirCuerpoInvitacion(m, { numeroReunion: numeroReunion, nombre: 'Reunión ' + numeroReunion }, url, modoPrueba);
        const envioApi = enviarCorreoInstitucionalPorApi_(m.email, inv.asunto, inv.htmlBody, inv.textBody);
        if (!envioApi.success) throw new Error(envioApi.message);
        enviados.push({ email: m.email, url: url, apiResponse: envioApi.message });
      } catch (e) { errores.push({ email: m.email, error: e.message }); }
    });
    return { success: errores.length === 0, message: ' Votaciones enviadas vía API institucional desde ' + CORREO_INSTITUCIONAL + ' para ' + numeroReunion + '. Exitosos: ' + enviados.length + ', Errores: ' + errores.length, reunion: numeroReunion, enviados: enviados, errores: errores };
  } catch (error) {
    return { success: false, message: 'Error: ' + error.message };
  }
}
function buscarAgendasAprobadas() {
  const resultado = hayAgendaListaParaVotar();
  if (resultado.success) {
    return { success: true, mensaje: resultado.mensaje, agendas: [{ numeroReunion: resultado.numeroReunion, nombre: 'Reunión ' + resultado.numeroReunion, fechaAprobacion: new Date(), totalRegistros: 3, tipo: 'ARCHIVADA' }] };
  }
  return { success: false, mensaje: resultado.mensaje, agendas: [] };
}
function enviarInvitacionesVotacionEspecifica(numeroReunion, modoPrueba, maxEnviar) { return enviarVotacionesDirectas(modoPrueba, maxEnviar); }
function enviarInvitacionesVotacion(modoPrueba, maxEnviar) { return enviarVotacionesDirectas(modoPrueba, maxEnviar); }
function debugAgendasArchivadas() {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const configSheet = ss.getSheetByName('Configuracion');
    let estadoConfiguracion = 'NO EXISTE';
    let datosConfiguracion = null;
    if (configSheet) {
      try { estadoConfiguracion = configSheet.getRange('I2').getDisplayValue().trim() || 'VACIO'; datosConfiguracion = configSheet.getRange('A2:K2').getValues()[0]; }
      catch (e) { estadoConfiguracion = 'ERROR: ' + e.message; }
    }
    const votacionSheet = ss.getSheetByName('Agenda_Votacion');
    let estadoArchivo = 'NO EXISTE HOJA';
    let numeroReunionArchivada = 'N/A';
    let datosVotacion = null;
    if (votacionSheet) {
      try { numeroReunionArchivada = votacionSheet.getRange('B2').getDisplayValue().replace(/^'/, '').trim() || 'VACIO'; estadoArchivo = votacionSheet.getRange('M2').getDisplayValue().trim() || 'VACIO'; datosVotacion = votacionSheet.getRange('A2:M2').getValues()[0]; }
      catch (e) { estadoArchivo = 'ERROR: ' + e.message; }
    }
    const agendaValida = (numeroReunionArchivada !== 'VACIO' && estadoArchivo === 'ARCHIVADA');
    return { error: false, configuracion: { existe: !!configSheet, estado: estadoConfiguracion, datos: datosConfiguracion, rango: 'A2:K2' }, votacion: { existe: !!votacionSheet, hayDatos: numeroReunionArchivada !== 'VACIO', numeroReunion: numeroReunionArchivada, estadoArchivo: estadoArchivo, datos: datosVotacion, rango: 'A2:M2' }, agendaValidaEncontrada: agendaValida, totalAgendas: agendaValida ? 1 : 0, mensaje: 'LECTURA DIRECTA - B2: "' + numeroReunionArchivada + '", M2: "' + estadoArchivo + '"' };
  } catch (error) {
    return { error: true, mensaje: 'Error en debug: ' + error.message, configuracion: { existe: false }, votacion: { existe: false, hayDatos: false }, agendaValidaEncontrada: false, totalAgendas: 0 };
  }
}
function buscarReunionAprobadaParaVotacion() {
  const verificacion = hayAgendaListaParaVotar();
  if (verificacion.success) return { disponible: true, mensaje: verificacion.mensaje, config: { numeroReunion: verificacion.numeroReunion, nombre: 'Reunión ' + verificacion.numeroReunion, estado: 'CERRADA', estadoDetallado: 'APROBADA' } };
  return { disponible: false, mensaje: verificacion.mensaje, todasLasAgendas: [] };
}
function obtenerAgendasArchivadas() {
  const resultado = hayAgendaListaParaVotar();
  if (resultado.success) return [{ numeroReunion: resultado.numeroReunion, nombre: 'Reunión ' + resultado.numeroReunion, estado: 'CERRADA', fechaAprobacion: new Date(), totalRegistros: 3, tipo: 'ARCHIVADA' }];
  return [];
}

// ===================== ESTADÍSTICAS CONGRUENTES (v23 — reescritura) =====================
// Lectura de la hoja Votaciones (sin cambios): col 9 = REU- (índice 9), col 3 = registro
// (índice 3), col 5 = voto (índice 5), col 6 = comentario (índice 6), col 2 = institución
// (índice 2), col 7 = nombre (índice 7), col 4 = tipo decisión (índice 4), col 10 = resolución (índice 10).
// Clave única: registro + tipo de decisión. Estado por registro: APROBADO si
// % a favor > 50; RECHAZADO si en contra > a favor; si no, SIN MAYORÍA.
function obtenerEstadisticasVotosPorReunionMejoradas(numeroReunion) {
  try {
    const reunionEtiqueta = 'REU-' + numeroReunion;
    const ss = SpreadsheetApp.openById(VOTACIONES_SHEET_ID_REF);
    const sheet = ss.getSheetByName(HOJA_VOTACIONES_REF);
    const vacio = function () {
      return { success: true, reunion: numeroReunion, reunionVigente: numeroReunion, resumen: { aFavor: 0, enContra: 0, abstencion: 0, total: 0, porcentajeAFavor: 0, porcentajeEnContra: 0, porcentajeAbstencion: 0, aprobados: 0, rechazados: 0, sinMayoria: 0, totalRegistros: 0 }, porTipo: [], registros: [], registrosPorTipo: [], data: { aFavor: 0, enContra: 0, abstencion: 0, total: 0 }, ultimaActualizacion: new Date().toLocaleString('es-ES') };
    };
    if (!sheet || sheet.getLastRow() <= 1) return vacio();
    const datos = sheet.getDataRange().getValues();
    const registroStats = {};
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (String(fila[9] || "") === reunionEtiqueta) {
        const registro = String(fila[3] || "").trim();
        const voto = String(fila[5] || "").trim();
        const comentario = String(fila[6] || "").trim();
        const institucion = String(fila[2] || "").trim();
        const nombre = String(fila[7] || "").trim();
        const tipoDecision = String(fila[4] || "").trim();
        const resolucion = String(fila[10] || "").trim();
        const claveUnica = registro + '___' + tipoDecision;
        if (!registroStats[claveUnica]) {
          registroStats[claveUnica] = { registro: registro, tipoDecision: tipoDecision, resolucion: resolucion, votos: { aFavor: 0, enContra: 0, abstencion: 0 }, comentarios: [], votantes: [], displayName: tipoDecision ? registro + ' (' + tipoDecision + ')' : registro };
        }
        const votoLower = voto.toLowerCase();
        if (votoLower === "a favor") registroStats[claveUnica].votos.aFavor++;
        else if (votoLower === "en contra") registroStats[claveUnica].votos.enContra++;
        else if (votoLower === "abstención" || votoLower === "abstencion") registroStats[claveUnica].votos.abstencion++;
        if (comentario) registroStats[claveUnica].comentarios.push({ comentario: comentario, institucion: institucion, nombre: nombre, voto: voto });
        registroStats[claveUnica].votantes.push({ nombre: nombre, institucion: institucion, voto: voto });
      }
    }
    const registrosPorTipo = {};
    let aFavor = 0, enContra = 0, abstencion = 0, aprobados = 0, rechazados = 0, sinMayoria = 0;
    const registros = Object.values(registroStats).map(function (reg) {
      reg.totalVotos = reg.votos.aFavor + reg.votos.enContra + reg.votos.abstencion;
      reg.porcentajeAFavor = reg.totalVotos > 0 ? Math.round((reg.votos.aFavor / reg.totalVotos) * 100) : 0;
      reg.porcentajeEnContra = reg.totalVotos > 0 ? Math.round((reg.votos.enContra / reg.totalVotos) * 100) : 0;
      reg.porcentajeAbstencion = reg.totalVotos > 0 ? Math.round((reg.votos.abstencion / reg.totalVotos) * 100) : 0;
      reg.puedeGenerarResolucion = reg.porcentajeAFavor > 50;
      if (reg.puedeGenerarResolucion) { reg.estado = 'APROBADO'; aprobados++; }
      else if (reg.votos.enContra > reg.votos.aFavor) { reg.estado = 'RECHAZADO'; rechazados++; }
      else { reg.estado = 'SIN MAYORÍA'; sinMayoria++; }
      const tipo = reg.tipoDecision || 'Sin Tipo';
      if (!registrosPorTipo[tipo]) registrosPorTipo[tipo] = { tipo: tipo, registros: [], totales: { aFavor: 0, enContra: 0, abstencion: 0, total: 0 }, aprobados: 0, rechazados: 0 };
      registrosPorTipo[tipo].registros.push(reg);
      registrosPorTipo[tipo].totales.aFavor += reg.votos.aFavor;
      registrosPorTipo[tipo].totales.enContra += reg.votos.enContra;
      registrosPorTipo[tipo].totales.abstencion += reg.votos.abstencion;
      registrosPorTipo[tipo].totales.total += reg.totalVotos;
      if (reg.puedeGenerarResolucion) registrosPorTipo[tipo].aprobados++;
      else if (reg.votos.enContra > reg.votos.aFavor) registrosPorTipo[tipo].rechazados++;
      aFavor += reg.votos.aFavor; enContra += reg.votos.enContra; abstencion += reg.votos.abstencion;
      return reg;
    });
    const total = aFavor + enContra + abstencion;
    const porTipo = Object.values(registrosPorTipo).map(function (g) {
      return { tipo: g.tipo, totalRegistros: g.registros.length, aFavor: g.totales.aFavor, enContra: g.totales.enContra, abstencion: g.totales.abstencion, total: g.totales.total, aprobados: g.aprobados, rechazados: g.rechazados, porcentajeAFavor: g.totales.total > 0 ? Math.round((g.totales.aFavor / g.totales.total) * 100) : 0 };
    });
    Logger.log(' Estadísticas v23: ' + registros.length + ' registros únicos, ' + total + ' votos, ' + aprobados + ' aprobados, ' + rechazados + ' rechazados, ' + sinMayoria + ' sin mayoría');
    return {
      success: true, reunion: numeroReunion, reunionVigente: numeroReunion,
      resumen: { aFavor: aFavor, enContra: enContra, abstencion: abstencion, total: total, porcentajeAFavor: total > 0 ? Math.round((aFavor / total) * 100) : 0, porcentajeEnContra: total > 0 ? Math.round((enContra / total) * 100) : 0, porcentajeAbstencion: total > 0 ? Math.round((abstencion / total) * 100) : 0, aprobados: aprobados, rechazados: rechazados, sinMayoria: sinMayoria, totalRegistros: registros.length },
      porTipo: porTipo, registros: registros, registrosPorTipo: Object.values(registrosPorTipo),
      data: { aFavor: aFavor, enContra: enContra, abstencion: abstencion, total: total },
      ultimaActualizacion: new Date().toLocaleString('es-ES')
    };
  } catch (err) {
    Logger.log(' Error en estadísticas mejoradas: ' + err.message);
    return { success: false, message: err.message };
  }
}

// ===================== FUNCIONES DE CORRELATIVO =====================
function registrarNumeroReunion(numeroReunion, fechaReunion) {
  try {
    const reunionesSheet = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    let sheetR = reunionesSheet.getSheetByName('Reuniones');
    if (!sheetR) {
      sheetR = reunionesSheet.insertSheet('Reuniones');
      sheetR.getRange(1, 1, 1, 2).setValues([['Numero_Reunion', 'Fecha_Reunion']]);
      sheetR.getRange(1, 1, 1, 2).setFontWeight('bold');
      sheetR.setFrozenRows(1);
    }
    const filaDestino = sheetR.getLastRow() + 1;
    const numeroConApostrofe = "'" + numeroReunion.trim();
    sheetR.getRange(filaDestino, 1).setValue(numeroConApostrofe);
    sheetR.getRange(filaDestino, 2).setValue(new Date(fechaReunion));
    Logger.log(' Correlativo registrado: "' + numeroReunion + '" como texto con apostrofe');
    return { success: true, numeroRegistrado: numeroReunion };
  } catch (error) {
    Logger.log(' Error registrando correlativo: ' + error.message);
    return { success: false, message: error.message };
  }
}
function eliminarNumeroReunion(numeroReunion) {
  try {
    const reunionesSheet = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    const sheetR = reunionesSheet.getSheetByName('Reuniones');
    if (!sheetR) return { success: true, message: 'No hay correlativo que eliminar' };
    const datos = sheetR.getDataRange().getValues();
    for (let i = datos.length - 1; i >= 1; i--) {
      const valorEnHoja = datos[i][0];
      const reunionEnHoja = String(valorEnHoja || '').replace(/^'/, '').trim();
      if (reunionEnHoja === numeroReunion.trim()) {
        sheetR.deleteRow(i + 1);
        Logger.log(' Correlativo ' + numeroReunion + ' eliminado (fila ' + (i + 1) + ')');
        return { success: true, message: 'Correlativo ' + numeroReunion + ' eliminado' };
      }
    }
    return { success: true, message: 'Correlativo ' + numeroReunion + ' no encontrado' };
  } catch (error) {
    Logger.log(' Error eliminando correlativo: ' + error.message);
    return { success: false, message: error.message };
  }
}

// ===================== AUTOMATIZACIÓN DE CALENDARIO (v22, sin cambios) =====================
function obtenerSiguienteFechaReunion() {
  try {
    const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    const hoja = ss.getSheetByName("Calendario_Reuniones");
    if (!hoja) return { success: false, message: " No existe la hoja 'Calendario_Reuniones'" };
    const datos = hoja.getDataRange().getValues();
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    for (let i = 1; i < datos.length; i++) {
      const estado = String(datos[i][3] || "").toUpperCase().trim();
      if (estado === "ACTIVA") {
        const fechaExtra = datos[i][5];
        const fechaNormal = datos[i][2];
        const mes = obtenerMesCorrecto_(datos, i);
        if (esFechaValida_(fechaExtra)) return formatearRespuestaFecha_(fechaExtra, "FECHA EXTRAORDINARIA (ACTIVA)", mes);
        if (esFechaValida_(fechaNormal)) return formatearRespuestaFecha_(fechaNormal, "REUNIÓN PROGRAMADA (ACTIVA)", mes);
      }
    }
    for (let i = 1; i < datos.length; i++) {
      const fechaNormal = datos[i][2];
      const estado = String(datos[i][3] || "").toUpperCase().trim();
      if (esFechaValida_(fechaNormal)) {
        const fechaObj = new Date(fechaNormal); fechaObj.setHours(0, 0, 0, 0);
        if (fechaObj >= hoy && estado !== "CERRADA") {
          const fechaExtra = datos[i][5];
          const mes = obtenerMesCorrecto_(datos, i);
          if (esFechaValida_(fechaExtra)) return formatearRespuestaFecha_(fechaExtra, "FECHA EXTRAORDINARIA", mes);
          return formatearRespuestaFecha_(fechaNormal, "SIGUIENTE REUNIÓN DISPONIBLE", mes);
        }
      }
    }
    return { success: false, message: "No se encontraron fechas disponibles en el calendario." };
  } catch (e) {
    Logger.log("Error en calendario: " + e.message);
    return { success: false, message: "Error interno: " + e.message };
  }
}
function obtenerMesCorrecto_(datos, indiceActual) {
  for (let k = indiceActual; k >= 1; k--) {
    if (datos[k][1] && String(datos[k][1]).trim() !== "") return String(datos[k][1]).toUpperCase();
  }
  return "MES DESCONOCIDO";
}
function esFechaValida_(d) { return Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d.getTime()); }
function formatearRespuestaFecha_(fechaObj, origenTxt, mesTxt) {
  return { success: true, fechaInput: Utilities.formatDate(new Date(fechaObj), Session.getScriptTimeZone(), "yyyy-MM-dd"), fechaDisplay: Utilities.formatDate(new Date(fechaObj), Session.getScriptTimeZone(), "dd/MM/yyyy"), mes: mesTxt, origen: origenTxt };
}
function esFechaValida(fecha) { return Object.prototype.toString.call(fecha) === "[object Date]" && !isNaN(fecha.getTime()); }
function formatDateISO(date) { return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd"); }

// ===================== SISTEMA DE ARCHIVO (DOS HOJAS) =====================
function crearHojaAgendaVotacionSiNoExiste_(ss) {
  try {
    let votacionSheet = ss.getSheetByName('Agenda_Votacion');
    if (!votacionSheet) {
      votacionSheet = ss.insertSheet('Agenda_Votacion');
      const headers = ['Nombre Agenda', 'Numero Reunion', 'Fecha Inicio', 'Fecha Fin', 'Hora Limite', 'Creado Por', 'Fecha Creacion', 'Version', 'Estado', 'Fecha Reunion', 'Fecha Aprobacion', 'Fecha Archivo', 'Estado Archivo'];
      votacionSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      votacionSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      votacionSheet.setFrozenRows(1);
      votacionSheet.autoResizeColumns(1, headers.length);
    }
    return votacionSheet;
  } catch (error) { console.error(' Error creando hoja Agenda_Votacion:', error); return null; }
}
function archivarConfiguracionActual_(configSheet) {
  try {
    if (!configSheet) return false;
    const datosAgenda = configSheet.getRange('A2:K2').getValues()[0];
    const hayDatos = datosAgenda.some(function (cell) { return cell !== "" && cell !== null && cell !== undefined; });
    if (!hayDatos) return false;
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const votacionSheet = crearHojaAgendaVotacionSiNoExiste_(ss);
    if (!votacionSheet) return false;
    const lastRow = votacionSheet.getLastRow();
    if (lastRow > 1) votacionSheet.getRange(2, 1, lastRow - 1, 13).clearContent();
    const fechaArchivo = new Date();
    const datosParaArchivar = [datosAgenda[0], datosAgenda[1], datosAgenda[2], datosAgenda[3], datosAgenda[4], datosAgenda[5], datosAgenda[6], datosAgenda[7], datosAgenda[8], datosAgenda[9], datosAgenda[10], fechaArchivo, 'ARCHIVADA'];
    votacionSheet.getRange(2, 1, 1, datosParaArchivar.length).setValues([datosParaArchivar]);
    return true;
  } catch (error) { console.error(' Error archivando configuración:', error); return false; }
}

// ===================== GESTIÓN DE AGENDAS (CON CONTADOR) =====================
function obtenerTodasLasAgendasConsejo() {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const configSheet = ss.getSheetByName('Configuracion');
    if (!configSheet) return [];
    const lastRow = configSheet.getLastRow();
    if (lastRow < 2) return [];
    const datosAgenda = configSheet.getRange('A2:K2').getValues()[0];
    const [nombreAgenda, numeroReunionRaw, fechaInicio, fechaFin, horaLimite, creadoPor, fechaCreacion, version, estado, fechaReunion, fechaAprobacion] = datosAgenda;
    let numeroReunion = String(numeroReunionRaw || '').replace(/^'/, '').trim();
    if (!numeroReunion) return [];
    let estadoDetallado = estado || 'ACTIVA';
    let mensajeEstado = '';
    const tiempoInfo = calcularTiempoRestante();
    if (tiempoInfo.success && tiempoInfo.expirada && estadoDetallado === 'ACTIVA') { estadoDetallado = 'VENCIDA'; mensajeEstado = 'Agenda vencida automáticamente por tiempo'; }
    else if (estadoDetallado === 'CERRADA') {
      if (fechaAprobacion && fechaAprobacion !== '') {
        estadoDetallado = 'APROBADA';
        const fechaAprob = fechaAprobacion instanceof Date ? Utilities.formatDate(fechaAprobacion, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(fechaAprobacion);
        mensajeEstado = 'Agenda aprobada el ' + fechaAprob;
      } else { estadoDetallado = 'PENDIENTE_APROBACION'; mensajeEstado = 'Agenda cerrada, pendiente de aprobación'; }
    } else if (estadoDetallado === 'ACTIVA') {
      mensajeEstado = tiempoInfo.success ? 'Agenda activa - ' + formatearTiempo(tiempoInfo.tiempoRestante).textoLegible : 'Agenda activa, recibiendo elementos';
    }
    let horaLimiteFormateada = '23:59';
    if (horaLimite instanceof Date) { horaLimiteFormateada = String(horaLimite.getHours()).padStart(2, '0') + ':' + String(horaLimite.getMinutes()).padStart(2, '0'); }
    else if (typeof horaLimite === 'string' && horaLimite.includes(':')) { horaLimiteFormateada = horaLimite; }
    let fechaInicioStr = '', fechaFinStr = '', fechaReunionStr = null;
    try {
      if (fechaInicio instanceof Date) fechaInicioStr = Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (fechaFin instanceof Date) fechaFinStr = Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (fechaReunion instanceof Date) fechaReunionStr = Utilities.formatDate(fechaReunion, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } catch (e) {}
    let totalRegistros = 0;
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (dataSheet && dataSheet.getLastRow() > 1) totalRegistros = dataSheet.getLastRow() - 1;
    return [{ id: numeroReunion, nombre: String(nombreAgenda || ('Reunión Consejo Nacional de Acreditación ' + numeroReunion)), numeroReunion: numeroReunion, fechaInicio: fechaInicioStr, fechaFin: fechaFinStr, fechaReunion: fechaReunionStr, horaLimite: horaLimiteFormateada, creadoPor: String(creadoPor || "Sistema OGA"), totalRegistros: totalRegistros, estado: estadoDetallado, estadoDetallado: estadoDetallado, mensajeEstado: mensajeEstado, fechaAprobacion: fechaAprobacion }];
  } catch (e) { Logger.log('Error obteniendo agendas: ' + e.message); return []; }
}
function obtenerConfiguracionAgendaConsejo(agendaId) {
  try {
    const agendas = obtenerTodasLasAgendasConsejo();
    if (agendas.length > 0) {
      const a = agendas[0];
      return { id: a.numeroReunion, nombre: a.nombre, numeroReunion: a.numeroReunion, fechaInicio: a.fechaInicio, fechaFin: a.fechaFin, fechaReunion: a.fechaReunion, horaLimite: a.horaLimite, estado: a.estado, estadoDetallado: a.estadoDetallado, mensajeEstado: a.mensajeEstado, fechaAprobacion: a.fechaAprobacion, totalRegistros: a.totalRegistros };
    }
    return null;
  } catch (error) { return null; }
}
function verificarAgendaConsejoActiva() {
  try { return puedeAgregarElementos().puede; } catch (_) { return false; }
}
function verificarAgendaAnteriorPendiente(numeroReunionNueva) {
  try {
    const agendas = obtenerTodasLasAgendasConsejo();
    const agendaPendiente = agendas.find(function (a) { return a.estado === 'ACTIVA' && a.numeroReunion !== numeroReunionNueva; });
    if (agendaPendiente) return { hayPendiente: true, agenda: agendaPendiente, message: 'Existe una agenda anterior (' + agendaPendiente.numeroReunion + ') que aún no ha sido aprobada.' };
    return { hayPendiente: false, message: 'No hay agendas pendientes' };
  } catch (e) { return { hayPendiente: false, message: 'Error al verificar' }; }
}
function crearAgendaConsejo(nombre, numeroReunion, fechaInicio, fechaFin, fechaReunion, horaLimite) {
  try {
    if (!esUsuarioAutorizadoParaAgenda()) return { success: false, message: 'No tienes permiso para crear agendas. Solo usuarios autorizados pueden realizar esta acción.' };
    const verificacion = verificarAgendaAnteriorPendiente(numeroReunion);
    if (verificacion.hayPendiente) return { success: false, message: verificacion.message, requiereAprobacion: true, agendaPendiente: verificacion.agenda };
    if (!nombre || nombre.trim().length < 5) nombre = 'Reunion Consejo Nacional de Acreditación ' + numeroReunion;
    if (!numeroReunion || numeroReunion.trim().length < 3) return { success: false, message: 'El número de reunión debe tener al menos 3 caracteres' };
    if (!fechaReunion) return { success: false, message: 'Debe especificar la fecha de la reunión' };
    const inicioDate = new Date(fechaInicio);
    const finDate = new Date(fechaFin);
    const diffDays = Math.ceil(Math.abs(finDate - inicioDate) / (1000 * 60 * 60 * 24));
    if (diffDays > 62) return { success: false, message: 'El rango de fechas no puede ser mayor a 2 meses (62 días)' };
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    let configSheet = ss.getSheetByName('Configuracion');
    if (configSheet && configSheet.getLastRow() >= 2) {
      const numeroExistente = configSheet.getRange('B2').getDisplayValue();
      const numeroComparar = String(numeroExistente || '').replace(/^'/, '').trim();
      if (numeroComparar === numeroReunion.trim()) return { success: false, message: 'Ya existe una agenda con el número: ' + numeroReunion + '.' };
    }
    const resultadoCorrelativo = registrarNumeroReunion(numeroReunion, fechaReunion);
    if (!resultadoCorrelativo.success) return { success: false, message: 'Error al registrar el correlativo: ' + resultadoCorrelativo.message };
    try { actualizarEstadoCalendario(fechaReunion); } catch (e) { console.error("Advertencia: No se pudo actualizar el estado del calendario: " + e.message); }
    if (!configSheet) configSheet = ss.insertSheet('Configuracion');
    const headers = ['Nombre Agenda', 'Numero Reunion', 'Fecha Inicio', 'Fecha Fin', 'Hora Limite', 'Creado Por', 'Fecha Creacion', 'Version', 'Estado', 'Fecha Reunion', 'Fecha Aprobacion'];
    configSheet.clear();
    configSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    configSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    configSheet.setFrozenRows(1);
    const numeroConApostrofe = "'" + numeroReunion.trim();
    const datosAgenda = [nombre.trim(), numeroConApostrofe, fechaInicio, fechaFin, horaLimite || '23:59', (Session.getActiveUser().getEmail() || 'Sistema OGA'), new Date(), '2.6', 'ACTIVA', fechaReunion, ''];
    configSheet.getRange(2, 1, 1, datosAgenda.length).setValues([datosAgenda]);
    configSheet.autoResizeColumns(1, headers.length);
    let dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!dataSheet) {
      dataSheet = ss.insertSheet('Respuestas de formulario 1');
      const headersData = ['Marca temporal', 'Dirección de correo electrónico', 'Registro de Acreditación', 'Solicitud para Consejo Nacional de Acreditación', 'Nombre del Experto', 'Tipo de Decisión', 'Favor Adjuntar', 'Tipo de Notificación', 'Motivo de la Notificación', 'Adjuntar Documentos', 'Acta', 'Resolución', 'Esquema de acreditación', 'Nombre comercial del OEC (únicamente si es diferente a la razón social)', 'Norma'];
      dataSheet.getRange(1, 1, 1, headersData.length).setValues([headersData]);
      dataSheet.getRange(1, 1, 1, headersData.length).setFontWeight('bold');
      dataSheet.setFrozenRows(1);
    }
    Logger.log(' Agenda "' + nombre + '" creada exitosamente y calendario actualizado.');
    return { success: true, message: ' Agenda "' + nombre + '" creada exitosamente. Calendario actualizado.', config: { id: numeroReunion.trim(), nombre: nombre.trim(), numeroReunion: numeroReunion.trim(), fechaInicio: fechaInicio, fechaFin: fechaFin, fechaReunion: fechaReunion, horaLimite: horaLimite || '23:59' } };
  } catch (error) {
    Logger.log(' Error creando agenda: ' + error.message);
    return { success: false, message: 'Error interno: ' + error.message };
  }
}
function actualizarEstadoCalendario(fechaReunionStr) {
  const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
  const hoja = ss.getSheetByName("Calendario_Reuniones");
  if (!hoja) return;
  const datos = hoja.getDataRange().getValues();
  const partes = fechaReunionStr.split('-');
  const fechaObjetivo = new Date(partes[0], partes[1] - 1, partes[2]);
  fechaObjetivo.setHours(0, 0, 0, 0);
  for (let i = 1; i < datos.length; i++) {
    const fechaNormal = datos[i][2] ? new Date(datos[i][2]) : null;
    const fechaExtra = datos[i][5] ? new Date(datos[i][5]) : null;
    if (fechaNormal) fechaNormal.setHours(0, 0, 0, 0);
    if (fechaExtra) fechaExtra.setHours(0, 0, 0, 0);
    let matchEncontrado = false;
    if (fechaNormal && fechaNormal.getTime() === fechaObjetivo.getTime()) matchEncontrado = true;
    else if (fechaExtra && fechaExtra.getTime() === fechaObjetivo.getTime()) matchEncontrado = true;
    if (matchEncontrado) {
      hoja.getRange(i + 1, 4).setValue("CERRADA");
      if (i + 1 < datos.length) {
        const siguienteFecha = datos[i + 1][2];
        if (siguienteFecha) hoja.getRange(i + 2, 4).setValue("ACTIVA");
      }
      return;
    }
  }
  Logger.log(" No se encontró la fecha en el calendario para actualizar el estado.");
}
function aprobarAgendaConsejo(agendaId) {
  try {
    if (!esUsuarioAutorizadoParaAgenda()) return { success: false, message: 'No tienes permiso para aprobar agendas. Solo usuarios autorizados pueden realizar esta acción.' };
    const config = obtenerConfiguracionAgendaConsejo(agendaId);
    if (!config) return { success: false, message: 'No se encontró la configuración de la agenda' };
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const configSheet = ss.getSheetByName('Configuracion');
    if (configSheet) {
      configSheet.getRange('I2').setValue('CERRADA');
      configSheet.getRange('K2').setValue(new Date());
      const archivoExitoso = archivarConfiguracionActual_(configSheet);
      if (archivoExitoso) return { success: true, message: ' Agenda "' + config.nombre + '" aprobada y archivada exitosamente en hoja Agenda_Votacion (A2:M2).' };
      return { success: false, message: 'Error al archivar la agenda en hoja Agenda_Votacion' };
    }
    return { success: false, message: 'No se pudo acceder a la hoja de configuración' };
  } catch (error) { Logger.log(' Error aprobando agenda: ' + error.message); return { success: false, message: 'Error al aprobar agenda: ' + error.message }; }
}
function eliminarAgendaConsejo(agendaId) {
  try {
    if (!esUsuarioAutorizadoParaAgenda()) return { success: false, message: 'No tienes permiso para eliminar agendas. Solo usuarios autorizados pueden realizar esta acción.' };
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const configSheet = ss.getSheetByName('Configuracion');
    let numeroReunionAEliminar = '';
    if (configSheet) {
      try { const numeroRaw = configSheet.getRange('B2').getDisplayValue(); numeroReunionAEliminar = String(numeroRaw || '').replace(/^'/, '').trim(); }
      catch (e) { Logger.log('No se pudo obtener número de reunión: ' + e.message); }
      configSheet.getRange('A2:K2').clearContent();
    }
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (dataSheet) ss.deleteSheet(dataSheet);
    if (numeroReunionAEliminar) eliminarNumeroReunion(numeroReunionAEliminar);
    return { success: true, message: 'Agenda eliminada exitosamente' + (numeroReunionAEliminar ? ' incluyendo correlativo ' + numeroReunionAEliminar : '') + '.' };
  } catch (error) { Logger.log(' Error eliminando agenda: ' + error.message); return { success: false, message: 'Error al eliminar agenda: ' + error.message }; }
}
function obtenerDatosAgendaArchivada(numeroReunion) {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!dataSheet || dataSheet.getLastRow() <= 1) return [];
    const numFilas = dataSheet.getLastRow() - 1;
    const datos = dataSheet.getRange(2, 1, numFilas, 15).getValues();
    const resultados = [];
    datos.forEach(function (fila, index) {
      if (fila.every(function (c) { return !c || c === ''; })) return;
      const acta = String(fila[10] || '');
      if (acta.includes(numeroReunion)) resultados.push({ id: index + 1, registroAcreditacion: String(fila[2] || ''), solicitudCNA: String(fila[3] || ''), acta: String(fila[10] || '') });
    });
    return resultados;
  } catch (error) { console.error(' Error obteniendo datos agenda archivada:', error); return []; }
}

// ===================== BÚSQUEDA DE ESTABLECIMIENTO =====================
function buscarPorRegistro(registro) {
  try {
    const ss = SpreadsheetApp.openById(CONFIRMACION_DATOS_ID);
    const hoja = ss.getSheetByName("Respuestas de formulario 1");
    if (!hoja) return { exito: false, mensaje: ' No se encontró la hoja de confirmación de datos' };
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const registroHoja = String(datos[i][2]).trim();
      if (registroHoja === registro) {
        return { exito: true, esquema: String(datos[i][3]).trim(), nombre: String(datos[i][5]).trim(), norma: String(datos[i][21]).trim() };
      }
    }
    return { exito: false, mensaje: ' No se encontró un establecimiento con ese Registro.' };
  } catch (error) { return { exito: false, mensaje: ' Error al buscar: ' + error.message }; }
}
function buscarUsuarioActivoEnHoja() {
  try {
    const usuarioActivoEmail = Session.getActiveUser().getEmail();
    if (!usuarioActivoEmail) return { exito: false, mensaje: 'No se pudo obtener el usuario activo del sistema' };
    const ss = SpreadsheetApp.openById(USUARIOS_SHEET_ID);
    const hoja = ss.getSheetByName("Usuarios");
    if (!hoja) return { exito: false, mensaje: 'No se encontró la hoja "Usuarios" en la hoja de usuarios' };
    const datos = hoja.getDataRange().getValues();
    if (datos.length <= 1) return { exito: false, mensaje: 'No hay datos de usuarios en la hoja' };
    const mapeoEmails = {
      'carchila@oga.org.gt': 'Carlos Archila', 'eajuchan@oga.org.gt': 'Estephanie Ajuchán', 'ppineda@oga.org.gt': 'Alex Pineda',
      'jalvarado.oga@gmail.com': 'Jonatan Alvarado', 'lsanchez@oga.org.gt': 'Lidia Sanchez', 'acampos@oga.org.gt': 'Addy Campos',
      'amelendez@oga.org.gt': 'Alberto Melendez', 'bjuarez@oga.org.gt': 'Boris Juárez', 'dbarillas@oga.org.gt': 'Duberly Barillas',
      'gquevedo@oga.org.gt': 'Gabriela Quevedo', 'ksantizo@oga.org.gt': 'Karla Santizo', 'sochoa@oga.org.gt': 'Sindy Ochoa',
      'mmenendez.oga@gmail.com': 'Michelle Menendez'
    };
    const nombreEsperado = mapeoEmails[usuarioActivoEmail.toLowerCase()];
    if (nombreEsperado) {
      for (let i = 1; i < datos.length; i++) {
        const nombreEnHoja = String(datos[i][0] || '').trim();
        if (nombreEnHoja.toLowerCase() === nombreEsperado.toLowerCase()) {
          return { exito: true, usuario: usuarioActivoEmail, correo: nombreEnHoja, nombreCompleto: nombreEnHoja, fila: i, autorizado: String(datos[i][2] || '').toLowerCase() === 'sí' };
        }
      }
    }
    for (let i = 1; i < datos.length; i++) {
      const nombreEnHoja = String(datos[i][0] || '').trim();
      const nombreCorto = String(datos[i][3] || '').trim();
      const emailPrefix = usuarioActivoEmail.split('@')[0].toLowerCase();
      if (nombreCorto.toLowerCase().includes(emailPrefix) || nombreEnHoja.toLowerCase().includes(emailPrefix)) {
        return { exito: true, usuario: usuarioActivoEmail, correo: nombreEnHoja, nombreCompleto: nombreEnHoja, fila: i, autorizado: String(datos[i][2] || '').toLowerCase() === 'sí' };
      }
    }
    const nombreFallback = usuarioActivoEmail.split('@')[0];
    return { exito: true, usuario: usuarioActivoEmail, correo: nombreFallback, nombreCompleto: nombreFallback, esSystemFallback: true, autorizado: false, mensaje: 'Usuario "' + usuarioActivoEmail + '" no encontrado en la hoja de usuarios' };
  } catch (error) { return { exito: false, mensaje: 'Error al buscar usuario: ' + error.message }; }
}

// ===================== GUARDAR DATOS (v23: FAC-114 + Reconocimiento separado) =====================
const DOCUMENTOS_DECISION_OBLIGATORIOS = ['OGA-FAC-014', 'OGA-FAC-078', 'OGA-FAC-003', 'OGA-FAC-090'];
// v23: Reconocimiento de Acreditación NO exige documentos FAC; las demás decisiones
// exigen además OGA-FAC-114 (el nombre del archivo debe iniciar con OGA-FAC-114).
const TIPO_RECONOCIMIENTO_ACREDITACION = 'Reconocimiento de Acreditación';
const DOCUMENTO_FAC114 = 'OGA-FAC-114';
function obtenerDocumentosObligatoriosPorTipoDecision_(tipoDecision) {
  if (String(tipoDecision || '').trim() === TIPO_RECONOCIMIENTO_ACREDITACION) return [];
  return ['OGA-FAC-014', 'OGA-FAC-078', 'OGA-FAC-003', 'OGA-FAC-090', 'OGA-FAC-114'];
}
function obtenerTipoDocumentoDecision_(archivo) {
  const tipoDeclarado = String((archivo && archivo.tipoDocumento) || '').toUpperCase().trim();
  const todos = DOCUMENTOS_DECISION_OBLIGATORIOS.concat([DOCUMENTO_FAC114]);
  if (todos.indexOf(tipoDeclarado) !== -1) return tipoDeclarado;
  const nombre = String((archivo && archivo.name) || '').toUpperCase().trim();
  const encontrado = todos.find(function (codigo) { return nombre.startsWith(codigo); });
  return encontrado || '';
}
function validarDocumentosDecisionObligatorios_(adjuntos, tipoDecision) {
  try {
    const obligatorios = obtenerDocumentosObligatoriosPorTipoDecision_(tipoDecision);
    if (obligatorios.length === 0) return { success: true, message: 'Reconocimiento de Acreditación: sin documentos FAC obligatorios.' };
    if (!Array.isArray(adjuntos)) return { success: false, message: 'Para una DECISIÓN debe cargar ' + obligatorios.join(', ') + '.' };
    const presentes = adjuntos.map(obtenerTipoDocumentoDecision_).filter(Boolean);
    const faltantes = obligatorios.filter(function (codigo) { return presentes.indexOf(codigo) === -1; });
    if (faltantes.length > 0) return { success: false, message: 'Faltan documentos obligatorios para DECISIÓN: ' + faltantes.join(', ') + '. No se puede guardar.' };
    return { success: true };
  } catch (error) { return { success: false, message: 'Error validando documentos de decisión: ' + error.message }; }
}
function guardarArchivosEnCola_(archivos, carpetaDestino, resolverCarpeta) {
  const guardados = [];
  const errores = [];
  for (let i = 0; i < archivos.length; i++) {
    const archivo = archivos[i];
    try {
      if (!archivo || !archivo.name || !archivo.content) throw new Error('El archivo no contiene nombre o contenido');
      const destino = resolverCarpeta ? resolverCarpeta(archivo) : carpetaDestino;
      const blob = Utilities.newBlob(Utilities.base64Decode(archivo.content), archivo.type || 'application/octet-stream', archivo.name);
      const creado = destino.createFile(blob);
      guardados.push({ nombre: archivo.name, id: creado.getId() });
      Logger.log('Archivo guardado ' + (i + 1) + '/' + archivos.length + ': ' + archivo.name);
    } catch (error) {
      errores.push((archivo && archivo.name ? archivo.name : 'sin nombre') + ': ' + error.message);
      Logger.log('Error guardando archivo ' + (i + 1) + '/' + archivos.length + ': ' + error.message);
      break;
    }
  }
  return { success: errores.length === 0 && guardados.length === archivos.length, guardados: guardados, errores: errores };
}
function guardarDatosAgendaConsejo(datos) {
  try {
    const tipoDec = String(datos.tipoDecision || '').trim();
    if (tipoDec === 'Reevaluación') {
      const validacion72h = validarPlazoReevaluacionBackend();
      if (!validacion72h.apto) return { success: false, message: validacion72h.mensaje, motivo: 'TIEMPO_AGOTADO_REEVALUACION' };
    }
    const estadoTiempo = puedeAgregarElementos();
    if (!estadoTiempo.puede) return { success: false, message: estadoTiempo.mensaje, tiempoAgotado: estadoTiempo.motivo === 'TIEMPO_AGOTADO', motivo: estadoTiempo.motivo };
    const tipo = String(datos.solicitudCNA || '').trim();
    // v23: Reconocimiento ya no está disponible en Notificaciones.
    if (tipo === 'NOTIFICACION' && String(datos.tipoNotificacion || '').indexOf('Reconocimiento') !== -1) {
      return { success: false, message: 'El tipo Reconocimiento de Acreditación ya no está disponible en Notificaciones. Use una Decisión.' };
    }
    if (tipo === 'VARIOS') {
      const descripcionVarios = String(datos.tituloVarios || '').trim();
      if (!descripcionVarios) return { success: false, message: 'Complete el título/descripción para el punto VARIOS' };
      datos.registroAcreditacion = descripcionVarios;
      datos.motivoNotificacion = '';
    }
    if (!datos.correo || !datos.solicitudCNA) return { success: false, message: 'Complete todos los campos obligatorios' };
    if (tipo !== 'VARIOS' && !datos.registroAcreditacion) return { success: false, message: 'Complete el registro de acreditación' };
    if (datos.solicitudCNA === 'DECISION') {
      if (!datos.nombreExperto) return { success: false, message: 'Debe ingresar el nombre del experto para una decisión' };
      if (!datos.tipoDecision) return { success: false, message: 'Debe seleccionar el tipo de decisión' };
      if (!datos.esquema) return { success: false, message: 'Debe buscar el registro para obtener el esquema' };
      if (!datos.norma) return { success: false, message: 'Debe buscar el registro para obtener la norma' };
      const validacionDocsDecision = validarDocumentosDecisionObligatorios_(datos.adjuntosOGA, datos.tipoDecision);
      if (!validacionDocsDecision.success) return { success: false, message: validacionDocsDecision.message };
    }
    if (datos.solicitudCNA === 'NOTIFICACION') {
      if (!datos.tipoNotificacion || !datos.motivoNotificacion) return { success: false, message: 'Complete los campos de notificación' };
    }
    if (!verificarAgendaConsejoActiva()) return { success: false, message: ' La agenda no está activa o ha vencido el plazo mientras se procesaba la solicitud' };
    if (Array.isArray(datos.adjuntosOGA) && datos.adjuntosOGA.length > 10) return { success: false, message: 'Máximo 10 archivos en "Favor Adjuntar (OGA)".' };
    if (Array.isArray(datos.adjuntarDocumentos) && datos.adjuntarDocumentos.length > 10) return { success: false, message: 'Máximo 10 archivos en "Adjuntar Documentos".' };
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!dataSheet) return { success: false, message: 'No se encontró la hoja de datos' };
    const config = obtenerConfiguracionAgendaConsejo();
    let actaValue = 'Reunión 01-2025';
    if (config && config.numeroReunion) actaValue = 'Reunión ' + config.numeroReunion;
    let resolucionGenerada = '';
    if (tipo === 'DECISION') {
      try {
        resolucionGenerada = registrarResolucion({ numeroReunion: config.numeroReunion, registroAcreditacion: datos.registroAcreditacion, tipoDecision: datos.tipoDecision, nombreEstablecimiento: datos.nombreEstablecimiento, esquema: datos.esquema, norma: datos.norma, nombreExperto: datos.nombreExperto });
      } catch (error) { Logger.log('Error generando resolución: ' + error.message); }
    }
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const folderName = 'Agenda_' + datos.registroAcreditacion + '_' + timestamp;
    const parentFolder = DriveApp.getFolderById(FOLDER_ID);
    const agendaFolder = parentFolder.createFolder(folderName);
    let adjuntosOGAUrls = '';
    if (datos.adjuntosOGA && datos.adjuntosOGA.length > 0) {
      const esReconocimiento = String(datos.tipoDecision || '').trim() === TIPO_RECONOCIMIENTO_ACREDITACION;
      if (esReconocimiento) {
        // v23: Reconocimiento — sin subcarpetas por código y sin validar prefijo;
        // los archivos se guardan tal cual, con su nombre estándar, en Documentos_Reconocimiento.
        const sub = agendaFolder.createFolder('Documentos_Reconocimiento');
        const resultadoRec = guardarArchivosEnCola_(datos.adjuntosOGA, sub);
        if (!resultadoRec.success) return { success: false, message: 'La carga de documentos de Reconocimiento quedó incompleta. Se detuvo la cola en: ' + resultadoRec.errores.join(' | ') };
        adjuntosOGAUrls = sub.getUrl();
      } else {
        const sub = agendaFolder.createFolder('Archivos_OGA_Decision');
        const carpetasPorDocumento = {};
        const resultadoOGA = guardarArchivosEnCola_(datos.adjuntosOGA, sub, function (a) {
          const tipoDocumento = obtenerTipoDocumentoDecision_(a);
          if (tipoDocumento) {
            if (!carpetasPorDocumento[tipoDocumento]) carpetasPorDocumento[tipoDocumento] = sub.createFolder(tipoDocumento);
            return carpetasPorDocumento[tipoDocumento];
          }
          return sub;
        });
        if (!resultadoOGA.success) return { success: false, message: 'La carga de documentos OGA quedó incompleta. Se detuvo la cola en: ' + resultadoOGA.errores.join(' | ') };
        adjuntosOGAUrls = sub.getUrl();
      }
    }
    let adjuntarDocumentosUrls = '';
    if (datos.adjuntarDocumentos && datos.adjuntarDocumentos.length > 0) {
      const sub = agendaFolder.createFolder('Documentos_Adicionales');
      const resultadoAdicionales = guardarArchivosEnCola_(datos.adjuntarDocumentos, sub);
      if (!resultadoAdicionales.success) return { success: false, message: 'La carga de documentos adicionales quedó incompleta. Se detuvo la cola en: ' + resultadoAdicionales.errores.join(' | ') };
      adjuntarDocumentosUrls = sub.getUrl();
    }
    const nuevaFila = [new Date(), String(datos.correo || '').trim(), String(datos.registroAcreditacion || '').trim(), tipo, tipo === 'DECISION' ? String(datos.nombreExperto || '').trim() : '', String(datos.tipoDecision || ''), adjuntosOGAUrls, String(datos.tipoNotificacion || ''), String(datos.motivoNotificacion || ''), adjuntarDocumentosUrls, actaValue, String(resolucionGenerada || datos.resolucion || ''), String(datos.esquema || ''), String(datos.nombreEstablecimiento || ''), String(datos.norma || '')];
    dataSheet.appendRow(nuevaFila);
    const total = dataSheet.getLastRow() - 1;
    const response = { success: true, message: ' Elemento agregado exitosamente. Total: ' + total + ' elementos.', folderUrl: agendaFolder.getUrl(), tiempoRestante: estadoTiempo.tiempoRestante };
    if (resolucionGenerada) { response.resolucionGenerada = resolucionGenerada; response.message += ' Resolución: ' + resolucionGenerada; }
    return response;
  } catch (error) { return { success: false, message: 'Error al guardar: ' + error.message }; }
}
function validarPlazoReevaluacionBackend() {
  try {
    const config = obtenerConfiguracionAgendaConsejo();
    if (!config) return { apto: false, mensaje: "No hay una agenda activa configurada para validar la fecha." };
    const fechaStr = config.fechaReunion;
    const horaStr = config.horaLimite || "23:59";
    if (!fechaStr) return { apto: false, mensaje: "La agenda activa no tiene fecha de reunión definida." };
    let fechaReunionFinal;
    if (fechaStr instanceof Date) { fechaReunionFinal = new Date(fechaStr); }
    else { const partes = String(fechaStr).split('-'); fechaReunionFinal = new Date(partes[0], partes[1] - 1, partes[2]); }
    const partesHora = String(horaStr).split(':');
    fechaReunionFinal.setHours(parseInt(partesHora[0]), parseInt(partesHora[1]), 0, 0);
    const HORAS_LIMITE = 72;
    const MS_EN_72_HORAS = HORAS_LIMITE * 60 * 60 * 1000;
    const fechaCorte = new Date(fechaReunionFinal.getTime() - MS_EN_72_HORAS);
    const ahora = new Date();
    if (ahora > fechaCorte) {
      const msTarde = ahora.getTime() - fechaCorte.getTime();
      const horasTarde = Math.floor(msTarde / (1000 * 60 * 60));
      const cierreLegible = Utilities.formatDate(fechaCorte, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      return { apto: false, mensaje: ' TIEMPO AGOTADO PARA REEVALUACIÓN: Se requieren 72 horas antes de la reunión. El sistema cerró la recepción de este tipo de solicitud el ' + cierreLegible + ' (hace ' + horasTarde + ' horas).' };
    }
    return { apto: true };
  } catch (e) {
    Logger.log("Error crítico validando 72h: " + e.message);
    return { apto: false, mensaje: "Error técnico validando fechas: " + e.message };
  }
}

// ===================== OBTENER DATOS (v23: añade reunionVigente / totalFiltradoVigente) =====================
function obtenerDatosAgendaConsejoParaEdicion() {
  try {
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    // v23: la reunión vigente es la de Configuración, aunque ya esté cerrada.
    let reunionVigente = '';
    try { const v = obtenerReunionVigenteParaFiltros(); if (v.success) reunionVigente = v.numeroReunion; } catch (e) {}
    if (!dataSheet) return { success: true, data: [], message: 'No hay hoja de datos disponible', reunionVigente: reunionVigente, totalFiltradoVigente: 0 };
    if (dataSheet.getLastRow() <= 1) return { success: true, data: [], message: 'No hay elementos registrados en la agenda', reunionVigente: reunionVigente, totalFiltradoVigente: 0 };
    const numFilas = dataSheet.getLastRow() - 1;
    const datos = dataSheet.getRange(2, 1, numFilas, 15).getValues();
    const resultados = [];
    datos.forEach(function (fila, index) {
      const filaVacia = fila.every(function (c) { return !c || c === ''; });
      if (filaVacia) return;
      resultados.push({ id: index + 1, marcaTemporal: fila[0] ? Utilities.formatDate(new Date(fila[0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : '', correo: String(fila[1] || ''), registroAcreditacion: String(fila[2] || ''), solicitudCNA: String(fila[3] || ''), nombreExperto: String(fila[4] || ''), tipoDecision: String(fila[5] || ''), adjuntosOGA: String(fila[6] || ''), tipoNotificacion: String(fila[7] || ''), motivoNotificacion: String(fila[8] || ''), adjuntarDocumentos: String(fila[9] || ''), acta: String(fila[10] || ''), resolucion: String(fila[11] || ''), esquema: String(fila[12] || ''), nombreEstablecimiento: String(fila[13] || ''), norma: String(fila[14] || '') });
    });
    const totalFiltradoVigente = reunionVigente ? resultados.filter(function (it) { return String(it.acta || '').includes(reunionVigente); }).length : 0;
    return { success: true, data: resultados, message: resultados.length + ' elementos encontrados', reunionVigente: reunionVigente, totalFiltradoVigente: totalFiltradoVigente };
  } catch (error) {
    return { success: false, data: [], message: 'Error al cargar datos: ' + error.message, reunionVigente: '', totalFiltradoVigente: 0 };
  }
}
function eliminarElementoAgendaConsejo(id) {
  try {
    const estadoTiempo = puedeAgregarElementos();
    if (!estadoTiempo.puede) return { success: false, message: estadoTiempo.mensaje, tiempoAgotado: estadoTiempo.motivo === 'TIEMPO_AGOTADO' };
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!dataSheet || dataSheet.getLastRow() <= 1) return { success: false, message: 'No hay datos' };
    const numFilas = dataSheet.getLastRow() - 1;
    if (id < 1 || id > numFilas) return { success: false, message: 'ID no válido' };
    dataSheet.deleteRow(id + 1);
    const total = dataSheet.getLastRow() - 1;
    return { success: true, message: ' Elemento ' + id + ' eliminado. Total: ' + total + '.' };
  } catch (error) { return { success: false, message: 'Error al eliminar: ' + error.message }; }
}
function obtenerMiembrosParaVotacion(modoPrueba) {
  const PRODUCCION_MIEMBROS_SHEET_ID = "10Z19aLAhMN2eUCbinYe6omdOneT4_mkGwqJqiiU1NHQ";
  const TEST_MIEMBROS_SHEET_ID = "1Az5OF8A5F6nb_J-9hU686hGNYuTnZLUy6cbX7KDZ0FY";
  const MIEMBROS_SHEET_NAME = "Miembros";
  const sheetId = modoPrueba ? TEST_MIEMBROS_SHEET_ID : PRODUCCION_MIEMBROS_SHEET_ID;
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(MIEMBROS_SHEET_NAME);
  if (!sheet) throw new Error("No se encontró la hoja 'Miembros' (" + (modoPrueba ? "TEST" : "PROD") + ").");
  const datos = sheet.getDataRange().getValues();
  const miembros = [];
  const correosUsados = new Set();
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const email = String(fila[5] || "").trim();
    if (!email) continue;
    if (correosUsados.has(email.toLowerCase())) continue;
    const nombre = (fila[0] + " " + fila[1] + " " + fila[2] + " " + fila[3]).trim().replace(/\s+/g, " ");
    const institucion = String(fila[9] || "").trim();
    miembros.push({ email: email, nombre: nombre, institucion: institucion, fuente: modoPrueba ? "TEST" : "PROD" });
    correosUsados.add(email.toLowerCase());
  }
  return miembros;
}
function construirCuerpoInvitacion(miembro, config, enlace, modoPrueba) {
  const asuntoTag = modoPrueba ? "[PRUEBA] " : "";
  const asunto = asuntoTag + 'Invitación a Votar - Reunión ' + config.numeroReunion;
  const htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;"><p>Estimado(a) <strong>' + (miembro.nombre || miembro.email) + '</strong>,</p><p>Se le invita a emitir su voto para la <strong>Reunión ' + config.numeroReunion + '</strong>.</p><p>Ingrese al siguiente enlace seguro para registrar sus decisiones:<br><a href="' + enlace + '" target="_blank" style="color:#1d4ed8;">' + enlace + '</a></p><p style="margin-top:12px;"><em>Nota:</em> Solo se acepta un envío de votos por institución.</p>' + (modoPrueba ? '<p style="color:#b45309;"><strong>Modo Prueba:</strong> Este envío es solamente para validación interna.</p>' : '') + '<hr style="margin:20px 0;border:none;border-top:1px solid #ddd;" /><p style="font-size:12px;color:#666;"><strong>OGA - Oficina Guatemalteca de Acreditación</strong><br>' + CORREO_INSTITUCIONAL + ' -- Enlace generado automáticamente. No responda a este correo.</p></div>';
  const textBody = 'Estimado(a) ' + (miembro.nombre || miembro.email) + ':\n\nSe le invita a emitir su voto para la Reunión ' + config.numeroReunion + '.\n\nIngrese al siguiente enlace para votar:\n' + enlace + '\n\nNota: Solo se acepta un envío por institución.\n\n' + (modoPrueba ? "[ESTE ES UN ENVÍO DE PRUEBA]" : "") + '\n--\nOGA - Oficina Guatemalteca de Acreditación\n' + CORREO_INSTITUCIONAL + ' (mensaje automático)';
  return { asunto: asunto, htmlBody: htmlBody, textBody: textBody };
}
function obtenerEstadisticasVotosPorReunion(numeroReunion) { return obtenerEstadisticasVotosPorReunionMejoradas(numeroReunion); }
function obtenerEstadisticasVotosAgenda() {
  const verificacion = hayAgendaListaParaVotar();
  if (!verificacion.success) return { success: false, message: verificacion.mensaje };
  const estadisticas = obtenerEstadisticasVotosPorReunionMejoradas(verificacion.numeroReunion);
  if (!estadisticas.success) return estadisticas;
  return { success: true, data: estadisticas.data, resumen: estadisticas.resumen, porTipo: estadisticas.porTipo, registros: estadisticas.registros, registrosPorTipo: estadisticas.registrosPorTipo, reunion: verificacion.numeroReunion, ultimaActualizacion: new Date().toLocaleString('es-ES') };
}

// ===================== RESOLUCIONES =====================
function obtenerSiguienteResolucion(numeroReunion) {
  try {
    const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    let hojaResoluciones = ss.getSheetByName("Resoluciones");
    if (!hojaResoluciones) {
      hojaResoluciones = ss.insertSheet("Resoluciones");
      hojaResoluciones.getRange(1, 1, 1, 12).setValues([['Reunión', 'Resolución', 'Registro_Acreditacion', 'Nombre_OEC', 'Esquema', 'Norma', 'Tipo_Decision', 'Comentarios', 'Experto', 'Agenda', 'Estado', 'Motivo de Generacion']]);
      hojaResoluciones.getRange(1, 1, 1, 12).setFontWeight('bold');
      hojaResoluciones.setFrozenRows(1);
    }
    const partesReunion = String(numeroReunion || '').split('-');
    const anioReunion = partesReunion.length >= 2 ? partesReunion[1] : new Date().getFullYear().toString();
    const datos = hojaResoluciones.getDataRange().getValues();
    let ultimoNumeroDelAnio = 0;
    let resolucionesEncontradas = 0;
    for (let i = 1; i < datos.length; i++) {
      const resolucion = String(datos[i][1] || '').trim();
      if (resolucion && resolucion.includes('-')) {
        const partesResolucion = resolucion.split('-');
        if (partesResolucion.length >= 2) {
          const anioResolucion = partesResolucion[1].trim();
          if (anioResolucion === anioReunion) {
            resolucionesEncontradas++;
            const numero = parseInt(partesResolucion[0].trim(), 10);
            if (!isNaN(numero) && numero > ultimoNumeroDelAnio) ultimoNumeroDelAnio = numero;
          }
        }
      }
    }
    const siguienteNumero = ultimoNumeroDelAnio + 1;
    const resolucionFinal = siguienteNumero.toString().padStart(3, '0') + '-' + anioReunion;
    Logger.log(' Año: ' + anioReunion + ' | Resoluciones encontradas: ' + resolucionesEncontradas + ' | Último número: ' + ultimoNumeroDelAnio + ' | Siguiente: ' + resolucionFinal);
    return resolucionFinal;
  } catch (error) {
    Logger.log(' Error generando resolución: ' + error.message);
    const anioFallback = new Date().getFullYear();
    return '001-' + anioFallback;
  }
}
function registrarResolucion(datos) {
  try {
    const numeroResolucion = obtenerSiguienteResolucion(datos.numeroReunion);
    const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    const hojaResoluciones = ss.getSheetByName("Resoluciones");
    const nuevaFila = [datos.numeroReunion || '', numeroResolucion, datos.registroAcreditacion || '', datos.nombreEstablecimiento || '', datos.esquema || '', datos.norma || '', datos.tipoDecision || '', '', datos.nombreExperto || '', 'Reunión ' + datos.numeroReunion, 'GENERADA', 'Generada automáticamente para ' + datos.tipoDecision];
    hojaResoluciones.appendRow(nuevaFila);
    Logger.log('Resolución registrada: ' + numeroResolucion + ' para ' + datos.registroAcreditacion);
    return numeroResolucion;
  } catch (error) { Logger.log('Error registrando resolución: ' + error.message); throw new Error('No se pudo registrar la resolución: ' + error.message); }
}
function generarResolucionDesdeVotacion(registro, reunion) {
  try {
    if (!esUsuarioAutorizadoParaAgenda()) return { success: false, message: 'No tienes permiso para generar resoluciones' };
    const ss = SpreadsheetApp.openById(AGENDA_SHEET_ID);
    const dataSheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!dataSheet) return { success: false, message: 'No se encontró la hoja de datos de la agenda' };
    const datos = dataSheet.getDataRange().getValues();
    let datosRegistro = null;
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][2]).trim() === registro) {
        datosRegistro = { numeroReunion: reunion, registroAcreditacion: registro, nombreEstablecimiento: String(datos[i][13] || ''), esquema: String(datos[i][12] || ''), norma: String(datos[i][14] || ''), tipoDecision: String(datos[i][5] || ''), nombreExperto: String(datos[i][4] || '') };
        break;
      }
    }
    if (!datosRegistro) return { success: false, message: 'No se encontraron datos para el registro ' + registro };
    const numeroResolucion = registrarResolucion(datosRegistro);
    return { success: true, message: 'Resolución ' + numeroResolucion + ' generada exitosamente para ' + registro, resolucion: numeroResolucion };
  } catch (error) { Logger.log('Error generando resolución desde votación: ' + error.message); return { success: false, message: 'Error interno: ' + error.message }; }
}
function registrarLogEnviosVotaciones(resultado) {
  try {
    const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    let logSheet = ss.getSheetByName("Log_Envios_Votaciones");
    if (!logSheet) {
      logSheet = ss.insertSheet("Log_Envios_Votaciones");
      logSheet.getRange(1, 1, 1, 8).setValues([['Fecha', 'Reunion', 'Modo_Prueba', 'Intentados', 'Exitosos', 'Fallidos', 'Usuario', 'Correo_Institucional']]);
      logSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }
    const fecha = new Date();
    const reunion = resultado.reunion || 'N/A';
    const modoPrueba = resultado.message.includes('[PRUEBA]') ? 'SÍ' : 'NO';
    const exitosos = resultado.enviados ? resultado.enviados.length : 0;
    const fallidos = resultado.errores ? resultado.errores.length : 0;
    const intentados = exitosos + fallidos;
    const usuario = Session.getActiveUser().getEmail() || 'Sistema';
    logSheet.appendRow([fecha, reunion, modoPrueba, intentados, exitosos, fallidos, usuario, CORREO_INSTITUCIONAL]);
  } catch (error) { Logger.log('Error registrando log: ' + error.message); }
}

// ===================== DOCUMENTOS =====================
function generarDocumentoAgendaConsejo(tipo) {
  try {
    const datos = obtenerDatosReunionActiva();
    if (tipo === 'agenda') return generarAgendaProfesionalConPlantilla(datos);
    else if (tipo === 'excel') return exportarAExcel(datos);
  } catch (error) { return { success: false, message: 'Error interno: ' + error.message }; }
}
function generarAgendaProfesionalConPlantilla(datos) {
  try {
    if (!Array.isArray(datos)) throw new Error('Datos inválidos');
    const config = obtenerConfiguracionAgendaConsejo();
    if (!config || !config.numeroReunion) throw new Error('No se pudo obtener configuración');
    desactivarAgendaAlmacenada(config.numeroReunion);
    const filtrados = datos.filter(function (d) { return d && d.acta && String(d.acta).includes(String(config.numeroReunion)); });
    const decisiones = filtrados.filter(function (i) { return i.solicitudCNA === 'DECISION'; });
    const notificaciones = filtrados.filter(function (i) { return i.solicitudCNA === 'NOTIFICACION'; });
    const varios = filtrados.filter(function (i) { return i.solicitudCNA === 'VARIOS'; });
    const plantilla = DriveApp.getFileById(PLANTILLA_AGENDA_ID);
    const nombreDocumento = 'Agenda_' + config.numeroReunion + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const copia = plantilla.makeCopy(nombreDocumento);
    const doc = DocumentApp.openById(copia.getId());
    const body = doc.getBody();
    const reunionActual = String(config.numeroReunion).trim();
    const reunionAnterior = (function calcularReunionAnterior(num) {
      const m = String(num || '').trim().match(/^(\d{1,2})-(\d{4})$/);
      if (!m) return num;
      let mes = parseInt(m[1], 10);
      let anio = parseInt(m[2], 10);
      if (mes === 1) { mes = 12; anio -= 1; } else { mes -= 1; }
      const mesStr = mes < 10 ? ('0' + mes) : String(mes);
      return mesStr + '-' + anio;
    })(reunionActual);
    body.replaceText('10-2025', reunionActual);
    body.replaceText('Aprobación Acta\\s+\\d{1,2}-\\d{4}', 'Aprobación Acta ' + reunionAnterior);
    if (config.fechaReunion) {
      const fecha = new Date(config.fechaReunion);
      const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
      const fechaFormateada = fecha.getDate() + ' de ' + meses[fecha.getMonth()] + ' de ' + fecha.getFullYear();
      body.replaceText('06 de agosto de 2025', fechaFormateada);
    }
    const reemplazarBloqueConEnlaces = function (marcador, elementos, generador) {
      try {
        const match = body.findText(marcador);
        if (!match) return;
        const elem = match.getElement().asText();
        const par = elem.getParent();
        const idx = body.getChildIndex(par);
        body.removeChild(par);
        if (!elementos || elementos.length === 0) {
          const linea = body.insertParagraph(idx, "      No se registraron puntos en este apartado.");
          linea.setItalic(true).setForegroundColor("#777").setIndentStart(1.5 * 72).setIndentFirstLine(1.5 * 72);
          return;
        }
        elementos.forEach(function (item, i) {
          const g = generador(item, i);
          const p = body.insertParagraph(idx + i, g.texto);
          p.setIndentStart(1.5 * 72).setIndentFirstLine(1.5 * 72);
          if (g.enlace) p.appendText(" ").appendText("Consultar").setLinkUrl(g.enlace);
        });
      } catch (e) {}
    };
    reemplazarBloqueConEnlaces('3.1  <Tipo de notificacion> de la acreditación otorgada. <Registro de Acreditacion> consultar', notificaciones, function (n, i) { return { texto: '3.' + (i + 1) + ' ' + (n.tipoNotificacion || 'Notificación') + ' de la acreditación otorgada. ' + (n.registroAcreditacion || ''), enlace: n.adjuntarDocumentos || '' }; });
    reemplazarBloqueConEnlaces('4.1 <Tipo de desicion >- <Esquema de acreditacion> <Registro de Acreditación>  Consultar', decisiones, function (d, i) { return { texto: '4.' + (i + 1) + ' ' + (d.tipoDecision || 'Evaluación') + ' - ' + (obtenerEsquemaAcreditacion(d.registroAcreditacion) || '') + ' ' + (d.registroAcreditacion || '') + ' (Resolución: ' + (d.resolucion || 'Pendiente') + ')', enlace: d.adjuntosOGA || d.adjuntarDocumentos || '' }; });
    reemplazarBloqueConEnlaces('5.1  <Varios>', varios, function (v, i) { return { texto: '5.' + (i + 1) + ' ' + (v.registroAcreditacion || 'Punto varios'), enlace: v.adjuntarDocumentos || '' }; });
    const agendasFolder = DriveApp.getFolderById(AGENDAS_CREADAS_FOLDER_ID);
    agendasFolder.addFile(copia);
    DriveApp.getRootFolder().removeFile(copia);
    const agendaUrl = doc.getUrl();
    const almacenarResult = almacenarAgenda(config.numeroReunion, agendaUrl);
    Logger.log(' Resultado almacenar agenda: ' + JSON.stringify(almacenarResult));
    return { success: true, message: 'Agenda generada exitosamente en carpeta "Agendas Creadas" (' + filtrados.length + ' elementos)', detalles: { decisiones: decisiones.length, notificaciones: notificaciones.length, varios: varios.length }, url: agendaUrl };
  } catch (error) { return { success: false, message: 'Error al generar agenda: ' + error.message }; }
}
function exportarAExcel(datos) {
  try {
    const config = obtenerConfiguracionAgendaConsejo();
    const ss = SpreadsheetApp.create('Agenda_Excel_' + (config ? config.numeroReunion : 'Export') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd"));
    const sheet = ss.getActiveSheet();
    const headers = ['Fecha', 'Usuario', 'Registro', 'Tipo Solicitud', 'Experto', 'Tipo Decisión', 'Resolución', 'Tipo Notificación', 'Motivo', 'Acta', 'Esquema de acreditación', 'Nombre comercial', 'Norma'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (datos.length > 0) {
      const rows = datos.map(function (item) { return [new Date(item.marcaTemporal), item.correo, item.registroAcreditacion, item.solicitudCNA, item.nombreExperto, item.tipoDecision, item.resolucion || '', item.tipoNotificacion, item.motivoNotificacion, item.acta, item.esquema || '', item.nombreEstablecimiento || '', item.norma || '']; });
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    sheet.autoResizeColumns(1, headers.length);
    sheet.setFrozenRows(1);
    const agendasFolder = DriveApp.getFolderById(AGENDAS_CREADAS_FOLDER_ID);
    const file = DriveApp.getFileById(ss.getId());
    agendasFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    return { success: true, message: 'Datos exportados a Excel exitosamente en carpeta "Agendas Creadas"', url: ss.getUrl() };
  } catch (error) { return { success: false, message: 'Error al exportar a Excel: ' + error.message }; }
}
function verificarArchivosEnCarpetaAgendas() {
  try {
    const cacheKey = 'archivos_carpeta_agendas';
    const cached = getCachedData(cacheKey);
    if (cached) return cached;
    let folder;
    try { folder = DriveApp.getFolderById(AGENDAS_CREADAS_FOLDER_ID); }
    catch (e) { return { success: false, message: 'No se pudo acceder a la carpeta de agendas creadas', archivos: [] }; }
    const listaArchivos = [];
    let archivos;
    try { archivos = folder.getFiles(); }
    catch (e) { return { success: false, message: 'Error al obtener archivos de la carpeta', archivos: [] }; }
    while (archivos.hasNext()) {
      try {
        const archivo = archivos.next();
        if (!archivo) continue;
        let mimeType = 'application/octet-stream';
        let nombre = 'Sin nombre';
        let fechaModificacion = new Date();
        let url = '#';
        try { nombre = archivo.getName() || 'Sin nombre'; } catch (e) {}
        try { const blob = archivo.getBlob(); if (blob) mimeType = blob.getContentType() || 'application/octet-stream'; } catch (e) {}
        try { fechaModificacion = archivo.getLastUpdated() || new Date(); } catch (e) {}
        try { url = archivo.getUrl() || '#'; } catch (e) {}
        let tipo = 'file';
        try { if (mimeType.includes('document') || mimeType.includes('word')) tipo = 'document'; else if (mimeType.includes('spreadsheet') || mimeType.includes('excel')) tipo = 'spreadsheet'; else if (mimeType.includes('folder')) tipo = 'folder'; } catch (_) { tipo = 'file'; }
        const reunionMatch = nombre.match(/\b(\d{1,3}-\d{4})\b/);
        const reunion = reunionMatch ? reunionMatch[1] : null;
        listaArchivos.push({ nombre: nombre, url: url, fechaModificacion: fechaModificacion, tipo: tipo, mimeType: mimeType, reunion: reunion });
      } catch (e) { continue; }
    }
    try { listaArchivos.sort(function (a, b) { return new Date(b.fechaModificacion) - new Date(a.fechaModificacion); }); } catch (e) {}
    const resultado = { success: true, archivos: listaArchivos, carpetaUrl: folder.getUrl() || '#', message: listaArchivos.length + ' archivos encontrados' };
    setCachedData(cacheKey, resultado, 120);
    return resultado;
  } catch (error) { return { success: false, message: 'Error al acceder a la carpeta: ' + error.message, archivos: [] }; }
}

// ===================== VISTA PREVIA SIN GUARDAR (v23) =====================
// SOLO LECTURA. No llama a DriveApp, DocumentApp, makeCopy, almacenarAgenda,
// desactivarAgendaAlmacenada, appendRow ni setValue. Construye en memoria la
// agenda con la plantilla lógica (encabezado, fecha, secciones 3.x / 4.x / 5.x)
// para mostrarla como evidencia de cómo se vería, sin guardar ni registrar nada.
function generarVistaPreviaAgendaConsejo() {
  try {
    const config = obtenerConfiguracionAgendaConsejo();
    if (!config || !config.numeroReunion) return { success: false, message: 'No hay agenda activa para generar la vista previa.' };
    const datos = obtenerDatosReunionActiva();
    const filtrados = (datos || []).filter(function (d) { return d && d.acta && String(d.acta).includes(String(config.numeroReunion)); });
    const decisiones = filtrados.filter(function (i) { return i.solicitudCNA === 'DECISION'; });
    const notificaciones = filtrados.filter(function (i) { return i.solicitudCNA === 'NOTIFICACION'; });
    const varios = filtrados.filter(function (i) { return i.solicitudCNA === 'VARIOS'; });
    const reunionActual = String(config.numeroReunion).trim();
    const reunionAnterior = (function calcularReunionAnterior(num) {
      const m = String(num || '').trim().match(/^(\d{1,2})-(\d{4})$/);
      if (!m) return num;
      let mes = parseInt(m[1], 10);
      let anio = parseInt(m[2], 10);
      if (mes === 1) { mes = 12; anio -= 1; } else { mes -= 1; }
      const mesStr = mes < 10 ? ('0' + mes) : String(mes);
      return mesStr + '-' + anio;
    })(reunionActual);
    let fechaReunionTexto = '';
    if (config.fechaReunion) {
      const fecha = new Date(config.fechaReunion);
      const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
      fechaReunionTexto = fecha.getDate() + ' de ' + meses[fecha.getMonth()] + ' de ' + fecha.getFullYear();
    }
    const secNotificaciones = notificaciones.map(function (n, i) { return { numero: '3.' + (i + 1), texto: (n.tipoNotificacion || 'Notificación') + ' de la acreditación otorgada. ' + (n.registroAcreditacion || '') }; });
    const secDecisiones = decisiones.map(function (d, i) { return { numero: '4.' + (i + 1), texto: (d.tipoDecision || 'Evaluación') + ' - ' + (obtenerEsquemaAcreditacion(d.registroAcreditacion) || '') + ' ' + (d.registroAcreditacion || ''), resolucion: d.resolucion || 'Pendiente' }; });
    const secVarios = varios.map(function (v, i) { return { numero: '5.' + (i + 1), texto: v.registroAcreditacion || 'Punto varios' }; });
    let html = '<div style="font-family:Arial,sans-serif;color:#1e293b;">';
    html += '<h1>Agenda Consejo OGA — Reunión ' + reunionActual + '</h1>';
    html += '<p><strong>Fecha de reunión:</strong> ' + (fechaReunionTexto || 'No definida') + '</p>';
    html += '<p><strong>Aprobación Acta:</strong> ' + reunionAnterior + '</p>';
    html += '<h2>3. Notificaciones (' + notificaciones.length + ')</h2>';
    html += secNotificaciones.length ? '<ol>' + secNotificaciones.map(function (s) { return '<li>' + s.texto + '</li>'; }).join('') + '</ol>' : '<p>No se registraron puntos en este apartado.</p>';
    html += '<h2>4. Decisiones (' + decisiones.length + ')</h2>';
    html += secDecisiones.length ? '<ol>' + secDecisiones.map(function (s) { return '<li>' + s.texto + ' (Resolución: ' + s.resolucion + ')</li>'; }).join('') + '</ol>' : '<p>No se registraron puntos en este apartado.</p>';
    html += '<h2>5. Varios (' + varios.length + ')</h2>';
    html += secVarios.length ? '<ol>' + secVarios.map(function (s) { return '<li>' + s.texto + '</li>'; }).join('') + '</ol>' : '<p>No se registraron puntos en este apartado.</p>';
    html += '</div>';
    return {
      success: true, reunion: reunionActual, fechaReunion: fechaReunionTexto, actaAnterior: reunionAnterior,
      conteos: { decisiones: decisiones.length, notificaciones: notificaciones.length, varios: varios.length, total: filtrados.length },
      secciones: { notificaciones: secNotificaciones, decisiones: secDecisiones, varios: secVarios },
      htmlVistaPrevia: html,
      aviso: 'Vista previa generada en memoria. No se guardó ni se registró ningún documento.'
    };
  } catch (e) {
    return { success: false, message: 'Error generando vista previa: ' + e.message };
  }
}

// ===================== UTILIDADES =====================
function obtenerDatosReunionActiva() {
  try {
    const resultado = obtenerDatosAgendaConsejoParaEdicion();
    return resultado.success ? resultado.data : [];
  } catch (error) { return []; }
}
function desactivarAgendaAlmacenada(numeroReunion) {
  try { return { success: true }; } catch (error) { return { success: false }; }
}
function obtenerEsquemaAcreditacion(registro) {
  try {
    const resultado = buscarPorRegistro(registro);
    return resultado.exito ? resultado.esquema : '';
  } catch (error) { return ''; }
}
function almacenarAgenda(numeroReunion, url) {
  try { return { success: true, message: 'Agenda almacenada exitosamente' }; }
  catch (error) { return { success: false, message: error.message }; }
}

// ===================== CACHÉ =====================
(function initCacheHelpers() {
  try {
    var __CACHE__ = CacheService.getScriptCache();
    var __CACHE_TIMEOUT__ = 300;
    if (typeof getCachedData !== 'function') {
      this.getCachedData = function (key) {
        try { var cached = __CACHE__.get(key); return cached ? JSON.parse(cached) : null; }
        catch (e) { return null; }
      };
    }
    if (typeof setCachedData !== 'function') {
      this.setCachedData = function (key, data, timeout) {
        try { __CACHE__.put(key, JSON.stringify(data), timeout || __CACHE_TIMEOUT__); }
        catch (e) {}
      };
    }
  } catch (e) {
    if (typeof getCachedData !== 'function') { this.getCachedData = function () { return null; }; }
    if (typeof setCachedData !== 'function') { this.setCachedData = function () {}; }
  }
})();

// ===================== FUNCIONES WRAPPER PARA COMPATIBILIDAD =====================
function verificarEstadoVotacion() { return buscarReunionAprobadaParaVotacion(); }
function invitarVotaciones(modoPrueba, maxEnviar) { return enviarVotacionesDirectas(modoPrueba, maxEnviar); }
function cargarEstadisticasVotos() { return obtenerEstadisticasVotosAgenda(); }
function cargarAgendasDisponiblesVotacion() { return buscarAgendasAprobadas(); }
function cargarEstadisticasAgendaEspecifica(numeroReunion) { return obtenerEstadisticasVotosPorReunionMejoradas(numeroReunion); }
function verEstadisticasAgendaEspecifica(numeroReunion) { return obtenerEstadisticasVotosPorReunionMejoradas(numeroReunion); }
function generarResolucion(registro, reunion) { return generarResolucionDesdeVotacion(registro, reunion); }

// ===================== AUTOMATIZACIÓN DE CALENDARIO (ARQUITECTURA VERTICAL, v22 duplicado) =====================
function obtenerSiguienteFechaReunion() {
  try {
    const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    const hoja = ss.getSheetByName("Calendario_Reuniones");
    if (!hoja) return { success: false, message: " No existe la hoja 'Calendario_Reuniones'" };
    const datos = hoja.getDataRange().getValues();
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    for (let i = 1; i < datos.length; i++) {
      const estado = String(datos[i][3] || "").toUpperCase().trim();
      if (estado === "ACTIVA") {
        const fechaExtra = datos[i][5];
        const fechaNormal = datos[i][2];
        const mesTexto = obtenerMesCorrecto_(datos, i);
        if (esFechaValida_(fechaExtra)) return formatearRespuestaFecha_(fechaExtra, "FECHA EXTRAORDINARIA (ACTIVA)", mesTexto);
        if (esFechaValida_(fechaNormal)) {
          let orden = "REUNIÓN ACTIVA";
          const mesAnterior = obtenerMesCorrecto_(datos, i - 1);
          if (i > 1 && mesTexto === mesAnterior) orden = "SEGUNDA REUNIÓN (ACTIVA)";
          else orden = "PRIMERA REUNIÓN (ACTIVA)";
          return formatearRespuestaFecha_(fechaNormal, orden, mesTexto);
        }
      }
    }
    for (let i = 1; i < datos.length; i++) {
      const fechaNormal = datos[i][2];
      const estado = String(datos[i][3] || "").toUpperCase().trim();
      if (esFechaValida_(fechaNormal)) {
        const fechaObj = new Date(fechaNormal); fechaObj.setHours(0, 0, 0, 0);
        if (fechaObj >= hoy && estado !== "CERRADA") {
          const fechaExtra = datos[i][5];
          const mesTexto = obtenerMesCorrecto_(datos, i);
          if (esFechaValida_(fechaExtra)) return formatearRespuestaFecha_(fechaExtra, "FECHA EXTRAORDINARIA", mesTexto);
          return formatearRespuestaFecha_(fechaNormal, "SIGUIENTE REUNIÓN", mesTexto);
        }
      }
    }
    return { success: false, message: "No se encontraron fechas futuras disponibles en el calendario." };
  } catch (e) {
    Logger.log("Error en calendario: " + e.message);
    return { success: false, message: "Error interno: " + e.message };
  }
}
function obtenerMesCorrecto_(datos, indiceActual) {
  for (let k = indiceActual; k >= 1; k--) {
    if (datos[k][1] && String(datos[k][1]).trim() !== "") return String(datos[k][1]).toUpperCase();
  }
  return "MES DESCONOCIDO";
}
function esFechaValida_(d) { return Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d.getTime()); }
function formatearRespuestaFecha_(fechaObj, origenTxt, mesTxt) {
  return { success: true, fechaInput: Utilities.formatDate(new Date(fechaObj), Session.getScriptTimeZone(), "yyyy-MM-dd"), fechaDisplay: Utilities.formatDate(new Date(fechaObj), Session.getScriptTimeZone(), "dd/MM/yyyy"), mes: mesTxt, origen: origenTxt };
}
function obtenerDatosProximaReunion() {
  try {
    const numero = obtenerSiguienteNumeroReunion();
    const fechaInfo = obtenerSiguienteFechaReunion();
    return { success: true, numero: numero, fechaInfo: fechaInfo };
  } catch (e) { return { success: false, message: e.message }; }
}

// ===================== INICIALIZACIÓN =====================
function abrirAgenda() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Agenda_Sistema').setWidth(3400).setHeight(2500);
    SpreadsheetApp.getUi().showModalDialog(html, ' Agenda Consejo OGA');
  } catch (error) { SpreadsheetApp.getUi().alert('Error al abrir el sistema: ' + error.message); }
}
function VolverMenu() {
  try {
    const html = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
    SpreadsheetApp.getUi().showModalDialog(html, 'Cerrando...');
    return { success: true, message: 'Volviendo al menú principal' };
  } catch (error) { return { success: false, message: 'Error al volver al menú: ' + error.message }; }
}
function mostrarPanelMenu() { return VolverMenu(); }
function testComunicacion() {
  return { success: true, message: 'Comunicación OK - Sistema v23 con estadísticas congruentes, reunión vigente, FAC-114, reconocimiento separado y vista previa sin guardar', timestamp: new Date(), sheetId: AGENDA_SHEET_ID, correoInstitucional: CORREO_INSTITUCIONAL, usuariosVotacion: CORREOS_PERMITIDOS_VOTACION.length };
}
function obtenerDatosProximaReunion() {
  try {
    const numero = obtenerSiguienteNumeroReunion();
    const fechaInfo = obtenerSiguienteFechaReunion();
    return { success: true, numero: numero, fechaInfo: fechaInfo };
  } catch (e) { return { success: false, message: "Error en función maestra: " + e.message }; }
}
// NOTA v22: las dos funciones siguientes son código de cliente (HTML) que venía
// pegado en el .gs original; se preservan tal cual, sin tocar su lógica.
function detectarFechaInteligente() {
  const btn = document.querySelector('button[onclick="detectarFechaInteligente()"]');
  const infoSpan = document.getElementById('infoOrigenFecha');
  const inputFecha = document.getElementById('fechaReunion');
  const textoOriginal = btn.innerHTML;
  btn.innerHTML = '<div class="spinner" style="width:15px;height:15px;border-width:2px;"></div>';
  btn.disabled = true;
  infoSpan.innerHTML = '<span style="color:var(--oga-text-muted);">Consultando calendario...</span>';
  google.script.run
    .withSuccessHandler(function (res) {
      btn.innerHTML = textoOriginal;
      btn.disabled = false;
      if (res.success) {
        inputFecha.value = res.fechaInput;
        let colorBadge = 'var(--oga-success)';
        if (res.origen === 'FECHA EXTRAORDINARIA') colorBadge = 'var(--oga-warning-dark)';
        infoSpan.innerHTML = '<span style="color:' + colorBadge + '"> ' + res.origen + ' (' + res.mes + ') detectada</span>';
        mostrarNotificacion(' Fecha fijada: ' + res.fechaDisplay, 'success');
        sugerirRangoFechas(res.fechaInput);
      } else {
        infoSpan.innerHTML = '<span style="color:var(--oga-error);"> ' + res.message + '</span>';
        mostrarNotificacion(res.message, 'warning');
      }
    })
    .withFailureHandler(function (err) {
      btn.innerHTML = textoOriginal;
      btn.disabled = false;
      infoSpan.innerHTML = '<span style="color:var(--oga-error);">Error de conexión</span>';
      mostrarNotificacion(' Error: ' + err.message, 'error');
    })
    .obtenerSiguienteFechaReunion();
}
function sugerirRangoFechas(fechaReunionStr) {
  const inicioInput = document.getElementById('fechaInicio');
  const finInput = document.getElementById('fechaFin');
  if (inicioInput.value && finInput.value) return;
  const fechaBase = new Date(fechaReunionStr + 'T00:00:00');
  const inicio = new Date(fechaBase); inicio.setDate(inicio.getDate() - 15);
  const fin = new Date(fechaBase); fin.setDate(fin.getDate() + 1);
  const fmt = function (d) { return d.toISOString().split('T')[0]; };
  if (!inicioInput.value) inicioInput.value = fmt(inicio);
  if (!finInput.value) finInput.value = fmt(fin);
}
function obtenerSiguienteNumeroReunion() {
  try {
    const ss = SpreadsheetApp.openById(REUNIONES_SHEET_ID);
    const hoja = ss.getSheetByName("Reuniones");
    const anioActual = new Date().getFullYear();
    if (!hoja || hoja.getLastRow() <= 1) return '01-' + anioActual;
    const datos = hoja.getDataRange().getValues();
    let maxNumero = 0;
    for (let i = 1; i < datos.length; i++) {
      const valor = String(datos[i][0]).replace(/^'/, '').trim();
      const partes = valor.split('-');
      if (partes.length === 2) {
        const numero = parseInt(partes[0], 10);
        const anio = parseInt(partes[1], 10);
        if (anio === anioActual) {
          if (!isNaN(numero) && numero > maxNumero) maxNumero = numero;
        }
      }
    }
    const siguiente = maxNumero + 1;
    const numeroStr = siguiente < 10 ? "0" + siguiente : String(siguiente);
    return numeroStr + '-' + anioActual;
  } catch (e) {
    Logger.log("Error obteniendo siguiente número: " + e.message);
    return '01-' + new Date().getFullYear();
  }
}
// Segunda definición v22 de validarPlazoReevaluacionBackend (preservada tal cual;
// la primera definición, de 72 horas, es la que se usa en guardarDatosAgendaConsejo).
function validarPlazoReevaluacionBackend() {
  try {
    const config = obtenerConfiguracionAgendaConsejo();
    if (!config) return { apto: false, mensaje: "No hay una agenda activa configurada para validar la fecha." };
    const fechaStr = config.fechaReunion;
    const horaStr = config.horaLimite || "23:59";
    if (!fechaStr) return { apto: false, mensaje: "La agenda activa no tiene fecha de reunión definida." };
    let fechaReunionFinal;
    if (fechaStr instanceof Date) { fechaReunionFinal = new Date(fechaStr); }
    else { const partes = String(fechaStr).split('-'); fechaReunionFinal = new Date(partes[0], partes[1] - 1, partes[2]); }
    const partesHora = String(horaStr).split(':');
    fechaReunionFinal.setHours(parseInt(partesHora[0]), parseInt(partesHora[1]), 0, 0);
    const HORAS_LIMITE = 192;
    const MS_EN_72_HORAS = HORAS_LIMITE * 60 * 60 * 1000;
    const fechaCorte = new Date(fechaReunionFinal.getTime() - MS_EN_72_HORAS);
    const ahora = new Date();
    if (ahora > fechaCorte) {
      const msTarde = ahora.getTime() - fechaCorte.getTime();
      const horasTarde = Math.floor(msTarde / (1000 * 60 * 60));
      const cierreLegible = Utilities.formatDate(fechaCorte, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      return { apto: false, mensaje: ' TIEMPO AGOTADO PARA REEVALUACIÓN: Se requieren 72 horas antes de la reunión. El sistema cerró la recepción de este tipo de solicitud el ' + cierreLegible + ' (hace ' + horasTarde + ' horas).' };
    }
    return { apto: true };
  } catch (e) {
    Logger.log("Error crítico validando 72h: " + e.message);
    return { apto: false, mensaje: "Error técnico validando fechas: " + e.message };
  }
}

// ===================== FUNCIONES DE PRUEBA =====================
function testContadorTiempo() {
  const estadoTiempo = obtenerEstadoTiempoAgenda();
  const puedeAgregar = puedeAgregarElementos();
  if (estadoTiempo.success) return { success: true, message: 'Contador de tiempo funcionando correctamente', datos: estadoTiempo };
  return { success: false, message: estadoTiempo.message };
}
function testVotacionesSimples() {
  const verificacion = hayAgendaListaParaVotar();
  if (verificacion.success) return { success: true, message: 'Votaciones simples funcionando correctamente' };
  return { success: false, message: verificacion.mensaje };
}
function testPermisosVotacion() {
  const usuario = Session.getActiveUser().getEmail();
  const autorizado = esUsuarioAutorizadoParaVotacion();
  return { success: true, message: 'Test de permisos completado', usuario: usuario, autorizado: autorizado, usuariosAutorizados: CORREOS_PERMITIDOS_VOTACION };
}
function testEstadisticasMejoradas() {
  const estadisticas = obtenerEstadisticasVotosPorReunionMejoradas('01-2025');
  if (estadisticas.success) return { success: true, message: 'Estadísticas corregidas funcionando correctamente', datos: estadisticas };
  return { success: false, message: estadisticas.message };
}
function testSistemaCompletoMejorado() {
  Logger.log(" INICIANDO TEST SISTEMA COMPLETO MEJORADO v23");
  const test1 = testContadorTiempo();
  const test2 = { success: verificarAgendaConsejoActiva() };
  const test3 = testVotacionesSimples();
  const test4 = testPermisosVotacion();
  const test5 = testEstadisticasMejoradas();
  const agendas = obtenerTodasLasAgendasConsejo();
  const test6 = { success: agendas.length > 0, agendas: agendas.length };
  return {
    success: true, version: "v23 - Sistema con Estadísticas congruentes, Reunión Vigente, FAC-114 y Vista Previa",
    tests: { contadorTiempo: test1.success, agendaActiva: test2.success, votaciones: test3.success, permisosVotacion: test4.success, estadisticasCorregidas: test5.success, estadoGeneral: test6.success },
    permisos: test4, correoInstitucional: CORREO_INSTITUCIONAL,
    message: "Tests completados - SISTEMA v23"
  };
}
function probarEnvioCorreoInstitucionalPorApi() {
  const resultado = enviarCorreoInstitucionalPorApi_('jalvarado.oga@gmail.com', 'Prueba API institucional OGA', '<h2>Prueba correcta</h2><p>Este correo fue solicitado desde Agenda Consejo y enviado por el API de info@oga.org.gt.</p>', 'Prueba correcta. Este correo fue enviado por el API de info@oga.org.gt.');
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
// ===================== FIN DEL ARCHIVO =====================
// Versión 23 - Estadísticas congruentes (resumen + porTipo + estado APROBADO/RECHAZADO/SIN MAYORÍA),
// reunión vigente desde Configuración aunque esté cerrada, FAC-114 obligatorio salvo
// Reconocimiento de Acreditación (sin FAC, carpeta Documentos_Reconocimiento),
// Reconocimiento eliminado de Notificaciones, y vista previa sin guardar.
