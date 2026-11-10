# Agenda Consejo OGA — v23: qué cambió y dónde copiarlo

**Versión:** v22 → v23 · **Fecha:** 10 de noviembre de 2026
**Repo:** [github.com/Jhons2004/TramitesOGA](https://github.com/Jhons2004/TramitesOGA) · Carpeta: `agenda-consejo/`
**Archivos del sistema:** `Agenda_Sistema.html` (frontend, se pega en el archivo HTML del proyecto Apps Script) y `Agenda_Consejo.gs` (backend, se pega en el archivo `.gs`).

Guía completa en HTML: [`LANDING_CAMBIOS_v23.html`](https://github.com/Jhons2004/TramitesOGA/blob/main/agenda-consejo/LANDING_CAMBIOS_v23.html)

## Resumen en 6 líneas

1. **Estadísticas congruentes:** resumen, barra de consenso y tabla por tipo de decisión, con estado APROBADO / RECHAZADO / SIN MAYORÍA y detalle de votantes.
2. **Filtro por última reunión vigente:** Ver Datos y Estadísticas arrancan siempre en la reunión vigente según `Configuración`, aunque esté cerrada, aprobada o vencida.
3. **Reconocimiento de Acreditación separado:** ya no exige FAC-014/078/003/090 ni FAC-114; solo documentación sin nombre estándar.
4. **Nuevo OGA-FAC-114:** 5.º documento obligatorio para todas las demás decisiones; el nombre del archivo debe iniciar con `OGA-FAC-114`.
5. **Notificaciones sin Reconocimiento:** se quitó "Reconocimiento de acreditación" del tipo de notificación y el backend lo rechaza si llega.
6. **Vista previa sin guardar:** el Generador de Documentos muestra la agenda construida en pantalla sin guardar ni registrar nada en Drive.

## Tabla de cambios

| # | Cambio | Dónde en el HTML | Dónde en el .gs | Qué copiar/pegar |
|---|--------|------------------|-----------------|------------------|
| 1 | Estadísticas más congruentes, claras y legibles | `mostrarEstadisticasDetalladasMejoradas(datos, targetId)`; contenedores `#contenedorPrincipalEstadisticas` y `#contenidoEstadisticas`; barra `.consenso-bar-container` | `obtenerEstadisticasVotosPorReunionMejoradas(numeroReunion)` (clave única registro + tipo de decisión; `porcentajeAFavor`, `puedeGenerarResolucion`); wrapper `obtenerEstadisticasVotosPorReunion(numeroReunion)` | Archivos completos `Agenda_Sistema.html` + `Agenda_Consejo.gs` |
| 2 | Filtro siempre por la última reunión vigente | Variable `reunionVigenteActual` + badge de reunión vigente; `cargarDatos()` y `cargarEstadisticasReunionEspecifica()` arrancan en esa reunión (`#filtroReunion`, `#selectReunionEstadisticas`) | `obtenerReunionVigenteParaFiltros()` — lee `Configuración` fila 2, columna B; vale aunque la agenda esté cerrada, aprobada o vencida | Archivos completos `Agenda_Sistema.html` + `Agenda_Consejo.gs` |
| 3 | Reconocimiento de Acreditación separado | Sección `#seccionDocsReconocimiento` (visible solo si `#tipoDecision` = "Reconocimiento de Acreditación"); `mostrarCamposEspecificos()` alterna las secciones; `guardarElemento()` salta la validación de FAC | `obtenerDocumentosObligatoriosPorTipoDecision_(tipoDecision)` — lista vacía para reconocimiento; `guardarDatosAgendaConsejo(datos)` guarda en subcarpeta `Documentos_Reconocimiento` sin validar prefijos | Archivos completos `Agenda_Sistema.html` + `Agenda_Consejo.gs` |
| 4 | Nuevo documento obligatorio OGA-FAC-114 | Tarjeta `fac114`: `#dropZoneFac114`, `#adjuntoFac114`, `#listaFac114`, `#estadoFac114`; entra en `DOCUMENTOS_DECISION_REQUERIDOS` con `codigo: 'OGA-FAC-114'` y en `archivosDecisionSeleccionados.fac114`; lo valida `validarDocumentosDecisionRequeridos()` | Constante `DOCUMENTOS_DECISION_OBLIGATORIOS` incluye `'OGA-FAC-114'`; `obtenerDocumentosObligatoriosPorTipoDecision_(tipoDecision)` devuelve los 5 documentos salvo reconocimiento; `validarDocumentosDecisionObligatorios_()` y `obtenerTipoDocumentoDecision_()` reconocen el prefijo `OGA-FAC-114` | Archivos completos `Agenda_Sistema.html` + `Agenda_Consejo.gs` |
| 5 | "Reconocimiento de acreditación" eliminado de Notificaciones | Select `#tipoNotificacion` en `#camposNotificacion`: se eliminó `<option value="Reconocimiento de acreditación">` | `guardarDatosAgendaConsejo(datos)`: si `tipoNotificacion` es "Reconocimiento de acreditación", devuelve error y no guarda | Archivos completos `Agenda_Sistema.html` + `Agenda_Consejo.gs` |
| 6 | Vista previa del documento sin guardar | Botón `#btnVistaPreviaAgenda` en la pestaña Documentos; `verVistaPreviaAgenda()` muestra el resultado en el modal `#modalConfirmacion` | `generarVistaPreviaAgendaConsejo()` — construye la agenda en memoria; NO llama a `DriveApp`, `makeCopy`, `almacenarAgenda()` ni `desactivarAgendaAlmacenada()` | Archivos completos `Agenda_Sistema.html` + `Agenda_Consejo.gs` |

## Archivos en el repo

| Archivo | Qué es |
|---------|--------|
| [`agenda-consejo/Agenda_Sistema.html`](https://github.com/Jhons2004/TramitesOGA/blob/main/agenda-consejo/Agenda_Sistema.html) | Frontend v23 (interfaz completa) |
| [`agenda-consejo/Agenda_Consejo.gs`](https://github.com/Jhons2004/TramitesOGA/blob/main/agenda-consejo/Agenda_Consejo.gs) | Backend v23 (lógica Apps Script) |
| [`agenda-consejo/LANDING_CAMBIOS_v23.html`](https://github.com/Jhons2004/TramitesOGA/blob/main/agenda-consejo/LANDING_CAMBIOS_v23.html) | Esta guía en formato página web |
| [`agenda-consejo/README.md`](https://github.com/Jhons2004/TramitesOGA/blob/main/agenda-consejo/README.md) | Este archivo |

## Cómo instalar (copiar y pegar) — 3 pasos

1. **Descargar del repo:** en [github.com/Jhons2004/TramitesOGA/tree/main/agenda-consejo](https://github.com/Jhons2004/TramitesOGA/tree/main/agenda-consejo), abrir `Agenda_Sistema.html` y `Agenda_Consejo.gs` y copiar el contenido completo de cada uno (botón *Raw* → Ctrl+A → Ctrl+C).
2. **Pegar en Apps Script:** en el proyecto de Google Apps Script, pegar `Agenda_Sistema.html` en el archivo HTML (reemplazo completo) y `Agenda_Consejo.gs` en el archivo `.gs` (reemplazo completo). Los dos archivos van juntos: no mezclar el HTML v23 con el .gs v22, porque las funciones nuevas (`obtenerReunionVigenteParaFiltros`, `generarVistaPreviaAgendaConsejo`, `obtenerDocumentosObligatoriosPorTipoDecision_`) no existirían en el backend viejo.
3. **Verificar:** guardar, desplegar la nueva versión y ejecutar `testComunicacion()`. Debe responder con éxito. Luego probar cada cambio según la sección "Cómo probarlo" de cada tarjeta en la landing (empezando por el cambio 4, FAC-114, y el cambio 6, vista previa).

> **Importante:** actualización por reemplazo completo de los dos archivos. No hay parches parciales: si se pega solo uno de los dos, las funciones nuevas del otro lado no existirán y el sistema fallará.

## Qué NO cambió

- **Reevaluación — regla de 72 horas:** sigue vigente (`validarPlazoReevaluacionBackend()`).
- **Calendario de reuniones:** detección de la próxima fecha (`obtenerSiguienteFechaReunion()`, `obtenerDatosProximaReunion()`), fechas extraordinarias y avance ACTIVA → CERRADA, iguales.
- **Votaciones:** envío desde `info@oga.org.gt`, usuarios autorizados (`esUsuarioAutorizadoParaVotacion()`) y hoja `Agenda_Votacion`, sin cambios.
- **Permisos:** quién crea agendas (`CORREOS_PERMITIDOS_CREAR`) y quién vota (`CORREOS_PERMITIDOS_VOTACION`), iguales.
- **Documentos FAC-014 / 078 / 003 / 090:** siguen obligatorios para todas las decisiones excepto Reconocimiento de Acreditación; solo se suma el FAC-114 como 5.º documento.
- **Resoluciones:** correlativo anual (`obtenerSiguienteResolucion()`) y registro en la hoja `Resoluciones`, sin cambios.
