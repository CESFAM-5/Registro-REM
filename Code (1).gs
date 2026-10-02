/**
 * Registro de actividades grupales · CESFAM N°5 (prototipo)
 * Recibe los registros del formulario web y los guarda en esta planilla.
 *
 * Instalación:
 * 1. Abre la planilla en Google Sheets > Extensiones > Apps Script.
 * 2. Borra el contenido de Code.gs, pega este archivo y guarda.
 * 3. Implementar > Nueva implementación > tipo "Aplicación web".
 *    Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
 * 4. Copia la URL que termina en /exec y pégala en la constante ENDPOINT de index.html.
 *
 * Cada envío trae { registro, salidas }. Si el código ya existe (p. ej. una actividad
 * programada que ahora se registra como realizada), se actualiza la misma fila y se
 * reemplazan sus salidas REM. Así no hay duplicados.
 */

const HOJA_REGISTRO = "Registro";
const HOJA_SALIDAS = "Salidas_REM";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const r = data.registro;
    const salidas = data.salidas || [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mes = String(r.fecha).slice(0, 7);
    const tipos = {
      promo: "Promoción o educación grupal", vidasana: "Taller de Vida Sana",
      amigables: "Taller en Espacios Amigables", gestion: "Actividad de gestión",
      participacion: "Participación social", orientacion: "Orientación o educación en terreno"
    };

    // --- Registro: actualizar si el código existe, si no agregar ---
    const reg = ss.getSheetByName(HOJA_REGISTRO);
    const fila = [
      r.codigo, r.fecha, mes, r.estado, r.nombre, r.responsable || "", r.equipo || "",
      tipos[r.tipo] || r.tipo, r.sesiones || 0, r.h || 0, r.m || 0, (r.h || 0) + (r.m || 0),
      r.po || 0, r.mig || 0, r.prais || 0, r.tm || 0, r.tf || 0, r.nb || 0, r.nr || 0,
      r.meta || 0, r.obs || "", new Date()
    ];
    const codigos = reg.getRange(1, 1, Math.max(reg.getLastRow(), 1), 1).getValues().flat();
    let n = codigos.indexOf(r.codigo) + 1;
    if (n < 1) n = reg.getLastRow() + 1;
    reg.getRange(n, 1, 1, 3).setNumberFormat("@"); // código, fecha y mes como texto
    reg.getRange(n, 1, 1, fila.length).setValues([fila]);

    // --- Salidas_REM: borrar las anteriores de este código y escribir las nuevas ---
    const sal = ss.getSheetByName(HOJA_SALIDAS);
    const last = sal.getLastRow();
    if (last > 1) {
      const cods = sal.getRange(2, 1, last - 1, 1).getValues().flat();
      for (let i = cods.length - 1; i >= 0; i--) {
        if (cods[i] === r.codigo) sal.deleteRow(i + 2);
      }
    }
    if (salidas.length) {
      const filas = salidas.map(o => [r.codigo, mes, r.estado, o.hoja, o.seccion, o.celda, o.fila, o.columna, o.valor]);
      const start = sal.getLastRow() + 1;
      sal.getRange(start, 1, filas.length, 2).setNumberFormat("@");
      sal.getRange(start, 1, filas.length, filas[0].length).setValues(filas);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, codigo: r.codigo }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
