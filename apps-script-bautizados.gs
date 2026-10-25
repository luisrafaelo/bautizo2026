/**
 * Apps Script — Consulta de certificado (lista de bautizados)
 *
 * Hoja esperada (fila 1 = encabezados), el orden de columnas no importa:
 *   RU | NOMBRE | APELLIDO | CERTIFICADO
 * En CERTIFICADO escribe SI / NO (o déjalo vacío si no tiene).
 *
 * Implementar como: Aplicación web
 *   - Ejecutar como: yo
 *   - Quién tiene acceso: cualquier persona
 * Después de editar, crea una "Nueva versión" de la implementación.
 */

const NOMBRE_HOJA = 'Bautizados'; // cambia por el nombre real de la pestaña

function doGet(e) {
  const ru = String((e.parameter && e.parameter.ru) || '').trim();

  if (!ru) return json({ error: 'Falta el RU', status: 400 });

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  if (!hoja) return json({ error: 'No existe la hoja "' + NOMBRE_HOJA + '"', status: 500 });

  const datos = hoja.getDataRange().getValues();
  const cab = datos[0].map(c => String(c).trim().toUpperCase());

  const iRU   = cab.indexOf('RU');
  const iNom  = cab.indexOf('NOMBRE');
  const iApe  = cab.indexOf('APELLIDO');
  const iCert = cab.indexOf('CERTIFICADO');

  if (iRU === -1 || iCert === -1) {
    return json({ error: 'Faltan las columnas RU o CERTIFICADO', status: 500 });
  }

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][iRU]).trim() === ru) {
      return json({
        ru: ru,
        nombre:      iNom !== -1 ? datos[i][iNom] : '',
        apellido:    iApe !== -1 ? datos[i][iApe] : '',
        certificado: String(datos[i][iCert]).trim().toUpperCase() // "SI" / "NO" / ""
      });
    }
  }

  return json({ error: 'RU no encontrado', status: 404 });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
