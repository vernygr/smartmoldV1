// SCRIPT PARA GOOGLE APPS SCRIPT (script.google.com)
// Reemplaza el contenido anterior con este script
// Funciona para LECTURA y ESCRITURA de datos

const SHEET_ID = "1yublxJGpihDZ89LCgqP7N1o0Pi6bCMF_TesvZi0RSUI";
const SHEET_NAME_MOLD = "problemas_por_molde";

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MOLD);
    const data = JSON.parse(e.postData.contents);

    const row = [
      data.id_molde,
      data.fecha,
      data.problema,
      data.diagnostico,
      data.accion,
      data.estado,
      data.severidad,
      data.tecnico
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(
      JSON.stringify({success: true, message: "Datos guardados correctamente"})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(
      JSON.stringify({success: false, error: error.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const accion = e.parameter.accion || 'moldes';

    if (accion === 'moldes') {
      return getMoldes();
    } else if (accion === 'buscar_molde') {
      const molde = e.parameter.molde || '';
      return buscarMolde(molde);
    } else if (accion === 'stats') {
      return getStats();
    }

    return ContentService.createTextOutput(
      JSON.stringify({success: false, error: "Acción no reconocida"})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(
      JSON.stringify({success: false, error: error.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function getMoldes() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MOLD);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const moldes = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    moldes.push(obj);
  }

  // Agrupar por molde
  const grouped = {};
  moldes.forEach(m => {
    if (!grouped[m.id_molde]) {
      grouped[m.id_molde] = [];
    }
    grouped[m.id_molde].push(m);
  });

  return ContentService.createTextOutput(
    JSON.stringify({success: true, moldes: moldes, grouped: grouped})
  ).setMimeType(ContentService.MimeType.JSON);
}

function buscarMolde(moldeId) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MOLD);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const resultados = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    const moldActual = String(data[i][0]).toUpperCase();
    if (moldActual.includes(moldeId.toUpperCase())) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j];
      }
      resultados.push(obj);
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify({success: true, moldes: resultados})
  ).setMimeType(ContentService.MimeType.JSON);
}

function getStats() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_MOLD);
  const data = sheet.getDataRange().getValues();

  let stats = {
    total: 0,
    moldes_unicos: 0,
    estado: {},
    severidad: {}
  };

  const moldes_set = new Set();

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    stats.total++;
    moldes_set.add(String(data[i][0]));

    const estado = String(data[i][5] || '');
    const severidad = String(data[i][6] || '');

    stats.estado[estado] = (stats.estado[estado] || 0) + 1;
    stats.severidad[severidad] = (stats.severidad[severidad] || 0) + 1;
  }

  stats.moldes_unicos = moldes_set.size;

  return ContentService.createTextOutput(
    JSON.stringify({success: true, stats: stats})
  ).setMimeType(ContentService.MimeType.JSON);
}
