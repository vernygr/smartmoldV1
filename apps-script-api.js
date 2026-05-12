// Este código va en Google Apps Script (script.google.com)
// 1. Copia este código en script.google.com
// 2. Reemplaza SHEET_ID con tu ID de Google Sheets
// 3. Deploy como Web App

const SHEET_ID = "1yublxJGpihDZ89LCgqP7N1o0Pi6bCMF_TesvZi0RSUI";

function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === "getMoldeProblems") {
      return getMoldeProblems();
    } else if (action === "getStats") {
      return getStats();
    }

    return ContentService.createTextOutput(
      JSON.stringify({success: false, error: "Invalid action"})
    ).setMimeType(ContentService.MimeType.JSON);

  } catch(error) {
    return ContentService.createTextOutput(
      JSON.stringify({success: false, error: error.toString()})
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function getMoldeProblems() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("problemas_por_molde");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Convert to array of objects
  const problems = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue; // Skip empty rows

    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    problems.push(obj);
  }

  // Sort by date descending
  problems.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

  return ContentService.createTextOutput(
    JSON.stringify({
      success: true,
      data: problems,
      total: problems.length
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

function getStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("problemas_por_molde");

  const data = sheet.getDataRange().getValues();

  let stats = {
    total: 0,
    pendiente: 0,
    en_proceso: 0,
    resuelto: 0,
    critica: 0,
    alta: 0,
    media: 0,
    baja: 0,
    moldes_unicos: new Set()
  };

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    stats.total++;
    const estado = data[i][5]; // Column F: estado
    const severidad = data[i][6]; // Column G: severidad
    const molde = data[i][0]; // Column A: id_molde

    if (estado) stats[estado]++;
    if (severidad) stats[severidad]++;
    if (molde) stats.moldes_unicos.add(molde);
  }

  stats.moldes_unicos = stats.moldes_unicos.size;

  return ContentService.createTextOutput(
    JSON.stringify({
      success: true,
      data: stats
    })
  ).setMimeType(ContentService.MimeType.JSON);
}
