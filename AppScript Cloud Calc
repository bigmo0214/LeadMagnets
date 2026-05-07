const SHEET_NAME = "FinOps Leads";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sheet  = ss.getSheetByName(SHEET_NAME);

    // Create sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Write headers on first run only
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Nombre",
        "Correo",
        "Empresa",
        "Cargo",
        "Factura Mensual (USD)",
        "Núm. Proveedores",
        "Última Auditoría FinOps",
        "Ahorro Est. Bajo / Mes (USD)",
        "Ahorro Est. Alto / Mes (USD)",
        "Ahorro Est. Bajo / Año (USD)",
        "Ahorro Est. Alto / Año (USD)",
        "Risk Score",
      ]);

      // Style header row
      const header = sheet.getRange(1, 1, 1, 13);
      header.setBackground("#ff5b00");
      header.setFontColor("#ffffff");
      header.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Append lead row (runs every submission)
    sheet.appendRow([
      new Date(data.timestamp),
      data.name           || "",
      data.email          || "",
      data.company        || "",
      data.role           || "",
      data.monthlyBill    || 0,
      data.providers      || "",
      data.auditRecency   || "",
      data.lowMonthly     || 0,
      data.highMonthly    || 0,
      data.lowAnnual      || 0,
      data.highAnnual     || 0,
      data.risk           || "",
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
