function sendEditedRowToN8n(e) {

  const webhookUrl = "";
  const range = e.range;
  const sheet = range.getSheet();

  const rowStart = range.getRow();
  const rowEnd = range.getLastRow();

  const columnStart = range.getColumn();
  const columnEnd = range.getLastColumn();

  const lastColumn = sheet.getLastColumn();

  // Full row values
  const rowValues = sheet
    .getRange(rowStart, 1, 1, lastColumn)
    .getValues();

  const data = {
    sheetId: sheet.getSheetId(),

    sheetName: sheet.getName(),

    rangeA1Notation: range.getA1Notation(),

    range: {
      columnEnd: columnEnd,
      columnStart: columnStart,
      rowEnd: rowEnd,
      rowStart: rowStart
    },

    oldValue: e.oldValue ?? null,

    value: e.value ?? null,

    user: {
      email: Session.getActiveUser().getEmail(),
      nickname: Session.getActiveUser().getEmail().split("@")[0]
    },

    rowValues: rowValues
  };

  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(data)
  });
}
