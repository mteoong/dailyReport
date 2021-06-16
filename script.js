function createTimebasedTrigger() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('dailyReports')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .onWeekDay(ScriptApp.WeekDay.TUESDAY)
        .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
        .onWeekDay(ScriptApp.WeekDay.THURSDAY)
        .onWeekDay(ScriptApp.WeekDay.FRIDAY)
        .atHour(4)
        .nearMinute(10)
        .inTimezone("Asia/Taipei")
        .create();
  }
  
  function dailyReports() {
    let today = new Date();
    let options = { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' };
    let date = today.toLocaleDateString("en-US", options);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
  
    for (let sheet of sheets) {
      ss.setActiveSheet(sheet);
  
      let info = sheet.getSheetValues(1, 1, 1, 3);
      let name = info[0][0];
      
      let email = info[0][1];
      let index = info[0][2]
      let doc = DocumentApp.create(`${name} Daily Report - ${date}`)
  
      //Format Document
      formatDoc(doc);
  
      //Send Document
      sheet.appendRow([date, doc.getUrl()]);
      cleanSheet(sheet, index);
      doc.addEditor(email);
      GmailApp.sendEmail(email, `${name} Daily Report - ${date}`, `Link: ${doc.getUrl()}`);
    }
  }
  
  function formatDoc(doc) {
    let body = doc.getBody();
  
    let cells = [
      ["Timestamp", "Work done"],
      ["9:00 - 12:00", "- "],
      ["1:00 - 3:00", "- "],
      ["3:00 - 5:00", "- "]
    ];
  
    let table = body.appendTable(cells);
    table.setColumnWidth(0, 112);
    table.setColumnWidth(1, 385);
  
    let bodyText = {};
    bodyText[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
    bodyText[DocumentApp.Attribute.FONT_SIZE] = 14;
  
    table.setAttributes(bodyText);
  
    let headText = {};
    headText[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
    headText[DocumentApp.Attribute.FONT_SIZE] = 16;
    headText[DocumentApp.Attribute.BOLD] = true;
    headText[DocumentApp.Attribute.WIDTH] = 10;
  
    let header = table.getRow(0);
    header.setAttributes(headText);
    header.getCell(0).setBackgroundColor('#BBB9B9');
     header.getCell(1).setBackgroundColor('#BBB9B9');
  }
  
  function cleanSheet(sheet, index) {
    if (sheet.getLastRow() - index > 14) {
      sheet.hideRow(index);
      index++;
    }
  }
  