/**
 * Create goolge sheet custom menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("其他功能")
    .addItem("過濾學年行事曆PDF必要資料", "filter_converted_file")
    .addItem("登錄到Goolge日曆", "push_to_google_calendar")
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu("從日曆中刪除此學年所有事件")
        .addItem("確定", "clean_all_events")
    )
    .addSeparator()
    .addItem('顯示操作說明', "show_instruction")
    .addToUi();
}


function show_instruction() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("instruction").setTitle("使用說明");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

/**
 * Convert PDF to GOOGLE DOC
 * @returns {document_id}
 */
function pdf_to_doc() {
  const ss = SpreadsheetApp.getActiveSheet();
  const ssID = ss.getParent().getId();
  const folderID = DriveApp.getFileById(ssID).getParents().next().getId();
  const folder = DriveApp.getFolderById(folderID);
  const files = folder.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    let splitName = file.getName().split('.');
    if (splitName[splitName.length - 1] == 'pdf') {
      let pdfID = file.getId();
      let fileBlob = DriveApp.getFileById(pdfID).getBlob();
      const resource = {
        title: fileBlob.getName(),
        mimeType: fileBlob.getContentType(),
        parents: [{ id: folderID }]
      };
      const options = {
        ocr: true
      };
      let doc = Drive.Files.insert(resource, fileBlob, options);
      Drive.Files.remove(pdfID);
      return doc.getId();
    }
  }
}

/**
 * onvert PDF to google DOC and filter necessary information.
 * @returns {event_information}
 */
function filter_converted_file() {
  const docID = pdf_to_doc();
  const doc = DocumentApp.openById(docID);
  const body = doc.getBody().getParagraphs();
  const arr_months = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二"];

  // Filter all valid data.
  let arr_bodyText = body
    .map(value => {
      return value.getText().split(" ").join("");
    })
    .filter((value) => value != "");

  // Get all necessary information.
  let arr_necessaryInformation = [];
  for (let i = 0; i < arr_bodyText.length; i++) {
    if (arr_bodyText[i].charAt(0) === '(') {
      let events = arr_bodyText[i].split('、(');
      let eventLength = events.length;
      if (eventLength > 1) {
        events.forEach(element => {
          if (element.charAt(0) !== '(') {
            element = '(' + element;
          }
          arr_necessaryInformation.push(element);
        });
        continue;
      }
      arr_necessaryInformation.push(arr_bodyText[i]);
      continue;
    }
    for (let j = 0; j < 12; j++) {
      if (arr_bodyText[i] + arr_bodyText[i + 1] === arr_months[j + 1]) {
        let month = convert_month_type(arr_bodyText[i++] + arr_bodyText[i]);
        arr_necessaryInformation.push(month);
        break;
      }
      if (arr_bodyText[i] + arr_bodyText[i + 1] === arr_months[j + 2]) {
        let month = convert_month_type(arr_bodyText[i++] + arr_bodyText[i]);
        arr_necessaryInformation.push(month);
        break;
      }
      if (arr_bodyText[i] === arr_months[j]) {
        let month = convert_month_type(arr_bodyText[i]);
        arr_necessaryInformation.push(month);
        break;
      }
    }
  }

  let sheetData = [];
  for (let i = 0, month = 8; i < arr_necessaryInformation.length; i++) {
    // Data is month.
    if (typeof (arr_necessaryInformation[i]) === 'number') {
      month = arr_necessaryInformation[i];
      continue;
    }
    // Data is an event.
    let day = arr_necessaryInformation[i].split(')')[0];
    let rowData = [
      month,
      day.slice(1),
      '',
      arr_necessaryInformation[i].slice((day.length + 1)),
      ''
    ];
    sheetData.push(rowData);
  }

  // Write data into google sheet.
  const ss = SpreadsheetApp.getActiveSheet();
  ss.deleteRows(2, ss.getLastRow() - 1);
  let formats = [];
  let alignments = [];
  for (let row = 0; row < ss.getLastRow() + 50; row++) {
    formats.push(["@", "@"]);
    alignments.push(["right", "right"]);
  }
  ss.getRange(2, 2, ss.getLastRow() + 50, 2).setNumberFormats(formats);
  ss.getRange(2, 1, sheetData.length, 5).setValues(sheetData);
  ss.getRange(2, 2, ss.getLastRow() + 50, 2).setHorizontalAlignments(alignments);
}

/**
 * Convert chinese month to the corresponding number.
 * @param {string} chinese_month 
 * @returns {number_month}
 */
function convert_month_type(chinese_month) {
  switch (chinese_month) {
    case '一':
      return 1;
    case '二':
      return 2;
    case '三':
      return 3;
    case '四':
      return 4;
    case '五':
      return 5;
    case '六':
      return 6;
    case '七':
      return 7;
    case '八':
      return 8;
    case '九':
      return 9;
    case '十':
      return 10;
    case '十一':
      return 11;
    case '十二':
      return 12;
    default:
      return -1;
  }
}

/**
 * Get sheet data & create events.
 */
function push_to_google_calendar() {
  const calendar = CalendarApp.getCalendarById('CALENDAR_ID');
  const ss = SpreadsheetApp.getActiveSheet();
  const sheetData = ss.getDataRange().getValues().slice(1);
  const obj_rowData = sheetData.map(arr => {
    return {
      month: arr[0],
      day: arr[1],
      time: arr[2],
      event: arr[3],
      description: arr[4],
    };
  });
  console.log(obj_rowData);
  obj_rowData.forEach(obj => {
    createSingleEvent(calendar, obj);
  });
}

/**
 * Create a calendar event.
 * @param {Calendar} calendar An object.
 * @param {object} obj A row data in the sheet.
 */
function createSingleEvent(calendar, obj) {
  const isInOneDay = obj.day.split('-').length === 1 ? true : false;
  const isSpecifiedTime = obj.time !== '' ? true : false;
  const year = new Date().getFullYear();

  // The Event is only in one day.
  if (isInOneDay) {
    if (isSpecifiedTime) {
      const [startTime, endTime] = obj.time.split('-');
      const startDateTime = 12 - obj.month <= 4 ?
        new Date([[year, obj.month, obj.day].join('/'), startTime].join(' ')) :
        new Date([[year + 1, obj.month, obj.day].join('/'), startTime].join(' '));
      const endDateTime = 12 - obj.month <= 4 ?
        new Date([[year, obj.month, obj.day].join('/'), endTime].join(' ')) :
        new Date([[year + 1, obj.month, obj.day].join('/'), endTime].join(' '));
      calendar.createEvent(obj.event, startDateTime, endDateTime, { description: obj.description });
      return;
    }
    const date = 12 - obj.month <= 4 ?
      new Date([year, obj.month, obj.day].join('/')) :
      new Date([year + 1, obj.month, obj.day].join('/'));
    calendar.createAllDayEvent(obj.event, date, { description: obj.description });
    return;
  }

  // The Event lasts more than one day.
  let [startDay, endDay] = obj.day.split('-');
  startDay = [obj.month, startDay].join('/');
  endDay = endDay.split('/').length > 1 ? endDay : [obj.month, endDay].join('/');
  if (isSpecifiedTime) {
    const [startTime, endTime] = obj.time.split('-');
    const startDateTime = 12 - obj.month <= 4 ?
      new Date([[year, startDay].join('/'), startTime].join(' ')) :
      new Date([[year + 1, startDay].join('/'), startTime].join(' '));
    const endDateTime = 12 - parseInt(endDay.split('/')[0]) <= 4 ?
      new Date([[year, endDay].join('/'), endTime].join(' ')) :
      new Date([[year + 1, endDay].join('/'), endTime].join(' '));
    calendar.createEvent(obj.event, startDateTime, endDateTime, { description: obj.description });
    return;
  }
  const startDate = 12 - obj.month <= 4 ?
    new Date([year, startDay].join('/')) :
    new Date([year + 1, startDay].join('/'));
  const endDate = 12 - parseInt(endDay.split('/')[0]) <= 4 ?
    new Date([year, endDay].join('/')) :
    new Date([year + 1, endDay].join('/'));
  calendar.createAllDayEvent(obj.event, startDate, endDate, { description: obj.description });
}
