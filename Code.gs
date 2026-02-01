function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Informasi Sertifikasi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * GOOGLE SHEETS CONNECTION SETUP
 */
var SPREADSHEET_ID = '1-DvFmX-Saq4RsyFC-XFNFvqRlyqWIn8SzxL5v7U4cHQ';

/* --- SINGLE SHEET DATABASE MODEL --- */
var MAIN_SHEET_NAME = 'Perencanaan';

/* --- TRIGGER FOR MANUAL ENTRY --- */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "Perencanaan") return;
  
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  if (row <= 1 || range.getNumRows() > 1) return;
  
  var idCell = sheet.getRange(row, 1);
  if (idCell.getValue() === "") {
    var isLat = (col >= 11 && col <= 15);
    var prefix = isLat ? "lat-" : "cert-";
    idCell.setValue(prefix + new Date().getTime());
  }
}

function connect(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(MAIN_SHEET_NAME);
    var headers = [
      'ID', 'SAP', 'NAMA', 
      'CERT_ITEM_ID', 'CERT_JUDUL', 'CERT_PERIODE', 'CERT_ANGGARAN', 'CERT_STATUS', 'CERT_MANDATORY', 'CERT_RESIKO',
      'LAT_ITEM_ID', 'LAT_JUDUL', 'LAT_INSTRUKTUR', 'LAT_PERIODE', 'LAT_RESIKO',
      'GENDER', 'UNIT_KERJA'
    ];
    sheet.appendRow(headers);
  }
  return sheet;
}

function getData() {
  try {
    var sheet = connect(MAIN_SHEET_NAME);
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if(!r[0]) continue;

        var idStr = r[0].toString();
        if (idStr.indexOf('lat-') === -1) { 
            data.push({
                id: r[0],
                sap: r[1],
                nama: r[2],
                itemId: r[3],
                judul: r[4],
                periode: r[5],
                jumlah: r[6],
                statusAnggaran: r[7],
                mandatory: r[8],
                resiko: r[9],
                gender: r[15],
                posCode: r[16],
                type: 'cert'
            });
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getData: ' + e.message);
    throw e;
  }
}

function getLATData() {
  try {
     var sheet = connect(MAIN_SHEET_NAME);
    var rows = sheet.getDataRange().getValues();
    var data = [];
    
    for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if(!r[0]) continue;

         var idStr = r[0].toString();
        if (idStr.indexOf('lat-') === 0) {
             data.push({
                id: r[0],
                sap: r[1],
                nama: r[2],
                itemId: r[10],
                judul: r[11],
                instruktur: r[12],
                periode: r[13],
                resiko: r[14],
                gender: r[15],
                posCode: r[16],
                type: 'lat'
            });
        }
    }
    return data;
  } catch (e) {
    Logger.log('ERROR getLATData: ' + e.message);
     throw e;
  }
}

function addData(formObject) {
  try {
    var sheet = connect(MAIN_SHEET_NAME);
    var id = 'cert-' + new Date().getTime();
    
    var newRow = [
        id, formObject.sap, formObject.nama,
        formObject.itemId, formObject.judul, formObject.periode, formObject.jumlah, formObject.statusAnggaran, formObject.mandatory, formObject.resiko,
        "", "", "", "", "",
        formObject.gender, formObject.posCode
    ];
    
    sheet.appendRow(newRow);
    
    // OPTIMIZATION: Return only the new object
    return {
        id: id,
        sap: formObject.sap,
        nama: formObject.nama,
        itemId: formObject.itemId,
        judul: formObject.judul,
        periode: formObject.periode,
        jumlah: formObject.jumlah,
        statusAnggaran: formObject.statusAnggaran,
        mandatory: formObject.mandatory,
        resiko: formObject.resiko,
        gender: formObject.gender,
        posCode: formObject.posCode,
        type: 'cert'
    };
  } catch(e) { throw e; }
}

function updateData(formObject) {
    var sheet = connect(MAIN_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    
    for(var i=1; i<data.length; i++) {
        if(data[i][0].toString() === formObject.id.toString()) {
            var r = i+1;
             var updatedRow = [
                formObject.id, formObject.sap, formObject.nama,
                formObject.itemId, formObject.judul, formObject.periode, formObject.jumlah, formObject.statusAnggaran, formObject.mandatory, formObject.resiko,
                "", "", "", "", "",
                formObject.gender, formObject.posCode
            ];
            sheet.getRange(r, 1, 1, updatedRow.length).setValues([updatedRow]);
            break;
        }
    }
    
    // OPTIMIZATION: Return only the updated object
    return {
        id: formObject.id,
        sap: formObject.sap,
        nama: formObject.nama,
        itemId: formObject.itemId,
        judul: formObject.judul,
        periode: formObject.periode,
        jumlah: formObject.jumlah,
        statusAnggaran: formObject.statusAnggaran,
        mandatory: formObject.mandatory,
        resiko: formObject.resiko,
        gender: formObject.gender,
        posCode: formObject.posCode,
        type: 'cert'
    };
}

function deleteData(id) {
    var sheet = connect(MAIN_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
        if(data[i][0].toString() === id.toString()) {
            sheet.deleteRow(i+1);
            break;
        }
    }
    // OPTIMIZATION: Return only deleted ID
    return id;
}

function addLATData(formObject) {
    var sheet = connect(MAIN_SHEET_NAME);
    var id = 'lat-' + new Date().getTime();
    
    var newRow = [
        id, formObject.sap, formObject.nama,
        "", "", "", "", "", "", "",
        formObject.itemId, formObject.judul, formObject.instruktur, formObject.periode, formObject.resiko,
        formObject.gender, formObject.posCode
    ];
    sheet.appendRow(newRow);
    
    return {
        id: id,
        sap: formObject.sap,
        nama: formObject.nama,
        itemId: formObject.itemId,
        judul: formObject.judul,
        instruktur: formObject.instruktur,
        periode: formObject.periode,
        resiko: formObject.resiko,
        gender: formObject.gender,
        posCode: formObject.posCode,
        type: 'lat'
    };
}

function updateLATData(formObject) {
    var sheet = connect(MAIN_SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    
    for(var i=1; i<data.length; i++) {
        if(data[i][0].toString() === formObject.id.toString()) {
            var r = i+1;
            var updatedRow = [
                formObject.id, formObject.sap, formObject.nama,
                "", "", "", "", "", "", "",
                formObject.itemId, formObject.judul, formObject.instruktur, formObject.periode, formObject.resiko,
                formObject.gender, formObject.posCode
            ];
            sheet.getRange(r, 1, 1, updatedRow.length).setValues([updatedRow]);
            break;
        }
    }
    
    return {
        id: formObject.id,
        sap: formObject.sap,
        nama: formObject.nama,
        itemId: formObject.itemId,
        judul: formObject.judul,
        instruktur: formObject.instruktur,
        periode: formObject.periode,
        resiko: formObject.resiko,
        gender: formObject.gender,
        posCode: formObject.posCode,
        type: 'lat'
    };
}

function deleteLATData(id) {
    var sheet = connect(MAIN_SHEET_NAME);
     var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
        if(data[i][0].toString() === id.toString()) {
            sheet.deleteRow(i+1);
            break;
        }
    }
    return id;
}