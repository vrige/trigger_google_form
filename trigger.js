function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var name_list_student_in_the_column = "Lista Studenti";
  var names = e.namedValues[name_list_student_in_the_column][0].split(",");

  var firstColumnHeader = e.values[1];
  var fieldNames = Object.keys(e.namedValues);
  // the entry columns are these: Informazioni cronologiche,	Nome del corso,	Docente del corso,	Data della lezione,	Durata della lezione,	Lista Studenti
  // we want to remove some columns for the new rows: nformazioni cronologiche,	Nome del corso and Lista Studenti
  var position_obj_to_remove = [0,1,5];

  //just to check the field names
  for (var i = 0; i < fieldNames.length; i++) {
    var fieldName = fieldNames[i];
    console.log("Field Name:", fieldName);
  }

  for (var i = 0; i < names.length; i++) {
    var row = [];
    for (var j = 0; j < e.values.length; j++) {
      if (position_obj_to_remove.includes(j)){
        continue;
      }

      row.push(e.values[j]);
    }
    student = names[i].trim().split("-");
    row.push(student[0].trim());
    row.push(student[1].trim());
    row.push(student[2].trim());
    //row[row.length] = names[i].trim();
    UpdateDocument(firstColumnHeader, row);
  }
}

function UpdateDocument(sheet_name, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheet_name);

  if (sheet) {
    // If the document exists, add the new rows to it
    sheet.appendRow(row);
  } else {
    ss.insertSheet(sheet_name);
    sheet = ss.getSheetByName(sheet_name);
    sheet.appendRow(row);
  }
}
