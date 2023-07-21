// function to call to compute all the statistics of the sheets in the list of "Lista materie"
function global_stitics(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main_subject_sheet_obj = getSheetByItsName(ss, "Lista materie");
  var main_subject_sheet = main_subject_sheet_obj["sheet"];

  var total_sum_students = 0;
  var last_row = 0;

  // check if the sheet exist
  if (main_subject_sheet_obj["abort"] === false){

    // avoid the first line (title of the list)
    var subjects = main_subject_sheet.getRange("A1:A" + main_subject_sheet.getLastRow()).getValues();

    for (var i = 1; i < subjects.length; i++) {
      
      var subject_sheet_obj = getSheetByItsName(ss, subjects[i][0]);

      if (subject_sheet_obj["abort"] === false){
        
        var subject_sheet = subject_sheet_obj["sheet"];
        var number_student_sub = compute_statitics(subject_sheet);

        var cell = main_subject_sheet.getRange(i + 1,2,1,4);
        //cell.setValue(number_student_sub);
        cell.setValues([number_student_sub])
        cell.setHorizontalAlignment("center");

        total_sum_students = total_sum_students + parseInt(number_student_sub[0]);
        last_row = i + 1;
      }
    }

    console.log("the last row is: ", last_row +1);
    console.log("total_sum_students: ", total_sum_students);

    var cell = main_subject_sheet.getRange(subjects.length + 1,2,1,1);
    cell.setValue(total_sum_students);
    cell.setHorizontalAlignment("center");
    cell.setFontWeight("bold");

    var cell = main_subject_sheet.getRange(2,2,1,4);
    var values = [
      ["Numero di studenti in totale","studenti partecipanti (>0h)","studenti che hanno superato il 70%","finanziamento (in euro):"]
    ];
    cell.setValues(values);

  }
}


// get the sheet with the name specified in "name_of_the_sheet"
function getSheetByItsName(ss, name_of_the_sheet){
  var abort = false;
  var error = "";
  subject_sheet = ss.getSheetByName(name_of_the_sheet);
  
  if (subject_sheet === null){
    abort = true;
    error = "Cannot find the sheet: " + name_of_the_sheet;
    console.log(error);
  }else{
    console.log("Sheet: " + name_of_the_sheet + " found");
  }

  return {abort: abort, error: error, name: name_of_the_sheet, sheet: subject_sheet};
}

// this function compute the statistics of the sheet in which you are (or the one specified in the parameter)
function compute_statitics(sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()) {
 
  var students = sheet.getRange("A1:A" + sheet.getLastRow()).getValues();
  
  console.log("Init");
  console.log("length: ", students.length);

  var count = 0;
  var check = false;
  var line_to_check = "";

  var count_students_school = 0;
  var count_students_with_70 = 0;
  var total_finanziamento_school = 0;
  var partecipanti_school = 0;

  var save_counting_school = [];
  var save_counting_with_70 = [];
  var save_finanziamento_school = [];
  var save_partecipanti_school = [];

  var total_sum_count_students = 0;
  var total_sum_count_students_with_70 = 0;
  var total_sum_finanziamento = 0;
  var total_sum_partecipanti = 0;

  var school_lines = [];

  for (var i = 0; i < students.length; i++) {
    
    if (students[i][0] !== "") {

      var range = sheet.getRange(i+1,5,1,5);
      range.setBackground("#FFFFFF")
      range = trimmingArray(range.getValues())[0].split(",");
      var summ = sum(range);
      var cell = sheet.getRange(i+1,11,1,1);
      var finanziamento = sheet.getRange(i+1,12,1,1);
      var certificate = sheet.getRange(i+1,13,1,1);
      
      cell.setValue(summ[0]);
      cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      cell.setHorizontalAlignment("center");

      finanziamento.setHorizontalAlignment("center");
      finanziamento.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      var fin_euro = 0;

      certificate.setHorizontalAlignment("center");
      certificate.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);

      if (summ[0] < 11){
        cell.setBackground("#D5202C"); 
        finanziamento.setBackground("#D5202C");
        certificate.setValue("no");
        certificate.setBackground("#D5202C");
      }else{
        cell.setBackground("#FFFFFF"); 
        fin_euro = 16.67 * summ[0] ;
        finanziamento.setBackground("#FFFFFF");
        certificate.setValue("yes");
        certificate.setBackground("#FFFFFF");
        count_students_with_70 = count_students_with_70 + 1;
      }

      finanziamento.setValue(fin_euro);
      total_finanziamento_school = total_finanziamento_school + fin_euro;
      count = count + 1;
      count_students_school = count_students_school + 1;

      if(summ[0] > 0){
        partecipanti_school = partecipanti_school + 1;
      }

      // in case of hours greater than 3, it is suggested to check it manually
      if(summ[1] === true){
        check = true;
        range = sheet.getRange(i+1,5,1,5);
        line_to_check = line_to_check + " " + (i+1).toString();
        range.setBackground("#FFFF00"); // yellow
        cell.setBackground("#FFFF00");
        finanziamento.setBackground("#FFFF00");
        certificate.setBackground("#FFFF00");
      }
    }

    var cell = sheet.getRange(i+1,2,1,1).getValue().trim();
    if (cell === "studenti" || i == students.length - 1){
      console.log("students at line: ", i+1);

      // remember to avoid the first elements 
      save_counting_school.push(count_students_school);
      save_counting_with_70.push(count_students_with_70);
      save_finanziamento_school.push(total_finanziamento_school);
      save_partecipanti_school.push(partecipanti_school);

      if (i != students.length - 1){
        school_lines.push(i+1);

        var cell = sheet.getRange(i+1,11,1,3);
        var values = [
          ["Ore totali", "Finanziamento", "certificato"]
        ];
        cell.setValues(values);
        cell.setHorizontalAlignment("center");
        cell.setFontWeight("bold");
        cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      }

      count_students_school = 0;
      count_students_with_70 = 0;
      total_finanziamento_school = 0;
      partecipanti_school = 0;
    }
  }

  // save some side statistics for each school
  for (var i = 0; i < school_lines.length; i++) {
    var cell = sheet.getRange(school_lines[i],14,8,1);
    var values = [
      ["Numero studenti iscritti"],
      [save_counting_school[i+1]],
      ["studenti partecipanti (>0h)"],
      [save_partecipanti_school[i+1]],
      ["studenti che hanno superato il 70%"],
      [save_counting_with_70[i+1]],
      ["finanziamento (in euro):"],
      [save_finanziamento_school[i+1]]
    ];
    cell.setValues(values);
    cell.setHorizontalAlignment("center");
    cell.setFontWeight("bold");
    
    total_sum_count_students = total_sum_count_students + save_counting_school[i+1];
    total_sum_partecipanti = total_sum_partecipanti + save_partecipanti_school[i+1];
    total_sum_count_students_with_70 = total_sum_count_students_with_70 + save_counting_with_70[i+1];
    total_sum_finanziamento = total_sum_finanziamento + save_finanziamento_school[i+1];
  }

  // save the total number of students in the sheet
  var cell = sheet.getRange(2,17,1,4);
  var values = [
    ["Numero di studenti in totale","studenti partecipanti (>0h)","studenti che hanno superato il 70%","finanziamento (in euro):"]
  ];
  cell.setValues(values);
  cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  cell.setHorizontalAlignment("center");
  cell.setFontWeight("bold");

  var cell = sheet.getRange(3,17,1,4);
  var values = [
    [total_sum_count_students,total_sum_partecipanti,total_sum_count_students_with_70,total_sum_finanziamento]
  ];
  cell.setValues(values);
  cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  cell.setHorizontalAlignment("center");
  cell.setFontWeight("bold");

  var cell = sheet.getRange(5,17,2,2);
  var values = [
    ["C'è qualche riga da controllare manualmente?", check ? "Sì" : "No"],
    ["Prova a Controllare le seguenti righe:", check ? line_to_check : ""]
  ];
  cell.setValues(values);
  cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  cell.setHorizontalAlignment("center");
  cell.setFontWeight("bold");
  
  return [total_sum_count_students,total_sum_partecipanti,total_sum_count_students_with_70,total_sum_finanziamento];
}

// sum all the value in the list, but if it is empty count it as zero
function sum(list) {
  var sum = 0;
  var check = false;
  for (var i = 0; i < list.length; i++) {
    if (list[i] !== "" && isNumeric(list[i])) {
      sum = sum + parseInt(list[i]);
      if (list[i] > 3){
        check = true;
      }
    }
  }
  return [sum, check];
}

function isNumeric(value) {
  return !isNaN(parseFloat(value)) && isFinite(value);
}

function trimmingArray(values){
  return values.map(function(row) {
    return row.toString().trim();
  });
}
