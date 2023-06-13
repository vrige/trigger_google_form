function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  /* assumptions on fixed column: 
     A - dynamic names for the form
     B - Name of the institue (in correspondece of a student)
     C - Name of the student
     D - Surname of the student
     from E to I possible lessons (assumption that there may be max 5 lessons)
  */
  // assumption: no duplicates in the first column (otherwise formRanger doesn't work)
  // assumption: when the names of the students are put in the riht column for the form,
  //             they must have a copy of the format of column A1 
  // assumption: the list of students must contain the keywords "Studenti"
  // assumption: the list of extra students is "Studenti extra"
  // assumption: the list of dates must contain the following keywork "Data della lezione"
  // assumption: there must one and only one column of students not empty (skip logic)
  // assumption: there must one and only one column of dates not empty (skip logic)

  // First of all, we need to filter the skip logic of the form
  // We need to find the column with the selected students using as keyword "Studenti" and then 
  // check which column is not empty
  var name_list_student_in_the_column;
  var searchString = "Studenti";
  var fieldNames = Object.keys(e.namedValues);
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && fieldNames[i].trim() !== "Studenti extra"){
      if (e.namedValues[fieldNames[i]][0] !== ""){
        name_list_student_in_the_column = fieldNames[i];
        console.log("name_list_student_in_the_column: ", fieldNames[i], " i: ", i );
      }
    }
  }
  var names = e.namedValues[name_list_student_in_the_column][0].split(",");

  // Now, we need to do the same with the column of the dates using as keyword "Data della lezione" and then 
  // check which column is not empty
  var date_of_the_lesson_column;
  var searchString = "Data della lezione";
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && e.namedValues[fieldNames[i]][0] !== ""){
      date_of_the_lesson_column = fieldNames[i];
      console.log("date_column: ", fieldNames[i], " i: ", i );
    }
  }
  var date_of_the_lesson = e.namedValues[date_of_the_lesson_column][0].trim();

  // Retrieve the name of the professor
  var professor_column;
  var searchString = "Nome e Cognome docente";
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && e.namedValues[fieldNames[i]][0] !== ""){
      professor_column = fieldNames[i];
      console.log("professor_column: ", fieldNames[i], " i: ", i );
    }
  }
  var professor = e.namedValues[professor_column][0].trim();

  var duration_column = "Durata";
  var duration = e.namedValues[duration_column][0];

  var name_of_the_course = e.values[1];

  console.log("name_of_the_course: ", name_of_the_course);
  console.log("duration: ", duration);
  console.log("date_of_the_lesson: ", date_of_the_lesson);
  console.log("professor: ", professor);
  console.log(names);

  
  var sheet_name = name_of_the_course;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // assumption: same name between pages and subjects
  var sheet_subject = ss.getSheetByName(sheet_name);
  

  // assumption: always look for the matching string in the first column
  var columnToSearch = 1;
  // save the lines with the students attending the course
  var row = [];
  for (var j = 0; j < names.length; j++){
    var searchString = names[j].trim();
      console.log("searchString: ", searchString);
      var subject_lines = sheet_subject.getRange("A1:A" + sheet_subject.getLastRow()).getValues();
      for (var i = 0; i < subject_lines.length; i++) {
        if (subject_lines[i][0] === searchString) {
          console.log("searchString at line: ", i + 1);
          row.push(i+1);
        }
      }
  }

  // assumption: students from the same institute attending the same course will attend the lesson at the same time
  // assumption: students from the same school attending the same course are on the same table in the same page
  // get the name of the institute for each student 
  var institute_row = [];
  var columnToSearch_number = 2;
  for (var j = 0; j < row.length; j++){
    var cell = sheet_subject.getRange(row[j], columnToSearch_number).getValue(); 
    institute_row.push(cell);
  }
  institute_row = trimmingArray(institute_row);
  // get the names of the institutes
  var institute_unique = [];
  for (var i = 0; i < institute_row.length; i++) {
    if (institute_unique.indexOf(institute_row[i]) == -1) {
      institute_unique.push(institute_row[i]);
      console.log("istitute: ", institute_row[i]);
    }
  }
  
  // assumption: for each school table there must be a column starting with "studenti"
  // assumption: the first student should be not further than 5 cells from the cell "studenti"
  // assumption: between the first student and the cell "studenti", the cells must be empty
  // assumption: the row with the dates is always - 2 wrt the row with cell "studenti"
  var columnToSearch = "B1:B";
  var searchString = "studenti";
  var count = 1;
  var data_lines = trimmingArray(sheet_subject.getRange(columnToSearch + sheet_subject.getLastRow()).getValues());
  var data_rows = [];

  // the idea is to search in the whole column B1 (the one with students' schools) for the "studenti" cells
  // when one is found, the algorithm check the following lines to see the school of the first student
  // if there is an empty space the cell below is checked instead (mechanism iterated up to 5 times)
  // In the end, we want to retrieve the row numbers of the lines with the dates for each school (in the list)
  for (var i = 0; i < data_lines.length; i++){  
    if (data_lines[i] === searchString) {
      count = 1;
      // check if there is a white space between the cell "studenti" and the first student
      while (count < 5) {
        var cellValue = data_lines[i + count];
        if (cellValue === "") {
          console.log("The cell is empty.");
        } else {
          console.log("The cell is not empty.");
          
          var val = institute_unique.indexOf(cellValue);
          if(val != -1){
            data_rows.push(i - 2);
            console.log("Institute: ", cellValue, " at line: ", i - 2);
            // we need also to put it in the correct order
            institute_unique.splice(val,1);
            institute_unique.push(cellValue);
            break;
          }
          
          break;
        }
        count++; 
      }
    }
    if (data_rows.length === institute_unique.length){
      break;
    }
  }
  
  // assumption: the calendar data is always in the format gg/mm/aa in the google sheet
  // assumption: the column in which the data may be are always E, F, G, H, I
  // now we will check the numbers of columns for the lesson that we are looking for (date_of_the_lesson)
  // the lesson can be in a different column for each school
  // we are going to use the previous row numbers with all the interesting dates for the selected students
  var searchString = createDateFromFormat(date_of_the_lesson);
  var date1 = new Date(searchString);
  var number_columns = [];
  for (var j = 0; j < data_rows.length; j++){ 
    var data_lines = trimmingArray(sheet_subject.getRange(data_rows[j] + 1, 5, 1, 6).getValues());
    data_lines = data_lines[0].split(",");
    for (var i = 0; i < data_lines.length; i++) {
      var date2 = new Date(data_lines[i]);
      if (date1.getTime() === date2.getTime()) {
        console.log("searchString at column: ", i + 5, " for the institute ",institute_unique[j]);
        number_columns.push(i + 5); 
        break;
      }
    }
  }
  console.log("number_columns: ", number_columns);

  //let's update the value in the cells
  for (var j = 0; j < row.length; j++){
    var column_to_check = number_columns[institute_unique.indexOf(institute_row[j])];
    var cell = sheet_subject.getRange(row[j], column_to_check); 
    cell.setValue(duration);
  }
  
  //Now we need to add the extra students as last row in the dedicated page
  // assumption: page for the extra students "Studenti extra"
  // assumption: page for the name of professor must contain the keyword "Nome e Cognome docente"
  var studenti_extra_column = "Studenti extra";
  var studenti_extra = e.namedValues[studenti_extra_column][0];
  console.log("studenti_extra: ", studenti_extra);
  if (studenti_extra.trim().length !== 0){
    var sheet_extraStudents = ss.getSheetByName("Studenti extra");
    var lastRow = sheet_extraStudents.getLastRow();
    var newRow = lastRow + 1;
    var cell = sheet_extraStudents.getRange(newRow, 1); 
    cell.setValue(name_of_the_course);
    cell = sheet_extraStudents.getRange(newRow, 2); 
    cell.setValue(professor);
    cell = sheet_extraStudents.getRange(newRow, 3); 
    cell.setValue(date_of_the_lesson);
    cell = sheet_extraStudents.getRange(newRow, 4); 
    cell.setValue(duration);
    cell = sheet_extraStudents.getRange(newRow, 5); 
    cell.setValue(studenti_extra);
  }
  
  // assumption: the name of the page with the lessons is "Riassunto lezioni " + name_of_the_subject
  var subject_lessons_page = "Riassunto lezioni " + name_of_the_course;
  var sheet_subject_lesson = ss.getSheetByName(subject_lessons_page);

  // assumption: fixed name for some columns: "Codice Fiscale Docente", "Settore Lavorativo Docente", "Docente Esterno"
  //             and "dottorando"
  // update the information about codice fiscale, settore lavorativo and docente esterno if they are not already 
  // present in the excel
  var professors_column = trimmingArray(sheet_subject_lesson.getRange("A1:A" + sheet_subject_lesson.getLastRow()).getValues());
  var professor_row = professors_column.indexOf(professor) + 1;
  console.log("professor_row: ", professor_row);
  if ( professor_row !== -1){
    var cell = sheet_subject_lesson.getRange(professor_row, 4); 
    console.log("professor_codice_fiscale: ", cell.getValue());
    if (cell.getValue() === ""){
      var codice_fiscale_column = "Codice Fiscale Docente";
      var codice_fiscale = e.namedValues[codice_fiscale_column][0];
      cell.setValue(codice_fiscale); 
    }
    var cell = sheet_subject_lesson.getRange(professor_row, 5); 
    console.log("professor_settore_lavorativo: ", cell.getValue());
    if (cell.getValue() === ""){
      var settore_lavorativo_column = "Settore Lavorativo Docente";
      var settore_lavorativo = e.namedValues[settore_lavorativo_column][0];
      cell.setValue(settore_lavorativo); 
    }
    var cell = sheet_subject_lesson.getRange(professor_row, 6); 
    console.log("professor_docente_esterno: ", cell.getValue());
    if (cell.getValue() === ""){
      var docente_esterno_column = "Docente Esterno";
      var docente_esterno = e.namedValues[docente_esterno_column][0];
      cell.setValue(docente_esterno); 
    }
    var cell = sheet_subject_lesson.getRange(professor_row, 7); 
    console.log("professor_dottorando: ", cell.getValue());
    if (cell.getValue() === ""){
      var dottorando_column = "Dottorando";
      var dottorando = e.namedValues[dottorando_column][0];
      cell.setValue(dottorando); 
    }
  }

}

function createDateFromFormat(dateString) {
  const [day, month, year] = dateString.split('/');
  const date = new Date(`${month}/${day}/${year}`);
  return date.toString().trim();
}

function trimmingArray(values){
  return values.map(function(row) {
    return row.toString().trim();
  });
}

