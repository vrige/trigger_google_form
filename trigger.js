function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // get the name of the subject 
  var name_of_the_course = e.values[1];

  // get the page of the subject
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_subject = ss.getSheetByName(name_of_the_course);
  
  console.log("name_of_the_course: ", name_of_the_course);


  // get the professor
  // if there are no professor mentioned in the form, then the form will be aborted
  var prof_form = getProfessor(e);
  if (prof_form["abort"] === true){
    // FIX-ME remove the row without professor and move it to the trouble page
    // abort the rest of the code
    console.log("form is aborted");
  }
  
  var professor = prof_form["professor"];

  if (professor === "Altro"){
    console.log("professor not in the list");
    var missing_professor_name = prof_form["missing_professor_name"];
    var missing_professor_surname = prof_form["missing_professor_surname"];
     // add missing professor
    addMissingProfessor(ss, name_of_the_course, professor, missing_professor_name, missing_professor_surname);
    professor = prof_form["missing_professor_name"] + " " + prof_form["missing_professor_surname"];
  }

  // update professor info if they are any in the form
  updateProfessorData(e, ss, name_of_the_course, professor);

  // get the list of students 
  var names = getStudentList(e);

  // get the date of the lesson
  var date_of_the_lesson = getDate(e);

  // get the duration of the lesson
  var duration = getDuration(e);

  // save the lines with the students attending the course
  var rows = getLines(names, sheet_subject);

  // save the institues involved in the form
  var institues = institutesInvolved(sheet_subject, rows);
  var institute_row = institues[0];
  var institute_unique = institues[1];

  // save the line numbers with the date of each institute
  var data_rows = dateRow(sheet_subject, institute_unique);

  // save the column for each student
  var number_columns = getColumnDate(sheet_subject, date_of_the_lesson, institute_unique, data_rows);

  // update the page of the students with their attendence
  updateStudents(sheet_subject, duration, rows, institute_unique, institute_row, number_columns);

  // update "Studenti Extra" page in case of extra students
  updateExtraStudents(e, ss, name_of_the_course, professor, date_of_the_lesson, duration)

}


// list of students attending the lesson
function getStudentList(e){
  var name_list_student_in_the_column;
  var searchString = "Studenti";
  // e.namedValues is a dict of each column with the argument inside
  // Object.keys(e.namedValues) extract just the key parts
  var fieldNames = Object.keys(e.namedValues);
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && fieldNames[i].trim() !== "Studenti extra"){
      if (e.namedValues[fieldNames[i]][0] !== ""){
        name_list_student_in_the_column = fieldNames[i];
        console.log("name_list_student_in_the_column: ", fieldNames[i], " i: ", i );
        break;
      }
    }
  }
  var names = e.namedValues[name_list_student_in_the_column][0].split(",");
  return names;
}


function getDate(e){
  // Now, we need to do the same with the column of the dates using as keyword "Data della lezione" and then 
  // check which column is not empty
  var date_of_the_lesson_column;
  var searchString = "Data della lezione";
  var fieldNames = Object.keys(e.namedValues);
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && e.namedValues[fieldNames[i]][0] !== ""){
      date_of_the_lesson_column = fieldNames[i];
      console.log("date_column: ", fieldNames[i], " i: ", i );
      break;
    }
  }
  var date_of_the_lesson = e.namedValues[date_of_the_lesson_column][0].trim();
  return date_of_the_lesson;
}

function getDuration(e){
  var duration_column = "Durata della lezione";
  var duration = e.namedValues[duration_column][0];
  return duration;
}

function getProfessor(e){
  // Retrieve the name of the professor
  var professor_column;
  var searchString = "Nome e Cognome docente";
  var fieldNames = Object.keys(e.namedValues);
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && e.namedValues[fieldNames[i]][0] !== ""){
      professor_column = fieldNames[i];
      console.log("professor_column: ", fieldNames[i], " i: ", i );
      break;
    }
  }
  var professor = e.namedValues[professor_column][0].trim();
  
  var missing_professor_name = "";
  var missing_professor_surname = "";

  if (professor === "Altro"){

    var searchString1 = "Nome docente mancante";
    var searchString2 = "Cognome docente mancante";
    var counter = 0;

    for (var i = 0; i < fieldNames.length; i++){

      if (fieldNames[i].includes(searchString1) && e.namedValues[fieldNames[i]][0] !== ""){
        professor_column = fieldNames[i];
        console.log("missing_professor_column_name: ", fieldNames[i], " i: ", i );
        missing_professor_name = e.namedValues[professor_column][0].trim();
        counter++;
      }

      if (fieldNames[i].includes(searchString2) && e.namedValues[fieldNames[i]][0] !== ""){
        professor_column = fieldNames[i];
        console.log("missing_professor_column_surname: ", fieldNames[i], " i: ", i );
        missing_professor_surname = e.namedValues[professor_column][0].trim();
        counter++;
      }
      if (counter == 2){
        break;
      }
    }

    // if no professor is mentioned in the form, abort the form
    if(missing_professor_name === "" || missing_professor_surname === ""){

      console.log("professor missing. Abort the form");
      return {abort: true, professor: professor, missing_professor_name : missing_professor_name, missing_professor_surname: missing_professor_surname};

    }else{
      return {abort: false, professor: professor, missing_professor_name : missing_professor_name, missing_professor_surname: missing_professor_surname};
    } 
  }
  return {abort: false, professor: professor, missing_professor_name : missing_professor_name, missing_professor_surname: missing_professor_surname};
}

// add the missing professor to the right list
function addMissingProfessor(ss, name_of_the_course, professor, name, surname){

  // assumption: the name of the page with the lessons is "Riassunto lezioni " + name_of_the_subject
  var subject_lessons_page = "Riassunto lezioni " + name_of_the_course;
  var sheet_subject_lesson = ss.getSheetByName(subject_lessons_page);

  var professors_column = trimmingArray(sheet_subject_lesson.getRange("A1:A" + sheet_subject_lesson.getLastRow()).getValues());
  var professor_row = professors_column.indexOf(professor) + 1;

  if ( professor_row !== -1){
    
    // add name and surname at the row with "Altro"
    if (professor === "Altro"){

      var cell = sheet_subject_lesson.getRange(professor_row, 1); 
      cell.setValue(name + " " + surname); 

      cell = sheet_subject_lesson.getRange(professor_row, 2); 
      cell.setValue(name); 
      
      cell = sheet_subject_lesson.getRange(professor_row, 3); 
      cell.setValue(surname); 

      // add "Altro" in the first column of the following row
      cell = sheet_subject_lesson.getRange(professor_row + 1, 1); 
      cell.setValue("Altro"); 
    }
  }
}

function updateProfessorData(e, ss, name_of_the_course, professor){

  // assumption: the name of the page with the lessons is "Riassunto lezioni " + name_of_the_subject
  var subject_lessons_page = "Riassunto lezioni " + name_of_the_course;
  var sheet_subject_lesson = ss.getSheetByName(subject_lessons_page);

  // assumption: fixed name for some columns: "Nome", "Cognome", "Codice Fiscale Docente", "Settore Lavorativo Docente",
  // "Docente Esterno" and "dottorando"
  // update the information about codice fiscale, settore lavorativo and docente esterno if they are not already 
  // present in the excel
  var professors_column = trimmingArray(sheet_subject_lesson.getRange("A1:A" + sheet_subject_lesson.getLastRow()).getValues());
  var professor_row = professors_column.indexOf(professor) + 1;

  if ( professor_row !== -1){
      
    var cell = sheet_subject_lesson.getRange(professor_row, 4); 
    console.log("professor_codice_fiscale: ", cell.getValue());
    if (cell.getValue() === ""){
      var codice_fiscale_column = "Codice Fiscale Docente";
      var codice_fiscale = e.namedValues[codice_fiscale_column][0];
      cell.setValue(codice_fiscale); 
    }

    cell = sheet_subject_lesson.getRange(professor_row, 5); 
    console.log("professor_settore_lavorativo: ", cell.getValue());
    if (cell.getValue() === ""){
      var settore_lavorativo_column = "Settore Scientifico Docente";
      var settore_lavorativo = e.namedValues[settore_lavorativo_column][0];
      cell.setValue(settore_lavorativo); 
    }

    cell = sheet_subject_lesson.getRange(professor_row, 6); 
    console.log("professor_docente_esterno: ", cell.getValue());
    if (cell.getValue() === ""){
      var docente_esterno_column = "Docente Esterno";
      var docente_esterno = e.namedValues[docente_esterno_column][0];
      cell.setValue(docente_esterno); 
    }

    cell = sheet_subject_lesson.getRange(professor_row, 7); 
    console.log("professor_dottorando: ", cell.getValue());
    if (cell.getValue() === ""){
      var dottorando_column = "Dottorando";
      var dottorando = e.namedValues[dottorando_column][0];
      cell.setValue(dottorando); 
    }
  }
}

// save the lines with the students attending the course
function getLines(names, sheet_subject){
  var row = [];
  for (var j = 0; j < names.length; j++){
    var searchString = names[j].trim();
    console.log("searchString: ", searchString);
    var subject_lines = sheet_subject.getRange("A1:A" + sheet_subject.getLastRow()).getValues();
    for (var i = 0; i < subject_lines.length; i++) {
      if (subject_lines[i][0] === searchString) {
        console.log("searchString at line: ", i + 1);
        row.push(i+1);
        break;
      }
    }
  }
  return row;
}

function institutesInvolved(sheet_subject, row){
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
  return [institute_row, institute_unique];
}

function dateRow(sheet_subject, institute_unique){
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
  return data_rows;
}


function getColumnDate(sheet_subject, date_of_the_lesson, institute_unique, data_rows){
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
  return number_columns;
}

function updateStudents(sheet_subject, duration, row, institute_unique, institute_row, number_columns){
  //let's update the value in the cells
  for (var j = 0; j < row.length; j++){
    var column_to_check = number_columns[institute_unique.indexOf(institute_row[j])];
    var cell = sheet_subject.getRange(row[j], column_to_check); 
    cell.setValue(duration);
  }
}

function updateExtraStudents(e, ss, name_of_the_course, professor, date_of_the_lesson, duration){
  // Now we need to add the extra students as last row in the dedicated page
  // assumption: page for the extra students "Studenti extra"
  // assumption: page for the name of professor must contain the keyword "Nome e Cognome docente"
  var studenti_extra_column = "Studenti extra";
  var studenti_extra = e.namedValues[studenti_extra_column][0];
  console.log("studenti_extra: ", studenti_extra);
  if (studenti_extra.trim().length !== 0){
    var sheet_extraStudents = ss.getSheetByName(studenti_extra_column);
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


