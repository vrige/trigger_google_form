// this function is set to be triggered when a form is received
function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // get the tab "Frequenze"
  var sheet_frequency_obj = getSheetByItsName(ss, "Frequenze");
  var sheet_frequency = sheet_frequency_obj["sheet"];

  // get the name of the subject 
  var name_of_the_course = e.values[1];

  // get the page of the subject
  var sheet_subject_obj = getSheetByItsName(ss, name_of_the_course);
  var sheet_subject = sheet_subject_obj["sheet"];

  // get the page "Riassunto lezioni " of a specific subject
  var subject_lessons_page = "Riassunto lezioni " + name_of_the_course;
  var sheet_subject_lesson_obj = getSheetByItsName(ss, subject_lessons_page);
  var sheet_subject_lesson = sheet_subject_lesson_obj["sheet"];

  // get the page for "Studenti extra"
  var sheet_extraStudents_obj = getSheetByItsName(ss, "Studenti extra");
  var sheet_extraStudents = sheet_extraStudents_obj["sheet"];

  // copy new row in the "Frequenze" tab
  // update it even in case of error
  var row_sheet = updateMainTab(e, sheet, sheet_frequency);
  
  // first layer of error handling: checking that the useful following pages exist:
  // "Studenti extra", name_of_the_course, "Riassunto lezioni " + name_of_the_course, "Frequenze"
  if( sheet_frequency_obj["abort"] === false && sheet_subject_obj["abort"] === false &&
      sheet_subject_lesson_obj["abort"] === false && sheet_extraStudents_obj["abort"] === false){
    
    console.log("The 4 useful pages exist")   

    // get the professor name and surname
    var professor_obj = getProfessor(e);  
    var professor = professor_obj["professor"];

    // get the date of the lesson
    var date_of_the_lesson_obj = getDataFromColumn(e, "Data della lezione");
    var date_of_the_lesson = date_of_the_lesson_obj["values"];

    // get the duration of the lesson
    var duration_obj = getDataFromColumn(e, "Durata della lezione");
    var duration = duration_obj["values"];

    // get data from the column "Nessuno studente è presente"
    var no_students_obj = getDataFromColumn(e, "Nessuno studente è presente");
    var no_students = no_students_obj["values"];

    // if the professor has selected the option "Sì" on "Nessuno studente è presente"
    // then it is possible to avoid the check on the student column
    console.log("nessuno studente è presente: ", no_students)

    // get the list of students 
    var names_obj = {values: [], abort: false, error: "", name: "Nessuno studente è presente",check: true};
    var names = names_obj["values"];
    if (no_students === "No"){
      names_obj = getDataFromColumn(e, "Studenti", "Studenti extra");
      names = names_obj["values"];
    }

    // second layer of error handling: checking form columns
    if(professor_obj["abort"] === false && names_obj["abort"] === false &&
    date_of_the_lesson_obj["abort"] === false && duration_obj["abort"] === false){

      console.log("All the necessary form columns are there")

      // save the lines with the students attending the course
      var rows_obj = getLines(names, sheet_subject);
      var rows = rows_obj["values"];

      // save the institues involved in the form
      var institues = institutesInvolved(sheet_subject, rows);
      var institute_row = institues[0];
      var institute_unique = institues[1];

      // save the line numbers with the date of each institute
      var data_rows_obj = dateRow(sheet_subject, institute_unique);
      var data_rows = data_rows_obj["values"];

      // save the column for each student
      var number_columns_obj = getColumnDate(sheet_subject, date_of_the_lesson, institute_unique, data_rows);
      var number_columns = number_columns_obj["values"];

      // third layer error handling: matching
      if (rows_obj["abort"] === false && data_rows_obj["abort"] === false && number_columns_obj["abort"] === false){
        
        console.log("form is valid");

        var extra_students_obj = getExtraStudenti(e);
        var extra_students = extra_students_obj["values"];

        var extra_prof_obj = getExtraProfessorData(e);

        // forth layers: security check on the entry data
        if(extra_students_obj["abort"] === false && extra_prof_obj["abort"] === false && professor_obj["sec_abort"] === false){

          console.log("Secuirty check passed");

          // update professors' list in case the are no errors
          // if there are no professor mentioned in the form, then the form will be aborted
          if (professor === "Altro"){
            console.log("professor not in the list");
            var missing_professor_name = professor_obj["missing_professor_name"];
            var missing_professor_surname = professor_obj["missing_professor_surname"];
            
            addMissingProfessor(sheet_subject_lesson, professor, missing_professor_name, missing_professor_surname);
            professor = missing_professor_name + " " + missing_professor_surname;
          }

          // update professor info if they are any in the form
          updateProfessorData(sheet_subject_lesson, professor, extra_prof_obj["codice_fiscale"], 
          extra_prof_obj["settore_lavorativo"], extra_prof_obj["docente_esterno"], extra_prof_obj["dottorando"]);

          // update the page of the students with their attendence
          updateStudents(sheet_subject, duration, rows, institute_unique, institute_row, number_columns);

          // update "Studenti Extra" page in case of extra students
          updateExtraStudents(extra_students, sheet_extraStudents, name_of_the_course, professor,
                               date_of_the_lesson, duration);

          if(extra_students_obj["check"] === true || professor_obj["check"] === true || names_obj["check"] === true || extra_prof_obj["check"] === true){
            var nature_check = "";
            if (extra_students_obj["check"] === true ){
              nature_check = nature_check + " Please check any extra student.\n";
            }
            if (professor_obj["check"] === true){
              nature_check = nature_check + " Please check any extra professor.\n";
            }
            if (extra_prof_obj["check"] === true ){
              nature_check = nature_check + " Please check any new information about the professor.\n";
            }
            if (names_obj["check"] === true ){
              nature_check = nature_check + " Please check why no student were selected from the list.\n" + "The student list in the form may be the wrong one.\n";
            }
            updateState(sheet_frequency_obj, row_sheet, nature_check, "", "#FBEF46", "Check");  // color yellow
            
          }else{
            // update that everything is fine
            updateState(sheet_frequency_obj, row_sheet, "", "", "#1BA937", "Ok");  // color green
          }
          

        }else{
          
          console.log("Secuirty check failed");

          // elaborate message error for the security layer: entry data checking
          var error = "";
          var solution = "";

          if(extra_students_obj["abort"] === true){
            error = error + extra_students_obj["error"] + "\n";
          }
          if(extra_prof_obj["abort"] === true){
            error = error + extra_prof_obj["error"] + "\n";
          }
          if(professor_obj["sec_abort"] === true){
            error = error + professor_obj["error"] + "\n";
          }

          solution = "Controllare manualmente che simboli sono stati inseriti. In questo caso è bene notificare un" + 
          " informatico \n per capire se c'è stato un tentativo di hacking oppure se è stato usato uno o più simboli che" + " solitamente non vengono usati. \n";

          updateState(sheet_frequency_obj, row_sheet, error, solution);
        }

      }else{

        // elaborate message error for the third layer: matching
        var error = "";
        var solution = "";

        if(rows_obj["abort"] === true){
          error = error + rows_obj["error"] + "\n";
          solution =  rows_obj["solution"] + "\n";
        }
        if(data_rows_obj["abort"] === true){
          error = error + data_rows_obj["error"] + "\n";
          solution = data_rows_obj["solution"] + "\n";
        }
        if(number_columns_obj["abort"] === true){
          error = error + number_columns_obj["error"] + "\n";
          solution = number_columns_obj["solution"] + "\n";
        }
      
        console.log(error)
        console.log(solution)

        updateState(sheet_frequency_obj, row_sheet, error, solution);
      }

    }else{
      console.log("There was some problems with the form columns");

      // elaborate message error for the second layer: missing form columns
      var error = "";
      var solution = "";

      if(professor_obj["abort"] === true){
        error = error + professor_obj["error"] + "\n";
        solution = solution + "Controllare che vi sia il nome e cognome del professore nel form. Potrebbe essere" + 
                  " un errore del professore che ha compilato il form senza il nome e/o il cognome.\n";
      }
      if(names_obj["abort"] === true){
        error = error + names_obj["error"] + "\n";
        solution = solution + "Controllare che esista una colonna del form che contenga la parola \'" + 
        names_obj["name"] + "\' e che il nome sia corretto.\n";
        solution = solution + "Può anche essere che il professore non abbia selezionato nessun studente dalla lista e abbia segnato l'aula come non vuota.\n"+ "Controllare il numero di studenti e la colonna \'Nessuno studente è presente\' per confermare tale ipotesi.\n";
      }
      if(date_of_the_lesson_obj["abort"] === true){
        error = error + date_of_the_lesson_obj["error"] + "\n";
        solution = solution + "Controllare che esista una colonna del form che contenga la parola \'" + 
        date_of_the_lesson_obj["name"] + "\' e che il nome sia corretto.\n";
      }
      if(duration_obj["abort"] === true){
        error = error + duration_obj["error"] + "\n";
        solution = solution + "Controllare che esista una colonna del form che contenga la parola \'" + duration_obj["name"] + "\'  e che il nome sia corretto.\n";
      }
      solution = solution + "Per controllare che il nome sia correto andare sulla corrispondente form e controllare \n" + 
                "che vi sia una domanda con la/e parola/e chiave/i precedentemente specificata/e.";
    
      console.log(error)
      console.log(solution)

      updateState(sheet_frequency_obj, row_sheet, error, solution);
    }
    
  }else{
    console.log("At least one the four useful pages doesn't exist. The form is aborted");

    // elaborate message error for the first layer on missing sheet(s)
    var error = "";
    var solution = "";
    if(sheet_frequency_obj["abort"] === true){
      error = error + sheet_frequency_obj["error"] + "\n";
      solution = solution + "Controllare che esista il foglio \'" + sheet_subject_obj["name"]  + 
                 "\' e che il nome sia corretto.\n";
    }
    if(sheet_subject_obj["abort"] === true){
      error = error + sheet_subject_obj["error"] + "\n";
      solution = solution + "Controllare che esista il foglio \'" + sheet_subject_obj["name"] + 
                "\' e che il nome sia corretto.\n";
    }
    if(sheet_subject_lesson_obj["abort"] === true){
      error = error + sheet_subject_lesson_obj["error"] + "\n";
      solution = solution + "Controllare che esista il foglio \'" + sheet_subject_lesson_obj["name"] + 
                "\' e che il nome sia corretto.\n";
    }
    if(sheet_extraStudents_obj["abort"] === true){
      error = error + sheet_extraStudents_obj["error"] + "\n";
      solution = solution + "Controllare che esista il foglio \'" + sheet_extraStudents_obj["name"] + 
                "\' e che il nome sia corretto.\n";
    }
    solution = solution + "Per controllare che il nome sia correto andare sul foglio \'Lista Materie\'" +
              " e fare un copia e incolla dalla lista."
  
    console.log(error)
    console.log(solution)

    updateState(sheet_frequency_obj, row_sheet, error, solution);
  }
}


// assumption: the state is in column Q (number 17), while the error in column R (number 18)
//             and the solutions are in column S (number 19)
// update status, error and solution columns in the "Frequenze" sheet
function updateState(sheet_frequency_obj, row_sheet, error, solution, color="#D5202C", text="Form aborted"){
  
  if(sheet_frequency_obj["abort"] !== true){
    sheet_frequency = sheet_frequency_obj["sheet"];
    var cell = sheet_frequency.getRange(row_sheet, 17, 1, 1);
    cell.setValue(text).setHorizontalAlignment("center").setVerticalAlignment('middle');
    cell.setBackground(color); 
    cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
    
    var cell = sheet_frequency.getRange(row_sheet, 18, 1, 1);
    cell.setValue(error);
    cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  
    var cell = sheet_frequency.getRange(row_sheet, 19, 1, 1);
    cell.setValue(solution);
    cell.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  }else{
    console.log("Cannot update the status, error and solution becasue the sheet \'Frequenze\' doesn't exist")
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

// copy new row in the "Frequenze" tab
function updateMainTab(e, sheet, targetSheet){
  var targetRange = 1;
  var new_row = e.range.getRow();

  targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1, 1, sheet.getLastColumn());
  sheet.getRange(new_row, 1, 1, sheet.getLastColumn()).copyTo(targetRange);

  return targetSheet.getLastRow();
}

// get the column with the word specified in "searchString". If you need to discriminate two or more columns,
// it is possible to specified some words to avoid in in the parameter "avoid"
function getDataFromColumn(e, searchString, avoid="-"){
  var name = 0;
  var abort = false;
  var error = "";
  var check = false;
  var values = "";
  // e.namedValues is a dict of each column with the argument inside
  // Object.keys(e.namedValues) extract just the key parts
  var fieldNames = Object.keys(e.namedValues);

  // look for the column with the searchStrign inside, but not avoid
  for (var i = 0; i < fieldNames.length; i++){
    if (fieldNames[i].includes(searchString) && fieldNames[i].trim() !== avoid){
      if (e.namedValues[fieldNames[i]][0] !== ""){
        name = fieldNames[i];
        //console.log("name: ", fieldNames[i], " i: ", i );
        break;
      }
    }
  }

  // check if the column was found
  if (name === 0){
    abort = true;
    error = "Cannot find the column with the words \'"+ searchString + "\'";
    console.log(error);
  }else{
    values = e.namedValues[name][0]; 
    console.log("Found column with the words \'"+ searchString + "\' in: ", values);
    if(searchString === "Studenti"){  // in case of the students, split the string in a list of students
      values = values.split(",");
      //console.log("students: "+  values);
    }else{
      values = values.trim();
    }
  }
 
  return {abort : abort, error: error, check: check, name: searchString, values: values};
}

// Retrieve the name and surname of the professor
function getProfessor(e){
  
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
  var abort = false;
  var error = "";
  var check = false;
  var sec_abort = false;

  // in case no professor is selected from the list, then we need to retrieve his/her 
  // name and surname from the optional fields
  if (professor === "Altro"){

    var searchString1 = "Nome docente mancante";
    var searchString2 = "Cognome docente mancante";
    var counter = 0;
    check = true;
    error = "check the new teacher added to the list. \n";

    // check name and surname of the new professor
    for (var i = 0; i < fieldNames.length; i++){

      // check name 
      if (fieldNames[i].includes(searchString1) && e.namedValues[fieldNames[i]][0] !== ""){
        professor_column = fieldNames[i];
        console.log("missing_professor_column_name: ", fieldNames[i], " i: ", i );
        missing_professor_name = e.namedValues[professor_column][0].trim();
        counter++;
      }

      // check surname
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

    // check if the data is not secure
    var name_obj = validateEntry(missing_professor_name);
    var surname_obj = validateEntry(missing_professor_surname);

    // if no professor is mentioned in the form, abort the form
    if(missing_professor_name === "" || missing_professor_surname === ""){

      console.log("professor missing. Abort the form");
      missing_professor_name = missing_professor_name;
      missing_professor_surname = missing_professor_surname;
      abort = true;
      check = false;
      error = error + "The professor has selected \'Altro\' on the form, but she/he didn't fill the field"+ 
              "about the name and/or surname.\n";
    }

    if(name_obj["abort"] === true){
      sec_abort = true;
      error = error + name_obj["error"] + " inside " + searchString1;
    }
    if(surname_obj["abort"] === true){
      sec_abort = true;
      error = error + surname_obj["error"] + " inside " + searchString2;
    }
  }
  return {abort: abort, error: error, check: check, sec_abort: sec_abort, professor: professor, missing_professor_name : missing_professor_name, missing_professor_surname: missing_professor_surname};
}

// add the missing professor to the right list
function addMissingProfessor(sheet_subject_lesson, professor, name, surname){

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

// check entry data on the professor data
function getExtraProfessorData(e){

  var abort = false;
  var error = "";
  var check = false;

  var codice_fiscale_column = "Codice Fiscale Docente";
  var codice_fiscale = e.namedValues[codice_fiscale_column][0];
  var codice_fiscale_obj = validateEntry(codice_fiscale);

  var settore_lavorativo_column = "Settore Scientifico Docente";
  var settore_lavorativo = e.namedValues[settore_lavorativo_column][0];
  var settore_lavorativo_obj = validateEntry(settore_lavorativo);

  var docente_esterno_column = "Docente Esterno";
  var docente_esterno = e.namedValues[docente_esterno_column][0];
  var docente_esterno_obj = validateEntry(docente_esterno);

  var dottorando_column = "Dottorando";
  var dottorando = e.namedValues[dottorando_column][0];
  var dottorando_obj = validateEntry(dottorando);

  if (codice_fiscale_obj["abort"]){
    error = error + codice_fiscale_obj["error"] + " in " + codice_fiscale_column + "\n";
    abort = true;
  }
  if (settore_lavorativo_obj["abort"]){
    error = error + settore_lavorativo_obj["error"] + " in " + settore_lavorativo_column + "\n";
    abort = true;
  }
  if (docente_esterno_obj["abort"]){
    error = error + docente_esterno_obj["error"] + " in " + docente_esterno_column + "\n";
    abort = true;
  }
  if (dottorando_obj["abort"]){
    error = error + dottorando_obj["error"] + " in " + dottorando_column + "\n";
    abort = true;
  }
  if (codice_fiscale !== "" || settore_lavorativo !== "" || docente_esterno !== "" || dottorando !== ""){
    check = true;
  }

  return {abort: abort, error: error, check: check, codice_fiscale: codice_fiscale, settore_lavorativo: settore_lavorativo, docente_esterno: docente_esterno, dottorando: dottorando};
}


// assumption: fixed name for some columns: "Nome", "Cognome", "Codice Fiscale Docente", "Settore Lavorativo Docente",
// "Docente Esterno" and "dottorando"
// update the information about codice fiscale, settore lavorativo and docente esterno if they are not already 
// present in the excel
function updateProfessorData(sheet_subject_lesson, professor, codice_fiscale, settore_lavorativo, docente_esterno, dottorando){

  var professors_column = trimmingArray(sheet_subject_lesson.getRange("A1:A" + sheet_subject_lesson.getLastRow()).getValues());
  var professor_row = professors_column.indexOf(professor) + 1;
  console.log("professor_row: ", professor_row);
  console.log("professor: ", professor);

  if ( professor_row !== -1){
      
    var cell = sheet_subject_lesson.getRange(professor_row, 4); 
    if (cell.getValue() === ""){
      cell.setValue(codice_fiscale); 
    }

    cell = sheet_subject_lesson.getRange(professor_row, 5); 
    if (cell.getValue() === ""){
      cell.setValue(settore_lavorativo); 
    }

    cell = sheet_subject_lesson.getRange(professor_row, 6); 
    if (cell.getValue() === ""){
      cell.setValue(docente_esterno); 
    }

    cell = sheet_subject_lesson.getRange(professor_row, 7); 
    if (cell.getValue() === ""){
      cell.setValue(dottorando); 
    }
  }
}

// save the lines with the students attending the course
function getLines(names, sheet_subject){
  var row = [];
  var missing = [];
  var subject_lines = sheet_subject.getRange("A1:A" + sheet_subject.getLastRow()).getValues();
  var flag;
  var abort = false;
  var error = "";
  var solution = "";

  // iterate over the student names from the form 
  for (var j = 0; j < names.length; j++){
    var searchString = names[j].trim();
    console.log("searchString: ", searchString);
    flag = false;

    // iterate over all the student names attending that specific course
    for (var i = 0; i < subject_lines.length; i++) {

      // collect the row number of the student
      if (subject_lines[i][0] === searchString) {
        console.log("searchString at line: ", i + 1);
        row.push(i+1);
        flag = true;
        break;
      }
      if( i === (subject_lines.length - 1) && flag === false){
        missing.push(searchString);
        break;
      }
    }
  }

  // check that all students were found 
  if (missing.length > 0){
    abort = true;
    error = "Some students are missing. Here there is a list:\n";
    for (var i = 0; i < missing.length; i++){
      error = error + "- " + missing[i] + "\n";
    }
    solution = "Potrebbero mancare alcuni studenti nella lista oppure potrebbero essere stati trascritti male.\n" + 
               "Provare a controllare la lista degli studenti e in particolare questi nomi:\n";
    for (var i = 0; i < missing.length; i++){
      solution = solution + "- " + missing[i] + "\n";
    } 
    console.log(error);
  }

  return {abort: abort, error: error, values: row, solution: solution};
}

// assumption: students from the same institute attending the same course will attend the lesson at the same time
// assumption: students from the same school attending the same course are on the same table in the same page
// get the name of the institute for each student 
function institutesInvolved(sheet_subject, row){
  var institute_row = [];
  var columnToSearch_number = 2;

  // get the school for each student
  for (var j = 0; j < row.length; j++){
    var cell = sheet_subject.getRange(row[j], columnToSearch_number).getValue(); 
    institute_row.push(cell);
  }
  institute_row = trimmingArray(institute_row);

  // get the names of the institutes
  var institute_unique = [];
  for (var i = 0; i < institute_row.length; i++) {
    if (institute_unique.indexOf(institute_row[i]) === -1) {
      institute_unique.push(institute_row[i]);
      console.log("istitute: ", institute_row[i]);
    }
  }
  return [institute_row, institute_unique];
}

// assumption: for each school table there must be a column starting with "studenti"
// assumption: the first student should be not further than 5 cells from the cell "studenti"
// assumption: between the first student and the cell "studenti", the cells must be empty
// assumption: the row with the dates is always - 2 wrt the row with cell "studenti"
// assumption: the school names on the second column must be all copies if they are refering to the same school
function dateRow(sheet_subject, institute_unique){

  var columnToSearch = "B1:B";
  var searchString = "studenti";
  var count = 1;
  var data_lines = trimmingArray(sheet_subject.getRange(columnToSearch + sheet_subject.getLastRow()).getValues());
  var data_rows = [];
  var abort = false;
  var error = "";
  var solution = "";

  // the idea is to search in the whole column B1 (the one with students' schools) for the "studenti" cells
  // when one is found, the algorithm check the following lines to see the school of the first student
  // if there is an empty space the cell below is checked instead (mechanism iterated up to 7 times)
  // In the end, we want to retrieve the row numbers of the lines with the dates for each school (in the list)
  for (var i = 0; i < data_lines.length; i++){  
    if (data_lines[i] === searchString) {
      count = 1;
      // check if there is a white space between the cell "studenti" and the first student
      while (count < 7) {

        var cellValue = data_lines[i + count].trim();
        var check = institute_unique.indexOf(cellValue);

        if (cellValue === "" || check == -1 ) {
          console.log("The cell is empty or no school is written inside: ", cellValue);

        } else {

          console.log("The cell is not empty.");       
          data_rows.push(i - 1);
          console.log("Institute: ", cellValue, " at line: ", i - 1);

          // we need also to put it in the correct order
          institute_unique.splice(check,1);
          institute_unique.push(cellValue);
          console.log("Institutes order: ", institute_unique);
          break;
        }
        count++; 
      }
    }
    if (data_rows.length === institute_unique.length){
      break;
    }
  }

  // check that all the institutes were found
  if (data_rows.length !== institute_unique.length){
     
    if(data_rows.length > institute_unique.length){
      error = "More institutes than the needed ones";
      abort = true;
      solution = "Provare a controllare che il nome delle scuole non sia stato scritto in maniera" + 
                 "diversa per studenti che frequentano la stessa scuola";
      console.log(error);
    } else{
      error = "Some institutes were not found";
      abort = true;
      solution = "Controllare che la keyword \'studenti\' sia presente all'altezza delle keywords \'nome\'"+
                 " e \'cognome\' per ogni scuola.\n" + "Provare a controllare le seguenti scuole: \n";
      for (var i = 0; i < institute_unique.length; i++){
        solution = solution + "- " + institute_unique[i] + "\n";
      }
      console.log(error);
    }
  }

  return {abort: abort, error: error, solution: solution, values: data_rows};
}

// assumption: the calendar data is always in the format gg/mm/aa in the google sheet
// assumption: the column in which the data may be are always E, F, G, H, I
// now we will check the numbers of columns for the lesson that we are looking for (date_of_the_lesson)
// the lesson can be in a different column for each school
// we are going to use the previous row numbers with all the interesting dates for the selected students
function getColumnDate(sheet_subject, date_of_the_lesson, institute_unique, data_rows){

  var searchString = createDateFromFormat(date_of_the_lesson);
  var date1 = new Date(searchString);
  var number_columns = [];
  var index_missing = [];
  var count;
  var abort = false;
  var error = "";
  var solution = "";

  for (var j = 0; j < data_rows.length; j++){ 
    var data_lines = trimmingArray(sheet_subject.getRange(data_rows[j], 5, 1, 6).getValues());
    data_lines = data_lines[0].split(",");
    count = 0;

    for (var i = 0; i < data_lines.length; i++) {
      var date2 = new Date(data_lines[i]);
      if (date1.getTime() === date2.getTime()) {
        console.log("searchString at column: ", i + 5, " for the institute ",institute_unique[j]);
        number_columns.push(i + 5); 
        count = 1;
        break;
      }
    }

    // if no data was found, save as missing
    if(count == 0){
      index_missing.push(j);
    }
  }

  if(index_missing.length > 0){
    abort = true;
    error = "Cannot find the right column date for all the institutes.";
    solution = "Potrebbe essere che le date per le scuole coinvolte non siano nel formato giusto o che " +
                "la data sia sbaglaita.\n" + "Controllare che il tipo di cella delle date " +
                "delle seguenti scuole sia settato su data:\n";
    for (var i = 0; i < index_missing.length; i++){
      solution = solution + "- " + institute_unique[index_missing[i]] + "\n";
    } 
    solution = solution + "Potrebbe essere anche un errore dal parte del docente " + 
              "che ha messo la data  sbagliata";          
  }
  console.log("number_columns: ", number_columns);

  return {abort: abort, error: error, solution: solution, values: number_columns};
}

function updateStudents(sheet_subject, duration, row, institute_unique, institute_row, number_columns){
  //let's update the value in the cells
  for (var j = 0; j < row.length; j++){
    var column_to_check = number_columns[institute_unique.indexOf(institute_row[j])];
    var cell = sheet_subject.getRange(row[j], column_to_check); 
    cell.setValue(duration);
  }
}

// assumption: page for the extra students "Studenti extra"
// get any extra student and validate the entry data
function getExtraStudenti(e){

  var studenti_extra_column = "Studenti extra";
  var studenti_extra = e.namedValues[studenti_extra_column][0];

  var studenti_extra_obj = validateEntry(studenti_extra);
  var error = studenti_extra_obj["error"] + " in " + studenti_extra_column + "\n";
  var check = false;

  // manually check in case of new students 
  if (studenti_extra.trim().length !== 0){
    check = true;
  }

  return {abort: studenti_extra_obj["abort"], error: error, check: check, values: studenti_extra};
}

// Now we need to add the extra students as last row in the dedicated page
// assumption: page for the name of professor must contain the keyword "Nome e Cognome docente"
function updateExtraStudents(studenti_extra, sheet_extraStudents, name_of_the_course, professor, date_of_the_lesson, duration){

  if (studenti_extra.trim().length !== 0){

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

function validateEntry(data){
  var abort = false;
  var error = "";

  // Dangerous symbols you want to check for
  var dangerousSymbols = ['<', '>', '&', '"', "'", '=','!','(',')','[',']','{','}'];

  // Check if the data contains any dangerous symbols
  for (var i = 0; i < dangerousSymbols.length; i++) {
    if (data.includes(dangerousSymbols[i])) {
      if( abort === false){
        error = "The following list of banned symbols was used as data entry: ";
      }
      error = error + dangerousSymbols[i] + " ";
      abort = true;
    }
  }
  return {abort: abort, error: error};
}

