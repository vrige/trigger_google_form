// this function is periodacally call by a trigger to check if the incoming-form list is empty or not
function time(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // get the sheet incoming frequencies
  var sheet_Incoming_freq_obj = getSheetByItsName(ss, "Incoming Frequencies");
  var sheet_Incoming_freq = sheet_Incoming_freq_obj["sheet"];

  // copy the row in the destination sheet "Frequenze"
  var sheet_frequency_obj = getSheetByItsName(ss, "Frequenze");
  var sheet_frequency = sheet_frequency_obj["sheet"];

  // assumption: the correct names of the columns are in the first row
  // assumption: the number of column is fix (if the number of questions in form is modified, then it must be modified)
  var fieldNames = sheet_Incoming_freq.getRange(1, 1, 1, 11).getValues()[0];
  console.log("fieldNames: " + fieldNames);

  var lastRow = sheet_Incoming_freq.getLastRow();
  console.log("lastRow: " + lastRow);

  if ( sheet_Incoming_freq_obj["abort"] === false && sheet_frequency_obj["abort"] === false ){

    // Check if there are rows to process 
    if (lastRow > 1) {

      console.log("There are incoming forms");

      for (var i = lastRow; i >= 2; i--){

        // copy the target row in the "Frequenze" sheet
        var row = updateMainTab(i, sheet_Incoming_freq, sheet_frequency);
        console.log("row \"incomingFrequencies\": " + i + ", row Frequenze: " + row);

        // start the main routine on the row
        routine_function(ss, sheet_frequency, row, fieldNames);

      }
      
    }else{
      console.log("There are no incoming forms");
    }

  }else{
    console.log("At least one the two main pages doesn't exist. The form is aborted");

    // elaborate message error for the first layer on missing sheet(s)
    var error = "";
    var solution = "";
    if(sheet_Incoming_freq_obj["abort"] === true){
      error = error + sheet_Incoming_freq_obj["error"] + "\n";
      solution = solution + "Controllare che esista il foglio \'" + sheet_Incoming_freq_obj["name"]  + 
                 "\' e che il nome sia corretto.\n";
    }
    if(sheet_frequency_obj["abort"] === true){
      error = error + sheet_frequency_obj["error"] + "\n";
      solution = solution + "Controllare che esista il foglio \'" + sheet_frequency_obj["name"] + 
                "\' e che il nome sia corretto.\n";
    }
  
    console.log(error)
    console.log(solution)

    //updateState(sheet_frequency_obj, row_sheet, error, solution);
  }
}

// copy target row from source sheet to destination sheet
function updateMainTab(row, source_sheet, dest_sheet){

  // get the row from the source sheet
  var targetRange = source_sheet.getRange(row, 1, 1, source_sheet.getLastColumn()).getValues()[0];

  // append the new row in the destination sheet
  dest_sheet.appendRow(targetRange);

  // remove the row from the source sheet
  source_sheet.deleteRow(row);

  return dest_sheet.getLastRow();
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

// get the column with the word specified in "searchString". If you need to discriminate two or more columns,
// it is possible to specified some words to avoid in in the parameter "avoid"
function getDataFromColumn(e, fieldNames, searchString, avoid="-"){
  var name = 0;
  var abort = false;
  var error = "";
  var check = false;
  var values = "";

  // look for the column with the searchStrign inside, but not avoid
  for (var i = 0; i < fieldNames.length; i++){
    console.log("fieldNames["+ i + "]: " + fieldNames[i]);
    console.log("e["+ i + "]: " + e[i]);
    if (fieldNames[i].includes(searchString) && fieldNames[i].trim() !== avoid){
      if (e[i] !== ""){
        name = fieldNames[i];
        values = e[i]; 
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
    console.log("Found column with the words \'"+ searchString + "\' in: ", values);
    if(searchString === "Studenti" || searchString === "Nome e Cognome docente"){  
      // in case of the students or professors, split the string in a list of students
      values = values.split(",");
      //console.log("students: "+  values);
    }else{
      values = values.trim();
    }
    if (searchString === "Durata della lezione"){
      if (values > 3){
        check = true;
        error = "Please check why the lesson last more than 3 hours.\n";
      }
    }
  }
 
  return {abort : abort, error: error, check: check, name: searchString, values: values};
}

// this function is the real automatization function
function routine_function(ss, sheet_frequency, row, fieldNames){

  // get the values of the row
  var e = sheet_frequency.getRange(row,1,1,sheet_frequency.getLastColumn()).getValues()[0];

  // get the name of the subject 
  var name_of_the_course = e[2];

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

  // first layer of error handling: checking that the useful following pages exist:
  // "Studenti extra", name_of_the_course, "Riassunto lezioni " + name_of_the_course
  if( sheet_subject_obj["abort"] === false && sheet_subject_lesson_obj["abort"] === false && 
    sheet_extraStudents_obj["abort"] === false){
    
    console.log("The 3 useful pages exist");

    // get the professor name and surname
    var professor_obj = getDataFromColumn(e, fieldNames, "Nome e Cognome docente");
    var professor = professor_obj["values"];

    // get the date of the lesson
    var date_of_the_lesson_obj = getDataFromColumn(e, fieldNames, "Data della lezione");
    var date_of_the_lesson = date_of_the_lesson_obj["values"];

    // get the duration of the lesson
    var duration_obj = getDataFromColumn(e, fieldNames, "Durata della lezione");
    var duration = duration_obj["values"];

    // get data from the column "Nessuno studente è presente"
    var no_students_obj = getDataFromColumn(e, fieldNames, "Nessuno studente è presente");
    var no_students = no_students_obj["values"];

    // if the professor has selected the option "Sì" on "Nessuno studente è presente"
    // then it is possible to avoid the check on the student column
    console.log("nessuno studente è presente: ", no_students)

    // get the list of students 
    var names_obj = {values: [], abort: false, error: "", name: "Nessuno studente è presente",check: true};
    var names = names_obj["values"];
    if (no_students === "No"){
      names_obj = getDataFromColumn(e, fieldNames, "Studenti", "Studenti extra");
      names = names_obj["values"];
    }

    // second layer of error handling: checking form columns
    if(professor_obj["abort"] === false && names_obj["abort"] === false &&
    date_of_the_lesson_obj["abort"] === false && duration_obj["abort"] === false && 
    no_students_obj["abort"] ===false){

      console.log("All the necessary form columns are there");

      console.log("professor: ", professor);
      console.log("date_of_the_lesson: ", date_of_the_lesson);
      console.log("duration: ", duration);
      console.log("no_students: ", no_students);
      console.log("no_students: ", names);


    } else{

      console.log("There was some problems with the form columns");

      // elaborate message error for the second layer: missing form columns
      var error = "";
      var solution = "";

      if(professor_obj["abort"] === true){
        error = error + professor_obj["error"] + "\n";
        solution = solution + "Controllare che esista una colonna del form che contenga la parola \'" + professor_obj["name"] + "\'  e che il nome sia corretto.\n";
      }
      if(names_obj["abort"] === true){
        error = error + names_obj["error"] + "\n";
        solution = solution + "Controllare che esista una colonna del form che contenga la parola \'" + 
        names_obj["name"] + "\' e che il nome sia corretto.\n";
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
      if(no_students_obj["abort"] === true){
        error = error + no_students_obj["error"] + "\n";
        solution = solution + "Controllare che esista una colonna del form che contenga la parola \'" + no_students_obj["name"] + "\'  e che il nome sia corretto.\n";
      }
      solution = solution + "Per controllare che il nome sia correto andare sulla corrispondente form e controllare \n" + "che vi sia una domanda con la/e parola/e chiave/i precedentemente specificata/e.";
    
      console.log(error)
      console.log(solution)

      //updateState(sheet_frequency_obj, row_sheet, error, solution);

    }
  } else{

    console.log("At least one the two useful pages doesn't exist. The form is aborted");

    // elaborate message error for the first layer on missing sheet(s)
    var error = "";
    var solution = "";
    
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
              " e fare un copia e incolla dalla lista.";
  
    console.log(error)
    console.log(solution)

    //updateState(sheet_frequency_obj, row_sheet, error, solution);
  }
}
