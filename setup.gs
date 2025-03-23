/***************************************************
 * onOpen(e): スプレッドシートを開いた際のメニュー作成
 ***************************************************/
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("ClassroomAPI 一括採点")
    .addItem("シート初期設定", "setupSheets")
    .addItem("課題作成", "confirmAndCreateAssignment")
    .addSeparator()
    .addItem("（評価）課題一覧取得", "fetchAssignmentsForSelectedClass")
    .addItem("（評価）提出一覧取得", "fetchSubmissionsForSelectedAssignment")
    .addItem("（評価）評価送信 (得点のみ)", "updateStudentGrades")       // 返却せず得点のみ
    .addItem("（評価）評価送信＆返却", "updateStudentGradesAndReturn") // 返却も行う
    .addToUi();
}


/***************************************************
 * シート初期設定: 必要なシートを作成し、Classroom APIからコース一覧を取得
 ***************************************************/
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) CourseListシート
  let courseListSheet = ss.getSheetByName("CourseList");
  if (!courseListSheet) {
    courseListSheet = ss.insertSheet("CourseList");
    courseListSheet.getRange("A1").setValue("courseId");
    courseListSheet.getRange("B1").setValue("courseName");
  }

  // 2) 課題作成用シート
  let creationSheet = ss.getSheetByName("AssignmentCreation");
  if (!creationSheet) {
    creationSheet = ss.insertSheet("AssignmentCreation");
    creationSheet.getRange("A1").setValue("コース名");
    creationSheet.getRange("B1").setValue("課題名");
    creationSheet.getRange("C1").setValue("配点");
    creationSheet.getRange("D1").setValue("課題の説明");
    creationSheet.getRange("A2").setValue("（プルダウン）");
    creationSheet.getRange("B2").setValue("API確認用課題");
    creationSheet.getRange("C2").setValue("10");
    creationSheet.getRange("D2").setValue("これはテスト課題です。");
  }

  // 3) 評価用シート
  let evaluationSheet = ss.getSheetByName("Evaluation");
  if (!evaluationSheet) {
    evaluationSheet = ss.insertSheet("Evaluation");
    evaluationSheet.getRange("A1").setValue("コース名");
    evaluationSheet.getRange("B1").setValue("課題名");
    evaluationSheet.getRange("A2").setValue("（プルダウン）");
    evaluationSheet.getRange("B2").setValue("（プルダウン）");
  }

  // 4) 課題一覧保管シート
  let assignmentListSheet = ss.getSheetByName("AssignmentList");
  if (!assignmentListSheet) {
    assignmentListSheet = ss.insertSheet("AssignmentList");
    assignmentListSheet.getRange("A1").setValue("courseId");
    assignmentListSheet.getRange("B1").setValue("courseWorkId");
    assignmentListSheet.getRange("C1").setValue("courseWorkTitle");
  }

  // 5) Submissionsシート
  let submissionsSheet = ss.getSheetByName("Submissions");
  if (!submissionsSheet) {
    submissionsSheet = ss.insertSheet("Submissions");
  }
  // 見出し行を設定（A～H列）
  submissionsSheet.clear();
  submissionsSheet.getRange(1,1).setValue("userId");
  submissionsSheet.getRange(1,2).setValue("studentName");
  submissionsSheet.getRange(1,3).setValue("submissionId");
  submissionsSheet.getRange(1,4).setValue("state");
  submissionsSheet.getRange(1,5).setValue("updateTime");
  submissionsSheet.getRange(1,6).setValue("assignedGrade");
  submissionsSheet.getRange(1,7).setValue("attachments");
  submissionsSheet.getRange(1,8).setValue("inputScore");

  // Classroom APIからコース一覧を取得し、CourseListに反映
  fetchAndListCourses();

  // コース名のプルダウンをAssignmentCreation!A2, Evaluation!A2に設定
  setCourseNameDropdown(creationSheet);
  setCourseNameDropdown(evaluationSheet);

  SpreadsheetApp.getUi().alert("初期セットアップ完了");
}


/***************************************************
 * fetchAndListCourses:
 *  Classroomからコース一覧を取得し、CourseListシート(A=ID,B=Name)に書き込む
 ***************************************************/
function fetchAndListCourses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CourseList");
  if (!sheet) throw new Error("CourseListシートがありません。");

  sheet.getRange("A2:B").clearContent();

  let response;
  try {
    response = Classroom.Courses.list({});
  } catch (e) {
    Logger.log("コース一覧取得失敗: " + e);
    return;
  }
  if (!response.courses || response.courses.length === 0) {
    Logger.log("利用可能なコースがありません。");
    return;
  }

  let row = 2;
  response.courses.forEach(course => {
    sheet.getRange(row,1).setValue(course.id);
    sheet.getRange(row,2).setValue(course.name);
    row++;
  });
  Logger.log("コース一覧をCourseListに反映しました。");
}


/***************************************************
 * setCourseNameDropdown:
 *  指定シートのA2セルにコース名のプルダウンを設定
 ***************************************************/
function setCourseNameDropdown(targetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const courseListSheet = ss.getSheetByName("CourseList");
  if (!courseListSheet) return;

  const lastRow = courseListSheet.getLastRow();
  if (lastRow < 2) return;

  const rangeForDropdown = courseListSheet.getRange(2, 2, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeForDropdown, true)
    .build();
  targetSheet.getRange("A2").setDataValidation(rule);
}
