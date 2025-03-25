/***************************************************
 * onOpen(e): スプレッドシートを開いたときのメニュー
 ***************************************************/
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("ClassroomAPI 一括採点")
    .addItem("シート初期設定", "setupSheets")
    .addItem("classroomの情報を更新", "updateClassroomData")
    .addSeparator()
    .addItem("課題作成", "confirmAndCreateAssignment")
    .addSeparator()
    .addItem("（評価）提出一覧取得", "fetchSubmissionsForSelectedAssignment")
    .addItem("（評価）評価送信 (得点のみ)", "updateStudentGrades")
    .addItem("（評価）評価送信＆返却", "updateStudentGradesAndReturn")
    .addToUi();
}


/***************************************************
 * setupSheets():
 *   必要シートを作成＆初期化し、
 *   コース名/課題名のプルダウン連動を実装
 ***************************************************/
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) CourseListシート (A=courseName, B=courseId)
  let courseList = ss.getSheetByName("CourseList");
  if (!courseList) {
    courseList = ss.insertSheet("CourseList");
    courseList.getRange("A1").setValue("courseName");
    courseList.getRange("B1").setValue("courseId");
  }

  // 2) 課題作成用シート
  let creationSheet = ss.getSheetByName("AssignmentCreation");
  if (!creationSheet) {
    creationSheet = ss.insertSheet("AssignmentCreation");
    creationSheet.getRange("A1").setValue("コース名");
    creationSheet.getRange("B1").setValue("課題名");
    creationSheet.getRange("C1").setValue("配点");
    creationSheet.getRange("D1").setValue("課題の説明");
    creationSheet.getRange("A2").setValue("（プルダウン:コース名）");
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
    evaluationSheet.getRange("A2").setValue("（プルダウン:コース名）");
    evaluationSheet.getRange("B2").setValue("（連動プルダウン）");
  }

  // 4) AssignmentListシート (A=courseName, B=workId, C=title)
  let assignmentList = ss.getSheetByName("AssignmentList");
  if (!assignmentList) {
    assignmentList = ss.insertSheet("AssignmentList");
    assignmentList.getRange("A1").setValue("courseName");
    assignmentList.getRange("B1").setValue("courseWorkId");
    assignmentList.getRange("C1").setValue("courseWorkTitle");
  }

  // 5) Submissionsシート
  let submissionsSheet = ss.getSheetByName("Submissions");
  if (!submissionsSheet) {
    submissionsSheet = ss.insertSheet("Submissions");
  }
  submissionsSheet.clear();
  submissionsSheet.getRange(1,1).setValue("userId");
  submissionsSheet.getRange(1,2).setValue("studentName");
  submissionsSheet.getRange(1,3).setValue("submissionId");
  submissionsSheet.getRange(1,4).setValue("state");
  submissionsSheet.getRange(1,5).setValue("updateTime");
  submissionsSheet.getRange(1,6).setValue("assignedGrade");
  submissionsSheet.getRange(1,7).setValue("attachments");
  submissionsSheet.getRange(1,8).setValue("inputScore");

  // 6) FilterHelperシート (連動プルダウン用)
  let filterSheet = ss.getSheetByName("FilterHelper");
  if (!filterSheet) {
    filterSheet = ss.insertSheet("FilterHelper");
  } else {
    filterSheet.clear();
  }

  // ラベル
  filterSheet.getRange("A1").setValue("Selected CourseName");
  filterSheet.getRange("B1").setValue("Filtered Titles");

  // A2: =Evaluation!A2
  filterSheet.getRange("A2").setFormula("=Evaluation!A2");

  // B2: =IF(A2="","", FILTER(AssignmentList!C2:C, AssignmentList!A2:A=$A$2))
  filterSheet.getRange("B2").setFormula(
    '=IF(A2="","", FILTER(AssignmentList!C2:C, AssignmentList!A2:A=$A$2))'
  );

  // 7) コース名プルダウン
  setCourseNameDropdown(creationSheet.getRange("A2"));
  setCourseNameDropdown(evaluationSheet.getRange("A2"));

  // 8) 課題名プルダウン(Evaluation!B2) → FilterHelper!B2:B
  setAssignmentDropdownForEvaluation();

  SpreadsheetApp.getUi().alert(
    "初期セットアップ完了。\n" +
    "次に「classroomの情報を更新」を行ってから、EvaluationシートA2のコース名を選ぶとB2が連動します。"
  );
}


/***************************************************
 * setCourseNameDropdown(targetCell)
 *   - 常にデータバリデーションを設定し、CourseList(A列)を参照
 *   - CourseListにデータがなくても設定だけは行う
 ***************************************************/
function setCourseNameDropdown(targetCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const courseListSheet = ss.getSheetByName("CourseList");
  if (!courseListSheet) return;

  // A列=コース名全体を範囲指定(例: A2:A1000)
  // あるいはシート全体で安全に設定
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(courseListSheet.getRange("A2:A1000"), true)
    .build();

  targetCell.clearDataValidations();
  targetCell.setDataValidation(rule);
}


/***************************************************
 * setAssignmentDropdownForEvaluation():
 *   Evaluation!B2 → FilterHelper!B2:B
 ***************************************************/
function setAssignmentDropdownForEvaluation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  const filterSheet = ss.getSheetByName("FilterHelper");
  if (!evalSheet || !filterSheet) return;

  // B2~B1000あたりを参照
  const dvRange = filterSheet.getRange("B2:B1000");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(dvRange, true)
    .build();

  evalSheet.getRange("B2").clearDataValidations();
  evalSheet.getRange("B2").setDataValidation(rule);
}
