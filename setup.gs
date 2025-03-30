/***************************************************
 * onOpen: スプレッドシートを開いたときのメニュー追加
 ***************************************************/
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("ClassroomAPI 一括採点")
    .addItem("シート初期設定", "setupSheets")
    .addItem("【一括】classroomの情報を取得", "updateClassroomDataAll")
    .addItem("評価対象クラスの課題を更新", "updateClassroomDataForEvaluation")
    .addSeparator()
    .addItem("課題を一括作成", "bulkCreateAssignments")
    .addItem("提出物を一括取得", "bulkFetchSubmissionsFromEvaluation")
    .addSeparator()
    .addItem("評価送信(得点のみ)", "bulkUpdateGrades")
    .addItem("評価送信＆返却", "bulkUpdateGradesAndReturn")
    .addToUi();
}


/***************************************************
 * setupSheets: 各シートの初期化
 *   ※実行前に「内容が消去されます」とアラート表示
 ***************************************************/
function setupSheets() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "初期設定の警告",
    "この操作を実行すると、すべてのシートの内容が消去されます。よろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) {
    Logger.log("初期設定をキャンセルしました。");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) CourseListシート: A="courseName", B="courseId"
  let courseList = ss.getSheetByName("CourseList");
  if (!courseList) {
    courseList = ss.insertSheet("CourseList");
  }
  courseList.clear();
  courseList.getRange("A1").setValue("courseName");
  courseList.getRange("B1").setValue("courseId");

  // 2) AssignmentListシート: A="courseName", B="courseWorkId", C="courseWorkTitle", D="maxPoints"
  let assignmentList = ss.getSheetByName("AssignmentList");
  if (!assignmentList) {
    assignmentList = ss.insertSheet("AssignmentList");
  }
  assignmentList.clear();
  assignmentList.getRange("A1").setValue("courseName");
  assignmentList.getRange("B1").setValue("courseWorkId");
  assignmentList.getRange("C1").setValue("courseWorkTitle");
  assignmentList.getRange("D1").setValue("maxPoints");

  // 3) AssignmentCreationシート
  let creationSheet = ss.getSheetByName("AssignmentCreation");
  if (!creationSheet) {
    creationSheet = ss.insertSheet("AssignmentCreation");
  }
  creationSheet.clear();
  creationSheet.getRange("A1").setValue("コース名");
  creationSheet.getRange("B1").setValue("課題名");
  creationSheet.getRange("C1").setValue("配点");
  creationSheet.getRange("D1").setValue("課題の説明");
  creationSheet.getRange("E1").setValue("課題ID(作成後)");

  // A列(2~101行)に「リストを範囲で指定」でコース名プルダウンを設定
  setCourseNameDropdownForCreation(creationSheet);

  // 4) Submissionsシート
  let subsSheet = ss.getSheetByName("Submissions");
  if (!subsSheet) {
    subsSheet = ss.insertSheet("Submissions");
  }
  subsSheet.clear();
  subsSheet.getRange("A1").setValue("courseName");
  subsSheet.getRange("B1").setValue("assignmentId");
  subsSheet.getRange("C1").setValue("assignmentName");
  subsSheet.getRange("D1").setValue("maxPoints");
  subsSheet.getRange("E1").setValue("userId");
  subsSheet.getRange("F1").setValue("studentName");
  subsSheet.getRange("G1").setValue("submissionId");
  subsSheet.getRange("H1").setValue("state");
  subsSheet.getRange("I1").setValue("updateTime");
  subsSheet.getRange("J1").setValue("assignedGrade");
  subsSheet.getRange("K1").setValue("attachments");
  subsSheet.getRange("L1").setValue("inputScore");
  subsSheet.getRange("M1").setValue("comment");

  // 5) Evaluationシート
  let evaluationSheet = ss.getSheetByName("Evaluation");
  if (!evaluationSheet) {
    evaluationSheet = ss.insertSheet("Evaluation");
  }
  evaluationSheet.clear();
  evaluationSheet.getRange("A1").setValue("コース名");
  evaluationSheet.getRange("B1").setValue("課題名");

  // A列(2~101行)に「リストを範囲で指定」でコース名プルダウンを設定
  setCourseNameDropdownForEvaluation(evaluationSheet);

  // 6) FilterHelperシート: 課題名の一時書き出し領域などに使う
  let filterSheet = ss.getSheetByName("FilterHelper");
  if (!filterSheet) {
    filterSheet = ss.insertSheet("FilterHelper");
  }
  filterSheet.clear();
  filterSheet.getRange("A1").setValue("Temporary usage for DataValidation");
  // このシートのZ列などをonEditで一時書き込み領域として使う

  SpreadsheetApp.getUi().alert("シート初期設定が完了しました。");
}


/***************************************************
 * setCourseNameDropdownForCreation:
 *   AssignmentCreationシート A列(2~101行)を
 *   "リストを範囲で指定" でコース名プルダウンにする
 ***************************************************/
function setCourseNameDropdownForCreation(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clSheet = ss.getSheetByName("CourseList");
  if (!clSheet) return;
  // コース名がA2:A1000に入っていると想定
  const rangeForCourseNames = clSheet.getRange("A2:A1000");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeForCourseNames, true)  // ドロップダウンを表示
    .build();
  for (let r = 2; r <= 101; r++) {
    let cell = sheet.getRange(r, 1); // A列
    cell.clearDataValidations();
    cell.setDataValidation(rule);
  }
}


/***************************************************
 * setCourseNameDropdownForEvaluation:
 *   Evaluationシート A列(2~101行)を
 *   "リストを範囲で指定" でコース名プルダウンにする
 ***************************************************/
function setCourseNameDropdownForEvaluation(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clSheet = ss.getSheetByName("CourseList");
  if (!clSheet) return;
  const rangeForCourseNames = clSheet.getRange("A2:A1000");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeForCourseNames, true)
    .build();
  for (let r = 2; r <= 101; r++) {
    let cell = sheet.getRange(r, 1);
    cell.clearDataValidations();
    cell.setDataValidation(rule);
  }
}


/***************************************************
 * onEdit: EvaluationシートのA列が編集されたら、
 *   該当行のB列に「リストを範囲で指定」データバリデーションを設定。
 *   ただし候補が多い場合(>500)でもOKなように、
 *   FilterHelperシートに一時書き込み → requireValueInRange で参照
 ***************************************************/
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Evaluation") return;
  // A列（列=1）、2行目以降
  if (e.range.getColumn() === 1 && e.range.getRow() >= 2) {
    const courseName = e.range.getValue();
    const row = e.range.getRow();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assignmentListSheet = ss.getSheetByName("AssignmentList");
    const filterSheet = ss.getSheetByName("FilterHelper");
    if (!assignmentListSheet || !filterSheet) return;

    // コース名に合致する課題名リストを取得
    const data = assignmentListSheet.getRange("A2:C1000").getValues(); 
    // A=courseName, B=workId, C=courseTitle
    let tasks = [];
    data.forEach(function(r) {
      if (r[0] === courseName && r[2]) {
        tasks.push(r[2]);
      }
    });

    const targetCell = sheet.getRange(row, 2); // B列
    targetCell.clearDataValidations();
    targetCell.clearContent();

    // FilterHelperシートのZ列(1~1000行)を一時書き込み領域に使う
    const tempRange = filterSheet.getRange("Z1:Z1000");
    tempRange.clearContent();

    if (tasks.length > 0) {
      // 書き込み
      for (let i = 0; i < tasks.length; i++) {
        filterSheet.getRange(1 + i, 26).setValue(tasks[i]); // Z1 + i
      }
      // tasks.length 行だけが課題リストとして書き込まれた
      const dvRange = filterSheet.getRange(1, 26, tasks.length, 1); // Z1:Z(tasks.length)
      // requireValueInRange(dvRange, true) でプルダウン
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(dvRange, true)
        .build();
      targetCell.setDataValidation(rule);
    }
  }
}
