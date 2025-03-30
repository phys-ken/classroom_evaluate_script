/***************************************************
 * (A) 【一括】classroomの情報を取得
 *    全クラス → CourseList (A=courseName, B=courseId)
 *    各クラスの課題 → AssignmentList (A=courseName, B=courseWorkId, C=courseWorkTitle, D=maxPoints)
 ***************************************************/
function updateClassroomDataAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let courses = fetchAllCourses();
  writeCourseList(courses);
  
  let assignmentArr = [];
  courses.forEach(c => {
    let works = fetchCourseWorks(c.id);
    if (works && works.length) {
      works.forEach(w => {
        assignmentArr.push({
          courseName: c.name,
          workId: w.id,
          title: w.title,
          maxPoints: (w.maxPoints !== undefined ? w.maxPoints : "")
        });
      });
    }
  });
  writeAssignmentList(assignmentArr);
  SpreadsheetApp.getUi().alert("全クラスの情報を更新しました。");
}


/***************************************************
 * (B) 評価対象クラスの課題を更新
 *    Evaluationシートの各行のコース名に対して、
 *    AssignmentListシートの内容（ヘッダー以外）を完全に初期化し、
 *    該当する課題情報を取得して再記入する
 ***************************************************/
function updateClassroomDataForEvaluation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  if (!evalSheet) {
    SpreadsheetApp.getUi().alert("Evaluationシートがありません。");
    return;
  }
  
  // 変更点: AssignmentListシートの全内容（ヘッダー以外）を完全にクリア
  const assignmentList = ss.getSheetByName("AssignmentList");
  if (assignmentList) {
    assignmentList.getRange("A2:D").clearContent();
  }
  
  let lastRow = evalSheet.getLastRow();
  let totalAssignments = 0;
  // Evaluationシートの各行（2行目以降）について処理
  for (let row = 2; row <= lastRow; row++) {
    let courseName = evalSheet.getRange(row, 1).getValue();
    if (!courseName) continue;
    const cId = getCourseIdByName(courseName);
    if (!cId) continue;
    let works;
    try {
      works = fetchCourseWorks(cId);
    } catch(e) {
      Logger.log("課題一覧取得失敗: " + e);
      continue;
    }
    if (!works || works.length === 0) continue;
    // 各評価対象の行ごとに、AssignmentListに追記する
    let assignmentArr = [];
    works.forEach(w => {
      assignmentArr.push({
        courseName: courseName,
        workId: w.id,
        title: w.title,
        maxPoints: (w.maxPoints !== undefined ? w.maxPoints : "")
      });
    });
    writeAssignmentListPartial(assignmentArr);
    totalAssignments += assignmentArr.length;
  }
  SpreadsheetApp.getUi().alert(`評価対象クラスの課題情報を更新しました（全${totalAssignments}件）。`);
}


/***************************************************
 * removeAssignmentListRows: 指定したコース名の行をAssignmentListから削除
 * （今回はupdateClassroomDataForEvaluationでシート全体をクリアするので未使用）
 ***************************************************/
function removeAssignmentListRows(courseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AssignmentList");
  if (!sh) return;
  let last = sh.getLastRow();
  if (last < 2) return;
  let data = sh.getRange(2, 1, last - 1, 4).getValues();
  let toRemove = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === courseName) {
      toRemove.push(i + 2);
    }
  }
  toRemove.reverse().forEach(r => {
    sh.deleteRow(r);
  });
}

/** writeAssignmentListPartial: AssignmentListに追記（既存行を消さず下に追加） */
function writeAssignmentListPartial(arr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AssignmentList");
  if (!sh) return;
  let last = sh.getLastRow();
  arr.forEach(a => {
    last++;
    sh.getRange(last, 1).setValue(a.courseName);
    sh.getRange(last, 2).setValue(a.workId);
    sh.getRange(last, 3).setValue(a.title);
    sh.getRange(last, 4).setValue(a.maxPoints);
  });
}

/***************************************************
 * fetchAllCourses: Classroom APIで全クラス取得 => [{id, name}, ...]
 ***************************************************/
function fetchAllCourses() {
  let all = [];
  let pageToken = "";
  do {
    let resp = Classroom.Courses.list({ pageToken });
    if (resp.courses && resp.courses.length) {
      all = all.concat(resp.courses);
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return all.map(c => ({ id: c.id, name: c.name }));
}

/***************************************************
 * fetchCourseWorks: 指定クラスの課題一覧取得
 ***************************************************/
function fetchCourseWorks(cId) {
  let all = [];
  let pageToken = "";
  do {
    let resp = Classroom.Courses.CourseWork.list(cId, { pageToken });
    if (resp.courseWork && resp.courseWork.length) {
      all = all.concat(resp.courseWork);
    }
    pageToken = resp.nextPageToken;
  } while (pageToken);
  return all;
}

/***************************************************
 * writeCourseList: CourseListシートに書き込み (A=courseName, B=courseId)
 ***************************************************/
function writeCourseList(courses) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CourseList");
  if (!sh) return;
  sh.getRange("A2:B").clearContent();
  let row = 2;
  courses.forEach(c => {
    sh.getRange(row, 1).setValue(c.name);
    sh.getRange(row, 2).setValue(c.id);
    row++;
  });
}

/***************************************************
 * writeAssignmentList: AssignmentListシートに書き込み
 *    (A=courseName, B=workId, C=courseWorkTitle, D=maxPoints)
 ***************************************************/
function writeAssignmentList(arr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AssignmentList");
  if (!sh) return;
  sh.getRange("A2:D").clearContent();
  let row = 2;
  arr.forEach(a => {
    sh.getRange(row, 1).setValue(a.courseName);
    sh.getRange(row, 2).setValue(a.workId);
    sh.getRange(row, 3).setValue(a.title);
    sh.getRange(row, 4).setValue(a.maxPoints);
    row++;
  });
}

/***************************************************
 * (C) 複数クラスへの一括課題作成: AssignmentCreationシート
 *    行1はヘッダー、行2～101が入力領域
 ***************************************************/
function bulkCreateAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AssignmentCreation");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("AssignmentCreationシートがありません。");
    return;
  }
  let countCreated = 0;
  for (let row = 2; row <= 101; row++) {
    const cName = sheet.getRange(row, 1).getValue();
    const title = sheet.getRange(row, 2).getValue();
    const pts = sheet.getRange(row, 3).getValue();
    const desc = sheet.getRange(row, 4).getValue();
    if (!cName || !title) continue;
    const cId = getCourseIdByName(cName);
    if (!cId) {
      Logger.log(`Row ${row}: コース名 "${cName}" が見つからずスキップ`);
      continue;
    }
    try {
      let newId = createAssignment(cId, title, desc || "", pts || 0);
      sheet.getRange(row, 5).setValue(newId);
      countCreated++;
      appendToAssignmentList(cName, newId, title, pts || 0);
    } catch (e) {
      Logger.log(`Row ${row} 課題作成失敗: ${e}`);
    }
  }
  SpreadsheetApp.getUi().alert(`一括課題作成完了: ${countCreated}件`);
}

function createAssignment(courseId, title, desc, maxPoints) {
  Classroom.Courses.get(courseId);
  let cw = {
    title: title,
    description: desc,
    maxPoints: maxPoints,
    state: "PUBLISHED",
    workType: "ASSIGNMENT"
  };
  let res = Classroom.Courses.CourseWork.create(cw, courseId);
  if (!res || !res.id) throw new Error("課題IDを取得できません");
  return res.id;
}

function appendToAssignmentList(cName, wId, title, maxPoints) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AssignmentList");
  if (!sh) return;
  let last = sh.getLastRow() + 1;
  sh.getRange(last, 1).setValue(cName);
  sh.getRange(last, 2).setValue(wId);
  sh.getRange(last, 3).setValue(title);
  sh.getRange(last, 4).setValue(maxPoints);
}

/** getCourseIdByName: CourseList (A=courseName, B=courseId)から取得 */
function getCourseIdByName(cName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cl = ss.getSheetByName("CourseList");
  if (!cl) return null;
  let last = cl.getLastRow();
  let data = cl.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === cName) {
      return data[i][1];
    }
  }
  return null;
}

/***************************************************
 * (D) 複数クラス・複数課題の提出物一括取得:
 *    Evaluationシートの各行(2～101)のコース名,課題名を参照して
 *    Submissionsシートに追記
 *    Submissionsシートの列:
 *      A=courseName, B=assignmentId, C=assignmentName, D=maxPoints,
 *      E=userId, F=studentName, G=submissionId, H=state, I=updateTime,
 *      J=assignedGrade, K=attachments, L=inputScore, M=comment
 ***************************************************/
function bulkFetchSubmissionsFromEvaluation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  const subsSheet = ss.getSheetByName("Submissions");
  if (!evalSheet || !subsSheet) {
    SpreadsheetApp.getUi().alert("EvaluationまたはSubmissionsシートがありません。");
    return;
  }
  let last = evalSheet.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert("Evaluationシートに対象データがありません。");
    return;
  }
  let data = evalSheet.getRange(2, 1, last - 1, 2).getValues(); // [courseName, assignmentTitle]
  let countFetched = 0;
  data.forEach(row => {
    const cName = row[0];
    const aTitle = row[1];
    if (!cName || !aTitle) return;
    const cId = getCourseIdByName(cName);
    if (!cId) return;
    const aId = getWorkIdByCourseAndTitle(cName, aTitle);
    if (!aId) return;
    let subArr;
    try {
      let resp = Classroom.Courses.CourseWork.StudentSubmissions.list(cId, aId);
      subArr = resp.studentSubmissions;
    } catch (e) {
      Logger.log(`取得失敗: ${cName}, assignmentId=${aId}: ${e}`);
      return;
    }
    if (!subArr || subArr.length === 0) return;
    subArr.forEach(s => {
      let stName = getStudentName(cId, s.userId);
      // maxPointsはAssignmentListから取得
      const mp = getMaxPointsFromAssignmentList(cName, aTitle);
      appendSubmissionRow(subsSheet, cName, aId, aTitle, mp, s, stName);
      countFetched++;
    });
  });
  SpreadsheetApp.getUi().alert(`提出物一括取得完了: ${countFetched}件`);
}

/** getWorkIdByCourseAndTitle: AssignmentList (A=courseName, B=workId, C=title)から課題IDを取得 */
function getWorkIdByCourseAndTitle(cName, aTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AssignmentList");
  if (!sh) return null;
  let last = sh.getLastRow();
  let data = sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === cName && data[i][2] === aTitle) {
      return data[i][1];
    }
  }
  return null;
}

/** getMaxPointsFromAssignmentList: AssignmentList (D列)から満点を取得 */
function getMaxPointsFromAssignmentList(cName, aTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("AssignmentList");
  if (!sh) return "";
  let last = sh.getLastRow();
  let data = sh.getRange(2, 1, last - 1, 4).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === cName && data[i][2] === aTitle) {
      return data[i][3];
    }
  }
  return "";
}

/***************************************************
 * appendSubmissionRow: Submissionsシートに1行追加
 *   列: A=courseName, B=assignmentId, C=assignmentName, D=maxPoints,
 *        E=userId, F=studentName, G=submissionId, H=state, I=updateTime,
 *        J=assignedGrade, K=attachments, L=inputScore, M=comment
 ***************************************************/
function appendSubmissionRow(sh, cName, aId, aTitle, mp, sub, stName) {
  let last = sh.getLastRow() + 1;
  sh.getRange(last, 1).setValue(cName);
  sh.getRange(last, 2).setValue(aId);
  sh.getRange(last, 3).setValue(aTitle);
  sh.getRange(last, 4).setValue(mp || "");
  sh.getRange(last, 5).setValue(sub.userId || "");
  sh.getRange(last, 6).setValue(stName || "");
  sh.getRange(last, 7).setValue(sub.id || "");
  sh.getRange(last, 8).setValue(sub.state || "");
  sh.getRange(last, 9).setValue(sub.updateTime || "");
  const ag = (sub.assignedGrade !== undefined) ? sub.assignedGrade : "";
  sh.getRange(last, 10).setValue(ag);
  sh.getRange(last, 11).setValue(getAttachmentsStr(sub));
  // inputScore: assignedGradeがあればその値を初期設定
  sh.getRange(last, 12).setValue(ag !== "" ? ag : "");
  sh.getRange(last, 13).setValue("");
}

function getStudentName(cId, userId) {
  try {
    let resp = Classroom.Courses.Students.get(cId, userId);
    if (resp.profile && resp.profile.name) {
      return resp.profile.name.fullName;
    }
  } catch(e) {}
  return "";
}

function getAttachmentsStr(sub) {
  if (!sub.assignmentSubmission) return "";
  if (!sub.assignmentSubmission.attachments) return "";
  let arr = [];
  sub.assignmentSubmission.attachments.forEach(a => {
    if (a.driveFile && a.driveFile.alternateLink) {
      arr.push(a.driveFile.alternateLink);
    } else if (a.link && a.link.url) {
      arr.push(a.link.url);
    }
  });
  return arr.join("; ");
}

/***************************************************
 * (E) 評価送信(得点のみ) および 返却: Submissionsシート一括処理
 *   送信前にダイアログで「新規評価／評価修正／未評価」の件数を確認
 ***************************************************/
function bulkUpdateGrades() {
  doUpdateGrades(false);
}
function bulkUpdateGradesAndReturn() {
  doUpdateGrades(true);
}

function doUpdateGrades(shouldReturn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subs = ss.getSheetByName("Submissions");
  if (!subs) {
    SpreadsheetApp.getUi().alert("Submissionsシートがありません。");
    return;
  }
  let last = subs.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert("Submissionsにデータがありません。");
    return;
  }
  let data = subs.getRange(2, 1, last - 1, 13).getValues(); // A～M
  // 正しいインデックス: A=0, B=1, C=2, D=3, E=4, F=5, G=submissionId (index 6),
  // H=7, I=8, J=assignedGrade (index 9), K=10, L=inputScore (index 11), M=12
  let countNew = 0, countModify = 0, countNoScore = 0;
  data.forEach(r => {
    const cName = r[0];
    const aId   = r[1];
    const subId = r[6];  // submissionId from G列
    if (!cName || !aId || !subId) return;
    const assigned = r[9]; // assignedGrade from J列
    const input    = r[11]; // inputScore from L列
    if (assigned === "" && input === "") {
      countNoScore++;
    } else if (assigned === "" && input !== "") {
      countNew++;
    } else if (assigned !== "" && input !== "" && assigned != input) {
      countModify++;
    }
  });
  const ui = SpreadsheetApp.getUi();
  const msg = `新規評価: ${countNew}人\n評価修正: ${countModify}人\n未評価: ${countNoScore}人\nこれらを送信しますか？`;
  let resp = ui.alert("評価送信確認", msg, ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) {
    Logger.log("評価送信キャンセル");
    return;
  }
  let countUpd = 0, countRet = 0;
  data.forEach(r => {
    const cName = r[0];
    const aId   = r[1];
    const subId = r[6];
    if (!cName || !aId || !subId) return;
    let input = r[11];
    if (input === "" || isNaN(input)) return;
    const cId = getCourseIdByName(cName);
    if (!cId) return;
    let body = { assignedGrade: input, draftGrade: input };
    try {
      Classroom.Courses.CourseWork.StudentSubmissions.patch(
        body, cId, aId, subId, { updateMask: "assignedGrade,draftGrade" }
      );
      countUpd++;
    } catch (e) {
      Logger.log("patch fail: " + e);
      return;
    }
    if (shouldReturn) {
      try {
        Classroom.Courses.CourseWork.StudentSubmissions['return'](
          {}, cId, aId, subId
        );
        countRet++;
      } catch (e) {
        Logger.log("return fail: " + e);
      }
    }
  });
  if (shouldReturn) {
    ui.alert(`得点更新: ${countUpd}件, 返却: ${countRet}件 完了`);
  } else {
    ui.alert(`得点更新のみ完了: ${countUpd}件`);
  }
}
