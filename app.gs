/***************************************************
 * 課題作成フロー
 ***************************************************/
function confirmAndCreateAssignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AssignmentCreation");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("AssignmentCreationシートがありません。");
    return;
  }

  const courseName = sheet.getRange("A2").getValue();
  const title      = sheet.getRange("B2").getValue();
  const maxPoints  = sheet.getRange("C2").getValue();
  const desc       = sheet.getRange("D2").getValue();

  if (!courseName || !title) {
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未入力です。");
    return;
  }
  if (!maxPoints || isNaN(maxPoints)) {
    SpreadsheetApp.getUi().alert("配点が数値でありません。");
    return;
  }

  const courseId = getCourseIdByName(courseName);
  if (!courseId) {
    SpreadsheetApp.getUi().alert(`コース名「${courseName}」のIDが見つかりません。`);
    return;
  }

  // ダイアログ確認
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "課題作成の確認",
    `コース: ${courseName}\n課題名: ${title}\n配点: ${maxPoints}\n\nよろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) {
    Logger.log("課題作成をキャンセルしました。");
    return;
  }

  try {
    const newId = createAssignment(courseId, title, desc, maxPoints);
    ui.alert(`課題作成完了\n課題ID: ${newId}`);
  } catch(e) {
    ui.alert("課題作成中にエラー: " + e);
    Logger.log(e);
  }
}


/** Classroom API で課題を作成 */
function createAssignment(courseId, title, desc, maxPoints) {
  // コースが存在するかチェック (権限含む)
  Classroom.Courses.get(courseId);

  const courseWork = {
    title: title,
    description: desc || "",
    maxPoints: maxPoints,
    state: "PUBLISHED",
    workType: "ASSIGNMENT"
  };
  const created = Classroom.Courses.CourseWork.create(courseWork, courseId);
  if (!created || !created.id) {
    throw new Error("課題IDを取得できません");
  }
  return created.id;
}


/** コース名 → コースID変換 (CourseListシート参照) */
function getCourseIdByName(courseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CourseList");
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2,1,lastRow-1,2).getValues();  // A=courseId, B=courseName
  for (let i=0; i<data.length; i++){
    if (data[i][1] === courseName) {
      return data[i][0];
    }
  }
  return null;
}


////////////////////////////////////////////////////
// 評価フロー
////////////////////////////////////////////////////

/**
 * (評価) 課題一覧取得: EvaluationシートA2のコース名から課題一覧を取得→AssignmentListに書き込み
 */
function fetchAssignmentsForSelectedClass() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  if (!evalSheet) {
    SpreadsheetApp.getUi().alert("Evaluationシートが見つかりません。");
    return;
  }

  const courseName = evalSheet.getRange("A2").getValue();
  if (!courseName) {
    SpreadsheetApp.getUi().alert("コース名が未選択です。");
    return;
  }
  const courseId = getCourseIdByName(courseName);
  if (!courseId) {
    SpreadsheetApp.getUi().alert(`コース名「${courseName}」のIDが見つかりません。`);
    return;
  }

  const listSheet = ss.getSheetByName("AssignmentList");
  if (!listSheet) {
    SpreadsheetApp.getUi().alert("AssignmentListシートがありません。");
    return;
  }
  listSheet.getRange("A2:C").clearContent();

  let courseWorks;
  try {
    const resp = Classroom.Courses.CourseWork.list(courseId);
    courseWorks = resp.courseWork;
  } catch(e) {
    SpreadsheetApp.getUi().alert("課題一覧取得に失敗: " + e);
    return;
  }

  if (!courseWorks || courseWorks.length === 0) {
    SpreadsheetApp.getUi().alert("課題がありません。");
    return;
  }

  let row = 2;
  courseWorks.forEach(work => {
    listSheet.getRange(row,1).setValue(courseId);
    listSheet.getRange(row,2).setValue(work.id);
    listSheet.getRange(row,3).setValue(work.title);
    row++;
  });

  setAssignmentNameDropdown();
  SpreadsheetApp.getUi().alert("課題一覧を取得し、プルダウンを設定しました。");
}


/** 課題名(C列)をEvaluationシートB2にプルダウン設定 */
function setAssignmentNameDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  const listSheet = ss.getSheetByName("AssignmentList");
  if (!evalSheet || !listSheet) return;

  const lastRow = listSheet.getLastRow();
  if (lastRow < 2) return;

  const rangeForDropdown = listSheet.getRange(2,3,lastRow-1,1); // C列=課題タイトル
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rangeForDropdown, true)
    .build();
  evalSheet.getRange("B2").setDataValidation(rule);
}


/**
 * (評価) 提出一覧取得: コース名(A2)と課題名(B2)に応じてStudentSubmissionsを取得→Submissionsへ
 */
function fetchSubmissionsForSelectedAssignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  const subsSheet = ss.getSheetByName("Submissions");
  const listSheet = ss.getSheetByName("AssignmentList");
  if (!evalSheet || !subsSheet || !listSheet) {
    SpreadsheetApp.getUi().alert("必要なシートが見つかりません。");
    return;
  }

  // コースID, 課題IDを特定
  const courseName = evalSheet.getRange("A2").getValue();
  const workTitle  = evalSheet.getRange("B2").getValue();
  if (!courseName || !workTitle) {
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未選択です。");
    return;
  }

  const courseId = getCourseIdByName(courseName);
  if (!courseId) {
    SpreadsheetApp.getUi().alert(`コース名「${courseName}」のIDが見つかりません。`);
    return;
  }
  const workId = getCourseWorkIdByTitle(courseId, workTitle);
  if (!workId) {
    SpreadsheetApp.getUi().alert(`課題名「${workTitle}」のIDが見つかりません。`);
    return;
  }

  // Roster取得(userId-> studentName)
  const nameMap = getRosterMap(courseId);

  // 提出一覧取得
  let submissions;
  try {
    const resp = Classroom.Courses.CourseWork.StudentSubmissions.list(courseId, workId);
    submissions = resp.studentSubmissions;
  } catch(e) {
    SpreadsheetApp.getUi().alert("提出一覧取得失敗: " + e);
    return;
  }

  // Submissionsシートをクリアしてヘッダ再設定
  subsSheet.clear();
  subsSheet.getRange(1,1).setValue("userId");
  subsSheet.getRange(1,2).setValue("studentName");
  subsSheet.getRange(1,3).setValue("submissionId");
  subsSheet.getRange(1,4).setValue("state");
  subsSheet.getRange(1,5).setValue("updateTime");
  subsSheet.getRange(1,6).setValue("assignedGrade");
  subsSheet.getRange(1,7).setValue("attachments");
  subsSheet.getRange(1,8).setValue("inputScore");

  if (!submissions || submissions.length === 0) {
    SpreadsheetApp.getUi().alert("提出がありません。");
    return;
  }

  // 各Submissionを書き込み
  let row = 2;
  submissions.forEach(sub => {
    const userId = sub.userId || "";
    const submissionId = sub.id || "";
    const state = sub.state || "";
    const updateTime = sub.updateTime || "";
    const assignedGrade = (sub.assignedGrade !== undefined) ? sub.assignedGrade : "";
    const attachmentsStr = getAttachmentsStr(sub);

    subsSheet.getRange(row,1).setValue(userId);
    subsSheet.getRange(row,2).setValue(nameMap[userId] || "不明");
    subsSheet.getRange(row,3).setValue(submissionId);
    subsSheet.getRange(row,4).setValue(state);
    subsSheet.getRange(row,5).setValue(updateTime);
    subsSheet.getRange(row,6).setValue(assignedGrade);
    subsSheet.getRange(row,7).setValue(attachmentsStr);
    // H列(inputScore)は空欄のまま

    row++;
  });

  SpreadsheetApp.getUi().alert("Submissionsシートに提出状況を表示しました。");
}


/** AssignmentListシート (courseId + workTitle) → workId検索 */
function getCourseWorkIdByTitle(courseId, workTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("AssignmentList");
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2,1,lastRow-1,3).getValues();
  // A=courseId, B=workId, C=courseWorkTitle
  for (let i=0; i<data.length; i++){
    if (data[i][0] === courseId && data[i][2] === workTitle) {
      return data[i][1]; // workId
    }
  }
  return null;
}


/** コースのRosterを取得 → { userId: studentName } にまとめる */
function getRosterMap(courseId) {
  const nameMap = {};
  let pageToken = "";
  do {
    let resp = Classroom.Courses.Students.list(courseId, { pageToken: pageToken });
    if (resp.students && resp.students.length) {
      resp.students.forEach(stu => {
        const uid = stu.userId;
        const fullname = (stu.profile && stu.profile.name) ? stu.profile.name.fullName : "";
        nameMap[uid] = fullname;
      });
    }
    pageToken = resp.nextPageToken;
  } while(pageToken);

  return nameMap;
}


/** SubmissionのattachmentsをまとめてURL文字列を返す */
function getAttachmentsStr(sub) {
  if (!sub.assignmentSubmission) return "";
  if (!sub.assignmentSubmission.attachments) return "";

  const attachments = sub.assignmentSubmission.attachments;
  const links = [];

  attachments.forEach(att => {
    if (att.driveFile && att.driveFile.alternateLink) {
      links.push(att.driveFile.alternateLink);
    } else if (att.link && att.link.url) {
      links.push(att.link.url);
    }
    // 他にyoutubeVideo等もあり得る
  });

  return links.join("; ");
}


/***************************************************
 * (評価) 得点のみ更新 (返却しない)
 ***************************************************/
function updateStudentGrades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  const subsSheet = ss.getSheetByName("Submissions");
  if (!evalSheet || !subsSheet) {
    SpreadsheetApp.getUi().alert("必要なシートが見つかりません。");
    return;
  }

  const courseName = evalSheet.getRange("A2").getValue();
  const workTitle  = evalSheet.getRange("B2").getValue();
  if (!courseName || !workTitle) {
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未選択です。");
    return;
  }

  const courseId = getCourseIdByName(courseName);
  const workId   = getCourseWorkIdByTitle(courseId, workTitle);
  if (!courseId || !workId) {
    SpreadsheetApp.getUi().alert("コースID or 課題IDの特定に失敗しました。");
    return;
  }

  const lastRow = subsSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Submissionsシートに採点対象データがありません。");
    return;
  }

  const data = subsSheet.getRange(2,1,lastRow-1,8).getValues();
  // A=userId, B=studentName, C=submissionId, ... H=inputScore

  let countUpdated = 0;
  data.forEach(row => {
    const submissionId = row[2];  // C列
    const inputScore   = row[7];  // H列

    if (!submissionId) return;
    if (inputScore === "" || isNaN(inputScore)) return;

    const resource = {
      assignedGrade: inputScore,
      draftGrade: inputScore
    };
    try {
      Classroom.Courses.CourseWork.StudentSubmissions.patch(
        resource,
        courseId,
        workId,
        submissionId,
        { updateMask: "assignedGrade,draftGrade" }
      );
      countUpdated++;
    } catch(e) {
      Logger.log("更新失敗: " + e);
    }
  });

  SpreadsheetApp.getUi().alert(`得点の更新が完了しました。\n更新件数: ${countUpdated}`);
}


/***************************************************
 * (評価) 得点更新後に返却も行う
 ***************************************************/
function updateStudentGradesAndReturn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet = ss.getSheetByName("Evaluation");
  const subsSheet = ss.getSheetByName("Submissions");
  if (!evalSheet || !subsSheet) {
    SpreadsheetApp.getUi().alert("必要なシートが見つかりません。");
    return;
  }

  const courseName = evalSheet.getRange("A2").getValue();
  const workTitle  = evalSheet.getRange("B2").getValue();
  if (!courseName || !workTitle) {
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未選択です。");
    return;
  }

  const courseId = getCourseIdByName(courseName);
  const workId   = getCourseWorkIdByTitle(courseId, workTitle);
  if (!courseId || !workId) {
    SpreadsheetApp.getUi().alert("コースIDまたは課題IDの特定に失敗しました。");
    return;
  }

  const lastRow = subsSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Submissionsシートに採点対象データがありません。");
    return;
  }

  const data = subsSheet.getRange(2,1,lastRow - 1,8).getValues();
  // A=userId, B=studentName, C=submissionId, ... H=inputScore

  let countUpdated = 0;
  let countReturned = 0;

  data.forEach(row => {
    const submissionId = row[2];  // C列
    const inputScore   = row[7];  // H列

    if (!submissionId) return;
    if (inputScore === "" || isNaN(inputScore)) return;

    // 1) 成績をPATCH (draftGrade, assignedGrade)
    const resource = {
      assignedGrade: inputScore,
      draftGrade: inputScore
    };
    try {
      Classroom.Courses.CourseWork.StudentSubmissions.patch(
        resource,
        courseId,
        workId,
        submissionId,
        { updateMask: "assignedGrade,draftGrade" }
      );
      countUpdated++;
    } catch(e) {
      Logger.log("採点更新失敗: " + e);
      return; // この提出物の返却は行わない
    }

    // 2) 返却: 「courses.courseWork.studentSubmissions.return」
    //    Apps Scriptでは予約語のため bracket記法を使う
    try {
      Classroom.Courses.CourseWork.StudentSubmissions['return'](
        {}, // リクエストボディは空オブジェクト
        courseId,
        workId,
        submissionId
      );
      countReturned++;
    } catch(e) {
      Logger.log("返却失敗: " + e);
      // すでに返却済み or stateがTURNED_INでない等の可能性
    }
  });

  SpreadsheetApp.getUi().alert(
    `得点更新: ${countUpdated}件\n返却処理: ${countReturned}件 完了`
  );
}
