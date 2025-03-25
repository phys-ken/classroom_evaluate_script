/***************************************************
 * classroomの情報を更新
 *  - 全コース取得 → CourseList(A=コース名, B=コースID)
 *  - 全コースの課題 → AssignmentList(A=コース名,B=workId,C=課題名)
 ***************************************************/
function updateClassroomData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let courses = fetchAllCourses();  // => [{id, name}, ...]
  writeCourseList(courses);

  let assignments = [];
  courses.forEach(c => {
    let works = fetchCourseWorks(c.id);
    if (works && works.length) {
      works.forEach(w => {
        assignments.push({
          courseName: c.name,
          workId: w.id,
          title: w.title
        });
      });
    }
  });
  writeAssignmentList(assignments);

  SpreadsheetApp.getUi().alert(
    "クラス＆課題一覧を更新しました。\n" +
    "EvaluationシートA2にコース名が出るので選択→B2が連動します。"
  );
}


/***************************************************
 * fetchAllCourses():
 *   Classroom APIからコース一覧
 ***************************************************/
function fetchAllCourses() {
  let all = [];
  let pageToken="";
  do {
    let resp = Classroom.Courses.list({ pageToken });
    if(resp.courses && resp.courses.length){
      all = all.concat(resp.courses);
    }
    pageToken = resp.nextPageToken;
  }while(pageToken);

  // => [{id, name}, ...]
  return all.map(c => ({ id:c.id, name:c.name }));
}


/***************************************************
 * fetchCourseWorks(courseId):
 *   指定コースIDの課題一覧
 ***************************************************/
function fetchCourseWorks(courseId) {
  let allWorks=[];
  let pageToken="";
  do {
    let r = Classroom.Courses.CourseWork.list(courseId,{ pageToken });
    if(r.courseWork && r.courseWork.length){
      allWorks = allWorks.concat(r.courseWork);
    }
    pageToken= r.nextPageToken;
  } while(pageToken);

  return allWorks; // array of {id, title, ...}
}


/***************************************************
 * writeCourseList(courses):
 *   CourseList(A=コース名, B=コースID)
 ***************************************************/
function writeCourseList(courses) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CourseList");
  if(!sheet)return;

  sheet.getRange("A2:B").clearContent();
  let row=2;
  courses.forEach(c=>{
    sheet.getRange(row,1).setValue(c.name);
    sheet.getRange(row,2).setValue(c.id);
    row++;
  });
}


/***************************************************
 * writeAssignmentList( arr ):
 *   AssignmentList(A=コース名, B=workId, C=課題名)
 ***************************************************/
function writeAssignmentList(arr) {
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const sheet= ss.getSheetByName("AssignmentList");
  if(!sheet)return;

  sheet.getRange("A2:C").clearContent();
  let row=2;
  arr.forEach(a=>{
    sheet.getRange(row,1).setValue(a.courseName);
    sheet.getRange(row,2).setValue(a.workId);
    sheet.getRange(row,3).setValue(a.title);
    row++;
  });
}


/***************************************************
 * 課題作成 (AssignmentCreationシート)
 ***************************************************/
function confirmAndCreateAssignment() {
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const sheet= ss.getSheetByName("AssignmentCreation");
  if(!sheet){
    SpreadsheetApp.getUi().alert("AssignmentCreationシートがありません。");
    return;
  }

  const courseName= sheet.getRange("A2").getValue();
  const title     = sheet.getRange("B2").getValue();
  const maxPoints = sheet.getRange("C2").getValue();
  const desc      = sheet.getRange("D2").getValue();

  if(!courseName||!title){
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未入力です。");
    return;
  }
  if(!maxPoints||isNaN(maxPoints)){
    SpreadsheetApp.getUi().alert("配点が数値でありません。");
    return;
  }

  // コース名→コースID
  const courseId= getCourseIdByName(courseName);
  if(!courseId){
    SpreadsheetApp.getUi().alert(`コース名「${courseName}」のIDが見つかりません。`);
    return;
  }

  const ui= SpreadsheetApp.getUi();
  const resp= ui.alert(
    "課題作成の確認",
    `コース: ${courseName}\n課題名: ${title}\n配点: ${maxPoints}\n\nよろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if(resp!==ui.Button.OK){
    Logger.log("課題作成をキャンセルしました。");
    return;
  }

  try{
    const newId= createAssignment(courseId,title,desc,maxPoints);
    ui.alert(`課題作成完了！(ID: ${newId})`);
  }catch(e){
    ui.alert("課題作成中にエラー: "+ e);
  }
}


/** Classroom APIで課題作成 */
function createAssignment(courseId, title, desc, maxPoints) {
  // コースが有効か確認
  Classroom.Courses.get(courseId);

  const courseWork={
    title,
    description: desc||"",
    maxPoints,
    state:"PUBLISHED",
    workType:"ASSIGNMENT"
  };
  const created= Classroom.Courses.CourseWork.create(courseWork, courseId);
  if(!created||!created.id){
    throw new Error("課題IDを取得できません。");
  }
  return created.id;
}


/***************************************************
 * コース名→コースID
 *   CourseList(A=courseName,B=courseId)を走査
 ***************************************************/
function getCourseIdByName(courseName){
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const sheet= ss.getSheetByName("CourseList");
  if(!sheet) return null;

  const lastRow= sheet.getLastRow();
  let data= sheet.getRange(2,1,lastRow-1,2).getValues();
  for(let i=0;i<data.length;i++){
    if(data[i][0]===courseName){
      return data[i][1];
    }
  }
  return null;
}


/***************************************************
 * (評価) 提出一覧取得 → Submissionsシート
 ***************************************************/
function fetchSubmissionsForSelectedAssignment(){
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet= ss.getSheetByName("Evaluation");
  const subsSheet= ss.getSheetByName("Submissions");
  if(!evalSheet||!subsSheet){
    SpreadsheetApp.getUi().alert("Evaluation or Submissionsシートがありません。");
    return;
  }

  const cName= evalSheet.getRange("A2").getValue();
  const wTitle= evalSheet.getRange("B2").getValue();
  if(!cName||!wTitle){
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未選択。");
    return;
  }

  const cId= getCourseIdByName(cName);
  if(!cId){
    SpreadsheetApp.getUi().alert(`コース名「${cName}」のIDがありません。`);
    return;
  }

  // 課題ID
  const wId= getWorkIdByCourseAndTitle(cName,wTitle);
  if(!wId){
    SpreadsheetApp.getUi().alert(`課題名「${wTitle}」の課題IDが見つかりません。`);
    return;
  }

  // Roster
  const nameMap= getRosterMap(cId);

  // Submissions
  let submissions;
  try{
    let resp= Classroom.Courses.CourseWork.StudentSubmissions.list(cId,wId);
    submissions= resp.studentSubmissions;
  }catch(e){
    SpreadsheetApp.getUi().alert("提出一覧取得失敗: "+ e);
    return;
  }

  subsSheet.clear();
  subsSheet.getRange(1,1).setValue("userId");
  subsSheet.getRange(1,2).setValue("studentName");
  subsSheet.getRange(1,3).setValue("submissionId");
  subsSheet.getRange(1,4).setValue("state");
  subsSheet.getRange(1,5).setValue("updateTime");
  subsSheet.getRange(1,6).setValue("assignedGrade");
  subsSheet.getRange(1,7).setValue("attachments");
  subsSheet.getRange(1,8).setValue("inputScore");

  if(!submissions||submissions.length===0){
    SpreadsheetApp.getUi().alert("提出がありません。");
    return;
  }

  let row=2;
  submissions.forEach(sub=>{
    const uid= sub.userId||"";
    const sid= sub.id||"";
    const st= sub.state||"";
    const ut= sub.updateTime||"";
    const ag= (sub.assignedGrade!==undefined)?sub.assignedGrade:"";
    const att= getAttachmentsStr(sub);

    subsSheet.getRange(row,1).setValue(uid);
    subsSheet.getRange(row,2).setValue(nameMap[uid]||"不明");
    subsSheet.getRange(row,3).setValue(sid);
    subsSheet.getRange(row,4).setValue(st);
    subsSheet.getRange(row,5).setValue(ut);
    subsSheet.getRange(row,6).setValue(ag);
    subsSheet.getRange(row,7).setValue(att);
    row++;
  });

  SpreadsheetApp.getUi().alert("Submissionsシートに提出一覧を表示しました。");
}


/***************************************************
 * AssignmentList(A=courseName,B=workId,C=title)
 *   → 課題IDを探す
 ***************************************************/
function getWorkIdByCourseAndTitle(courseName, workTitle){
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const sheet= ss.getSheetByName("AssignmentList");
  if(!sheet) return null;

  const lastRow= sheet.getLastRow();
  let data= sheet.getRange(2,1,lastRow-1,3).getValues();
  for(let i=0;i<data.length;i++){
    if(data[i][0]===courseName && data[i][2]===workTitle){
      return data[i][1];
    }
  }
  return null;
}


/***************************************************
 * Roster: userId -> studentName
 ***************************************************/
function getRosterMap(courseId){
  let nm={};
  let pageToken="";
  do{
    let resp= Classroom.Courses.Students.list(courseId,{pageToken});
    if(resp.students && resp.students.length){
      resp.students.forEach(s=>{
        const uid= s.userId;
        const fn= (s.profile && s.profile.name)? s.profile.name.fullName:"";
        nm[uid]= fn;
      });
    }
    pageToken= resp.nextPageToken;
  }while(pageToken);
  return nm;
}


/***************************************************
 * 添付ファイルURL を連結
 ***************************************************/
function getAttachmentsStr(sub){
  if(!sub.assignmentSubmission)return "";
  if(!sub.assignmentSubmission.attachments)return "";

  let arr=[];
  sub.assignmentSubmission.attachments.forEach(a=>{
    if(a.driveFile && a.driveFile.alternateLink){
      arr.push(a.driveFile.alternateLink);
    }else if(a.link && a.link.url){
      arr.push(a.link.url);
    }
  });
  return arr.join("; ");
}


/***************************************************
 * (評価) 得点更新
 ***************************************************/
function updateStudentGrades(){
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet= ss.getSheetByName("Evaluation");
  const subsSheet= ss.getSheetByName("Submissions");
  if(!evalSheet||!subsSheet){
    SpreadsheetApp.getUi().alert("必要なシートがありません。");
    return;
  }

  const cName= evalSheet.getRange("A2").getValue();
  const wTitle= evalSheet.getRange("B2").getValue();
  if(!cName||!wTitle){
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未選択。");
    return;
  }

  const cId= getCourseIdByName(cName);
  const wId= getWorkIdByCourseAndTitle(cName,wTitle);
  if(!cId||!wId){
    SpreadsheetApp.getUi().alert("コースID or 課題IDが見つかりません。");
    return;
  }

  const lastRow= subsSheet.getLastRow();
  if(lastRow<2){
    SpreadsheetApp.getUi().alert("Submissionsにデータがありません。");
    return;
  }

  let data= subsSheet.getRange(2,1,lastRow-1,8).getValues();
  let cnt=0;
  data.forEach(r=>{
    const subId= r[2];
    const sc= r[7];
    if(!subId) return;
    if(sc===""||isNaN(sc)) return;

    let resource={assignedGrade: sc, draftGrade: sc};
    try{
      Classroom.Courses.CourseWork.StudentSubmissions.patch(
        resource, cId, wId, subId,
        {updateMask:"assignedGrade,draftGrade"}
      );
      cnt++;
    }catch(e){
      Logger.log("採点更新失敗: "+ e);
    }
  });
  SpreadsheetApp.getUi().alert(`得点更新完了: ${cnt}件`);
}


/***************************************************
 * (評価) 得点更新 → 返却
 ***************************************************/
function updateStudentGradesAndReturn(){
  const ss= SpreadsheetApp.getActiveSpreadsheet();
  const evalSheet= ss.getSheetByName("Evaluation");
  const subsSheet= ss.getSheetByName("Submissions");
  if(!evalSheet||!subsSheet){
    SpreadsheetApp.getUi().alert("必要なシートがありません。");
    return;
  }

  const cName= evalSheet.getRange("A2").getValue();
  const wTitle= evalSheet.getRange("B2").getValue();
  if(!cName||!wTitle){
    SpreadsheetApp.getUi().alert("コース名 or 課題名が未選択。");
    return;
  }

  const cId= getCourseIdByName(cName);
  const wId= getWorkIdByCourseAndTitle(cName,wTitle);
  if(!cId||!wId){
    SpreadsheetApp.getUi().alert("コースIDまたは課題IDが見つかりません。");
    return;
  }

  const lastRow= subsSheet.getLastRow();
  if(lastRow<2){
    SpreadsheetApp.getUi().alert("Submissionsにデータがありません。");
    return;
  }

  let data= subsSheet.getRange(2,1,lastRow-1,8).getValues();
  let countUpd=0;
  let countRet=0;

  data.forEach(r=>{
    const subId= r[2];
    const sc= r[7];
    if(!subId) return;
    if(sc===""||isNaN(sc)) return;

    // 1) patch
    let resource={assignedGrade:sc, draftGrade:sc};
    try{
      Classroom.Courses.CourseWork.StudentSubmissions.patch(
        resource,cId,wId,subId,
        {updateMask:"assignedGrade,draftGrade"}
      );
      countUpd++;
    }catch(e){
      Logger.log("採点更新失敗: "+ e);
      return;
    }

    // 2) return
    try{
      Classroom.Courses.CourseWork.StudentSubmissions['return'](
        {},cId,wId,subId
      );
      countRet++;
    }catch(e){
      Logger.log("返却失敗: "+ e);
    }
  });

  SpreadsheetApp.getUi().alert(`得点更新: ${countUpd}件 / 返却: ${countRet}件 完了`);
}
