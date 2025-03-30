/********************************
 * saitenui.gs
 ********************************/

// --- シート構成定義 ---
// Submissionsシートの各カラム（1列目から）
// A: courseName, B: assignmentId, C: assignmentName, D: maxPoints,
// E: userId, F: studentName, G: submissionId, H: state, I: updateTime,
// J: assignedGrade, K: attachments, L: inputScore, M: comment
const SHEET_NAME_WEBUI = 'Submissions';
const START_ROW_WEBUI = 2;

const COURSE_COLUMN = 1;             // A: courseName
const ASSIGNMENT_ID_COLUMN = 2;      // B: assignmentId (内部用)
const ASSIGNMENT_NAME_COLUMN = 3;    // C: assignmentName
const MAX_POINTS_COLUMN = 4;         // D: maxPoints
const USER_ID_COLUMN = 5;            // E: userId (内部用)
const STUDENT_NAME_COLUMN = 6;       // F: studentName
const SUBMISSION_ID_COLUMN = 7;      // G: submissionId (内部用)
const STATE_COLUMN = 8;              // H: state
const UPDATE_TIME_COLUMN = 9;        // I: updateTime
const ASSIGNED_GRADE_COLUMN = 10;    // J: assignedGrade (内部用)
const ATTACHMENTS_COLUMN = 11;       // K: attachments (ファイルリンク)
const INPUT_SCORE_COLUMN = 12;       // L: inputScore
const COMMENT_COLUMN = 13;           // M: comment

// --- 表示用定数 ---
// 採点ページのヘッダーには「Classroom一括採点アプリ」と表示する
const PAGE_TITLE_WEBUI = 'Classroom一括採点アプリ';
// 各行の満点はシートの maxPoints (D列) をそのまま利用するため、デフォルトは参考値
const DEFAULT_MAX_SCORE = 10; // ※ヘッダー上で表示しないため、内部処理のみ使用

/**
 * doGet(e)
 * URL のパスに応じて表示するテンプレートを切り替えます。
 *  - パスが指定されない場合は index.html（トップページ）を表示
 *  - パスに "saiten.html" が含まれる場合は採点用UI (saiten.html) を表示
 */
function doGet(e) {
  var templateName = "index"; // デフォルトは index.html
  if (e && e.pathInfo) {
    var path = e.pathInfo.toLowerCase();
    if (path.indexOf("saiten.html") !== -1) {
      templateName = "saiten";
    } else if (path.indexOf("kadaisakusei.html") !== -1) {
      templateName = "kadaisakusei";
    }
  }
  
  var template = HtmlService.createTemplateFromFile(templateName);
  // 採点UIの場合、ヘッダータイトルは "Classroom一括採点アプリ" となる
  template.title = PAGE_TITLE_WEBUI;
  // ヘッダー上で満点表示は不要なので maxScore は渡しません
  
  // 採点用UIの場合、Submissionsシートから取得したデータをセットする
  if (templateName === "saiten") {
    var submissionsData = getSubmissionsData_webui();
    template.data = submissionsData;
  }
  
  return template.evaluate().setTitle(PAGE_TITLE_WEBUI);
}

/**
 * getSubmissionsData_webui()
 * Submissionsシートから各行のデータを取得し、採点UI用に整形して返す。
 * 表示用名称は「課題名 (クラス名) - 生徒名」として組み立て、各行の満点はシートのD列を利用します。
 */
function getSubmissionsData_webui() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_WEBUI);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW_WEBUI) return [];
  
  const dataRange = sheet.getRange(START_ROW_WEBUI, 1, lastRow - START_ROW_WEBUI + 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  
  const results = [];
  for (let i = 0; i < values.length; i++) {
    const rowIndex = START_ROW_WEBUI + i;
    const row = values[i];
    
    const courseName = row[COURSE_COLUMN - 1];             // A: courseName
    const assignmentName = row[ASSIGNMENT_NAME_COLUMN - 1];  // C: assignmentName
    const maxPoints = row[MAX_POINTS_COLUMN - 1];            // D: maxPoints
    const studentName = row[STUDENT_NAME_COLUMN - 1];        // F: studentName
    const attachments = row[ATTACHMENTS_COLUMN - 1];         // K: attachments
    const inputScore = row[INPUT_SCORE_COLUMN - 1];          // L: inputScore
    
    // 必要な情報がなければスキップ
    if (!assignmentName || !studentName) continue;
    
    // 表示用名称: 「課題名 (クラス名) - 生徒名」
    var displayName = assignmentName;
    if (courseName) {
      displayName += " (" + courseName + ")";
    }
    displayName += " - " + studentName;
    
    // 添付ファイルのリンク取得（未提出なら空文字）
    var linkValue = "";
    if (attachments) {
      linkValue = attachments.toString().trim();
    }
    
    // 採点欄の初期値は inputScore（未入力なら空文字）
    var score = (inputScore === "" || inputScore === null) ? "" : inputScore;
    
    // プレビュー用の埋め込みコードまたはリンク一覧の生成
    let embedCode = "";
    let multipleLinks = false;
    let linksArray = [];
    
    if (linkValue === "") {
      embedCode = "";
    } else {
      const countHttp = (linkValue.match(/http/g) || []).length;
      if (countHttp > 1) {
        multipleLinks = true;
        linksArray = linkValue.match(/http\S+/g) || [];
      } else {
        const res = generateEmbedCode_webui(linkValue);
        embedCode = res.embedCode;
      }
    }
    
    results.push({
      rowIndex: rowIndex,
      displayName: displayName,
      link: linkValue,
      score: score,
      maxPoints: maxPoints,  // 各行の満点
      embedCode: embedCode,
      multipleLinks: multipleLinks,
      linksArray: linksArray
    });
  }
  
  return results;
}

/**
 * generateEmbedCode_webui(url)
 * 単一リンクから埋め込みコードを生成する関数。
 * 対応例: Googleスライド、Canva、Googleドキュメント、スプレッドシート、Googleマイマップ、Googleドライブ、YouTube、Vimeo、Scratchなど
 * @param {string} url - 対象のURL
 * @return {Object} - { embedCode: string, type: string }
 */
function generateEmbedCode_webui(url) {
  var embedCode = "";
  var type = "";
  
  if (url.indexOf('docs.google.com/presentation') !== -1) {
    embedCode = '<iframe src="' + url.replace(/\/[^/]*$/, '/embed') +
                '" width="100%" height="300" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'google';
  } else if (url.indexOf('canva.com') !== -1) {
    var canvaURL = url;
    var embedUrl = "";
    if (canvaURL.indexOf('/watch') !== -1) {
      embedUrl = canvaURL.replace('/watch', '/view');
    } else if (canvaURL.indexOf('/view') !== -1) {
      embedUrl = canvaURL;
    } else if (canvaURL.indexOf('/edit') !== -1) {
      embedUrl = canvaURL.replace('/edit', '/view')
                         .replace('utm_content=', 'view?utm_content=')
                         .split('?')[0];
    } else {
      embedUrl = canvaURL;
    }
    embedCode = '<div class="canva-embed-container"><iframe loading="lazy" src="' + embedUrl +
                '?embed" allowfullscreen="allowfullscreen" allow="fullscreen"></iframe></div>';
    type = 'canva';
  } else if (url.indexOf('docs.google.com/document') !== -1) {
    embedCode = '<iframe src="' + url.replace(/\/edit.*$/, '/preview') +
                '" width="100%" height="480" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'google';
  } else if (url.indexOf('docs.google.com/spreadsheets') !== -1) {
    embedCode = '<iframe src="' + url.replace(/\/edit.*$/, '/preview') +
                '" width="100%" height="480" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'google';
  } else if (url.indexOf('www.google.com/maps/d/') !== -1) {
    var embedUrl = url.replace('/edit', '/view');
    embedCode = '<iframe src="' + embedUrl +
                '" width="100%" height="480" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'mymap';
  } else if (url.indexOf('drive.google.com') !== -1 && url.indexOf('file/d/') !== -1) {
    var fileIdMatch = url.match(/[-\w]{25,}/);
    if (fileIdMatch) {
      var fileId = fileIdMatch[0];
      if (url.indexOf('view') !== -1) {
        embedCode = '<iframe src="https://drive.google.com/file/d/' + fileId +
                    '/preview" width="100%" height="360" frameborder="0" allow="autoplay; fullscreen" allowfullscreen></iframe>';
        type = 'video';
      } else {
        var imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
        embedCode = '<div style="width: 100%; max-height: 500px; display: flex; justify-content: center;' +
                    ' align-items: center; overflow: hidden; border-radius: 8px; box-shadow: 0 2px 8px rgba(63,69,81,0.16);">' +
                    '<img src="' + imageUrl + '" style="max-width: 100%; max-height: 100%; object-fit: contain;" alt="Google Drive Image" />' +
                    '</div>';
        type = 'image';
      }
    }
  } else if (url.indexOf('drive.google.com') !== -1 && url.indexOf('open?id=') !== -1) {
    var fileIdMatch = url.match(/[-\w]{25,}/);
    if (fileIdMatch) {
      var fileId = fileIdMatch[0];
      if (url.indexOf('view') !== -1) {
        embedCode = '<iframe src="https://drive.google.com/file/d/' + fileId +
                    '/preview" width="100%" height="360" frameborder="0" allow="autoplay; fullscreen" allowfullscreen></iframe>';
        type = 'video';
      } else {
        var imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
        embedCode = '<div style="width: 100%; max-height: 500px; display: flex; justify-content: center;' +
                    ' align-items: center; overflow: hidden; border-radius: 8px; box-shadow: 0 2px 8px rgba(63,69,81,0.16);">' +
                    '<img src="' + imageUrl + '" style="max-width: 100%; max-height: 100%; object-fit: contain;" alt="Google Drive Image" />' +
                    '</div>';
        type = 'image';
      }
    }
  } else if (url.indexOf('youtube.com') !== -1 || url.indexOf('youtu.be') !== -1) {
    var videoIdMatch = url.match(/(?:youtube\.com\/(?:[^\/\n\s]+\/\S+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^"&?\/\s]{11})/);
    if (videoIdMatch) {
      var videoId = videoIdMatch[1];
      var videoUrl = 'https://www.youtube.com/embed/' + videoId;
      embedCode = '<iframe width="100%" height="480" src="' + videoUrl +
                  '" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>';
      type = 'video';
    }
  } else if (url.indexOf('vimeo.com') !== -1) {
    var videoIdMatch = url.match(/vimeo\.com\/(\d+)/);
    if (videoIdMatch) {
      var videoId = videoIdMatch[1];
      var videoUrl = 'https://player.vimeo.com/video/' + videoId;
      embedCode = '<iframe src="' + videoUrl +
                  '" width="100%" height="480" frameborder="0" allow="autoplay; fullscreen; picture-in-picture" allowfullscreen></iframe>';
      type = 'video';
    }
  } else if (url.indexOf('scratch.mit.edu/projects/') !== -1) {
    var parts = url.split('/');
    var projectId = parts[parts.length - 1];
    embedCode = '<div class="scratch-embed" data-project-id="' + projectId + '">' +
                '<iframe src="https://scratch.mit.edu/projects/' + projectId +
                '/embed" allowtransparency="true" width="100%" height="100%" frameborder="0" scrolling="no" allowfullscreen></iframe>' +
                '</div>';
    type = 'scratch';
  }
  
  return { embedCode: embedCode, type: type };
}

/**
 * saveScores_webui(scores)
 * Submissionsシートの入力採点（L列）に、各行のスコアを書き込む
 * @param {Array} scores - [ { rowIndex: number, score: number|string }, ... ]
 */
function saveScores_webui(scores) {
  if (!scores || !scores.length) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_WEBUI);
  if (!sheet) return;
  
  scores.forEach(function(item) {
    const row = item.rowIndex;
    const cell = sheet.getRange(row, INPUT_SCORE_COLUMN);
    cell.setValue(item.score);
  });
}
