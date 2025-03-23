/********************************
 * saitenui.gs
 ********************************/

// 設定: シート名、開始行、各列番号、満点などの定数
const SHEET_NAME_WEBUI = 'Submissions'; // シート名
const START_ROW_WEBUI = 2;             // 2行目からデータ開始
const NAME_COLUMN_WEBUI = 2;           // B列
const LINK_COLUMN_WEBUI = 7;           // G列
const SCORE_COLUMN_WEBUI = 8;          // H列

// 画面タイトルと満点
const PAGE_TITLE_WEBUI = '1年英語ワーク(仮)';
const MAX_SCORE_WEBUI = 10;

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  template.title = PAGE_TITLE_WEBUI;
  template.maxScore = MAX_SCORE_WEBUI;
  
  // シートからデータ取得
  const submissionsData = getSubmissionsData_webui();
  template.data = submissionsData;
  
  return template.evaluate().setTitle(PAGE_TITLE_WEBUI);
}

/**
 * スプレッドシートからカードデータを取得し、必要な情報を整形して返す
 * 各行のデータオブジェクト: { rowIndex, name, link, score, embedCode, multipleLinks, linksArray }
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
    const name = row[NAME_COLUMN_WEBUI - 1];
    
    let linkValue = row[LINK_COLUMN_WEBUI - 1];
    // 未提出の場合、リンクは空文字にする
    if (!linkValue) {
      linkValue = '';
    } else {
      linkValue = linkValue.toString().trim();
    }
    
    // スコアは未入力の場合は空文字とする（「0」とは区別）
    const rawScore = row[SCORE_COLUMN_WEBUI - 1];
    const score = (rawScore === "" || rawScore === null) ? "" : rawScore;
    
    // 名前がない場合はスキップ
    if (!name) continue;
    
    let embedCode = '';
    let multipleLinks = false;
    let linksArray = [];
    
    if (linkValue === '') {
      // 未提出の場合はプレビューはなし
      embedCode = '';
    } else {
      // 複数リンクかどうか判定。 'http' の出現回数でチェック
      const countHttp = (linkValue.match(/http/g) || []).length;
      if (countHttp > 1) {
        multipleLinks = true;
        // 正規表現で全てのURLを抽出
        linksArray = linkValue.match(/http\S+/g) || [];
      } else {
        // 単一リンクの場合、プレビュー用の埋め込みコードを生成
        const res = generateEmbedCode_webui(linkValue);
        embedCode = res.embedCode;
      }
    }
    
    results.push({
      rowIndex: rowIndex,
      name: name,
      link: linkValue,
      score: score,
      embedCode: embedCode,
      multipleLinks: multipleLinks,
      linksArray: linksArray
    });
  }
  
  return results;
}

/**
 * 単一リンクからプレビュー表示用の埋め込みコードを生成する関数
 * 対応するのは、Googleスライド、Canva、Googleドキュメント、スプレッドシート、Googleマイマップ、Googleドライブ、YouTube、Vimeo、Scratch
 * @param {string} url - 対象のURL
 * @return {Object} - { embedCode: string, type: string }
 */
function generateEmbedCode_webui(url) {
  var embedCode = '';
  var type = '';
  
  if (url.indexOf('docs.google.com/presentation') !== -1) {
    embedCode = '<iframe src="' + url.replace(/\/[^/]*$/, '/embed') + '" width="100%" height="300" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'google';
  } else if (url.indexOf('canva.com') !== -1) {
    var canvaURL = url;
    var embedUrl = '';
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
    embedCode = '<div class="canva-embed-container"><iframe loading="lazy" src="' + embedUrl + '?embed" allowfullscreen="allowfullscreen" allow="fullscreen"></iframe></div>';
    type = 'canva';
  } else if (url.indexOf('docs.google.com/document') !== -1) {
    embedCode = '<iframe src="' + url.replace(/\/edit.*$/, '/preview') + '" width="100%" height="480" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'google';
  } else if (url.indexOf('docs.google.com/spreadsheets') !== -1) {
    embedCode = '<iframe src="' + url.replace(/\/edit.*$/, '/preview') + '" width="100%" height="480" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'google';
  } else if (url.indexOf('www.google.com/maps/d/') !== -1) {
    var embedUrl = url.replace('/edit', '/view');
    embedCode = '<iframe src="' + embedUrl + '" width="100%" height="480" frameborder="0" allowfullscreen="true"></iframe>';
    type = 'mymap';
  } else if (url.indexOf('drive.google.com') !== -1 && url.indexOf('file/d/') !== -1) {
    var fileIdMatch = url.match(/[-\w]{25,}/);
    if (fileIdMatch) {
      var fileId = fileIdMatch[0];
      if (url.indexOf('view') !== -1) {
        embedCode = '<iframe src="https://drive.google.com/file/d/' + fileId + '/preview" width="100%" height="360" frameborder="0" allow="autoplay; fullscreen" allowfullscreen></iframe>';
        type = 'video';
      } else {
        var imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
        embedCode = '<div style="width: 100%; max-height: 500px; display: flex; justify-content: center; align-items: center; overflow: hidden; border-radius: 8px; box-shadow: 0 2px 8px rgba(63,69,81,0.16);">' +
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
        embedCode = '<iframe src="https://drive.google.com/file/d/' + fileId + '/preview" width="100%" height="360" frameborder="0" allow="autoplay; fullscreen" allowfullscreen></iframe>';
        type = 'video';
      } else {
        var imageUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
        embedCode = '<div style="width: 100%; max-height: 500px; display: flex; justify-content: center; align-items: center; overflow: hidden; border-radius: 8px; box-shadow: 0 2px 8px rgba(63,69,81,0.16);">' +
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
      embedCode = '<iframe width="100%" height="480" src="' + videoUrl + '" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>';
      type = 'video';
    }
  } else if (url.indexOf('vimeo.com') !== -1) {
    var videoIdMatch = url.match(/vimeo\.com\/(\d+)/);
    if (videoIdMatch) {
      var videoId = videoIdMatch[1];
      var videoUrl = 'https://player.vimeo.com/video/' + videoId;
      embedCode = '<iframe src="' + videoUrl + '" width="100%" height="480" frameborder="0" allow="autoplay; fullscreen; picture-in-picture" allowfullscreen></iframe>';
      type = 'video';
    }
  } else if (url.indexOf('scratch.mit.edu/projects/') !== -1) {
    var parts = url.split('/');
    var projectId = parts[parts.length - 1];
    embedCode = '<div class="scratch-embed" data-project-id="' + projectId + '">' +
                '<iframe src="https://scratch.mit.edu/projects/' + projectId + '/embed" allowtransparency="true" width="100%" height="100%" frameborder="0" scrolling="no" allowfullscreen></iframe>' +
                '</div>';
    type = 'scratch';
  }
  
  return { embedCode: embedCode, type: type };
}

/**
 * 採点結果をスプレッドシートに書き込む
 * @param {Array} scores - [ { rowIndex: number, score: number|string }, ... ]
 */
function saveScores_webui(scores) {
  if (!scores || !scores.length) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_WEBUI);
  if (!sheet) return;
  
  // 各行ごとに書き込む（未入力の場合は空文字のまま書き込む）
  scores.forEach(function(item) {
    const row = item.rowIndex;
    const cell = sheet.getRange(row, SCORE_COLUMN_WEBUI);
    cell.setValue(item.score);
  });
  
  // 成功時のハンドリングはクライアント側で行います
}
