function main() {

  // シート全体を取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // フォームプロパティを配列に読み込む
  var form_props = ss.getSheetByName('PropertyList').getDataRange().getValues();

  // 問題リストを配列に読み込み、オブジェクト化する
  var prob_data = ss.getSheetByName('ProblemList').getDataRange().getValues();
  var prob_head = prob_data.shift();
  var problems  = new Problems(prob_head, prob_data);

  // bodyの不適切行を先に削除しておく
  // TODO:ここはまだやっていない
  cleanupBody(problems);

  // QA作成のために、ランダムにn個の数字（インデックス）を作成
  // いまのところ3個と直打ち
  //（problems.dataBodyの最大行数より小さいもの）
  var idx_rows = pickupRows(3, problems.dataBody.length);

  // QA#1-3をつくる
  problems.qa0 = generateQA(problems, idx_rows[0]);
  problems.qa1 = generateQA(problems, idx_rows[1]);
  problems.qa2 = generateQA(problems, idx_rows[2]);

  // フォーム名称を作成してから、フォームオブジェクトを生成
  var YMD  = getYYMMDD_(new Date());
  var NAME = form_props[0][1];
  var form = copyTemplateToNewForm(YMD + NAME);

  // フォームをmove
  moveForm(form);

  // フォームのプロパティを設定
  setFormProperties(form, form_props);

  // フォームにQAを追加
  addQAtoForm(form, problems.qa0);
  addQAtoForm(form, problems.qa1);
  addQAtoForm(form, problems.qa2);

}


/**
 * 問題オブジェクトのコンストラクタ
 */
function Problems(dataHead, dataBody) {
  this.dataHead = dataHead;
  this.dataBody = dataBody;
  this.idx = { title:3, corAns:6, firstChoice:7, lastChoice:10}; // HACK:直打ち
  // this.qa0 = {};
  // this.qa1 = {};
  // this.qa2 = {};
}

/**
 * まだつくっていないが、
 * problems.dataBodyを１行ずつ取り出し、OKなら残す、NGなら削除
 * 
 * @param {Object} 問題オブジェクト
 * 
 */
function cleanupBody(problems) {
  // problems.dataBodyを１行ずつ取り出し、OKなら残す、NGなら削除
  // mapとかで一気にかけないかなあ
}

/**
 * 重複のないN個のインデックス（行数）を取得する
 *
 * @param {Number}  何個のインデックスを返してほしいか
 * @param {Number}  インデックスの最大値
 * @return {Array}  インデックスを並べた配列
 * @customfunction
 * 
 * idxOfRows = [ 9, 3, 5 ]
 */
function pickupRows(numPicks, maxRows) {
  var idxOfRows = [];

  for (var i=1; i<=numPicks*2; i++) {
    idxOfRows.push(Math.floor(Math.random()*maxRows));
  } 

  idxOfRows         = uniq(idxOfRows);
  idxOfRows.length  = numPicks;

  return idxOfRows;
}

/**
 * 配列から重複を取り除く
 *
 * @param {Array}   入力配列
 * @return {Array}  入力配列から重複を取り除いた配列
 * @customfunction
 * 
 * JavaScriptのArrayでuniqする8つの方法
 * https://qiita.com/piroor/items/02885998c9f76f45bfa0
 */
function uniq(array) {
  return [...new Set(array)];
}

/**
 * Q&Aを１つ作成する
 *
 * @param {Object}  問題オブジェクト  
 * @param {Number}  使用する行
 * @return {Object} Q&Aオブジェクト
 * @customfunction
 * 
 * qa.title   = "好きな動物は？"
 * qa.corAns  = "ネコ"
 * qa.choices = [ ['イヌ', false], ['ネコ', true], ['ネズミ', false],['ヘビ', false] ]
 */
function generateQA(problems, idx_of_row) {
  var qa = {};
  qa.line     = problems.dataBody[idx_of_row]; // １行取得
  qa.title    = qa.line[problems.idx.title];    // 質問文
  qa.corAns   = qa.line[problems.idx.corAns];  // 正答
  qa.choices  = [];

  var ibg = problems.idx.firstChoice;
  var ied = problems.idx.lastChoice;

  // 配列にpush
  for (var i=ibg; i<=ied; i++) {
    var isCorrect = (qa.line[i] == qa.corAns);
    qa.choices.push([qa.line[i] , isCorrect]);
  }

  return qa;
}

/**
 * "YYMMDD_"形式の日付Stringを得る
 *
 * @param {Object}  日付オブジェクト
 * @return {String} "YYMMDD_"形式の日付String
 * @customfunction
 * 
 */
function getYYMMDD_(dt) {
  var YY  = dt.getFullYear().toString().slice(-2); // "21"
  var MM  = ("0" + (dt.getMonth()+1)).slice(-2);   // "03"
  var DD  = ("0" + (dt.getDate())).slice(-2);      // "05"
  return YY + MM + DD + "_";                       // "210305_"
}

/**
 * テンプレートをコピーし、新しいフォームを作成する
 * いくつかの項目は、GASからセットできない。
 * （たとえば成績の表示 - 送信直後）
 * テンプレートでセットしておいて、設定をコピーする必要がある。
 *
 * @param {String}  ファイル名＝フォーム名
 * @return {Object} 新しくつくったフォームオブジェクト
 * @customfunction
 * 
 */
function copyTemplateToNewForm(fileName) {
  var FILE_NAME = String(fileName);
  
  // スクリプトプロパティを登録した
  var SF_URL = PropertiesService.getScriptProperties().getProperty('SOURCE_FORM_URL');
  var source_form = FormApp.openByUrl(SF_URL);
  
  var sourceFile = DriveApp.getFileById(source_form.getId());
  var copiedFile = sourceFile.makeCopy();
  copiedFile.setName(FILE_NAME);

  var form = FormApp.openById(copiedFile.getId());
  form.setTitle(FILE_NAME);
  
  return form;
}

/**
 * フォームを特定のフォルダに移動する（FOLDER_ID）
 * コピーをFOLDER_IDに作成し、マイドライブのものを消している
 * （他に方法はないのか？）
 * FOLDER_IDはスクリプトプロパティ
 *
 * @param {Object}  フォームオブジェクト
 * @customfunction
 * 
 */
function moveForm(form) {
  var F_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  var formFile = DriveApp.getFileById(form.getId());
  DriveApp.getFolderById(F_ID).addFile(formFile);
  DriveApp.getRootFolder().removeFile(formFile);
}

/**
 * フォームのプロパティを設定する
 *
 * @param {Object}  フォームオブジェクト
 * @param {Array}   フォームプロパティの入っている配列
 * @customfunction
 * 
 */
function setFormProperties(form, form_props) {
  var FORM_DESCRIP = form_props[1][1];    //概要
  var FORM_CNF_MSG = form_props[2][1];    //終了時メッセージ
  var SHEET_ID = PropertiesService.getScriptProperties().getProperty('RESPONSE_SHEET_ID');

  form.setDescription(FORM_DESCRIP)       // 説明文
    // スクリプトプロパティ化を行った
    .setDestination(FormApp.DestinationType.SPREADSHEET, SHEET_ID)
    // 【全般タブ】
    .setCollectEmail(true)                // "メールアドレスを収集する" ON
    // 回答のコピーを送信 OFF
    .setLimitOneResponsePerUser(true)     // "回答を1回に制限する" ON
    .setAllowResponseEdits(false)         // "送信後に編集" OFF
    .setPublishingSummary(false)          // "概要グラフとテキストの回答を表示" OFF
    // 【プレゼンテーションタブ】
    .setProgressBar(false)                // "進行状況バーを表示" OFF
    .setShuffleQuestions(true)            // "質問の順序をシャッフルする" ON
    .setConfirmationMessage(FORM_CNF_MSG) // 回答後メッセージをセット
    // 【テストタブ】
    .setIsQuiz(true);                     // "テストにする"をON
    // 成績の表示 - 送信直後
    // 回答者が表示できる項目 - 不正解だった質問 ON 正解 ON 点数 ON
}

/**
 * ラジオボタン形式の質問を作成する
 *
 * @param {Object}  フォームオブジェクト
 * @param {Object}  質問文と選択肢の入っているオブジェクト
 * @customfunction
 * 
 * qa.title   = "好きな動物は？"
 * qa.choices = [ ['イヌ', false], ['ネコ', true], ['ネズミ', false],['ヘビ', false] ]
 */
function addQAtoForm(form, qa) {
  const item = form.addMultipleChoiceItem();
  item
  .setRequired(true)    // 回答を要求
  .setPoints(1)         // 1問1点固定
  .setTitle(qa.title)
  .setChoices([
    item.createChoice(qa.choices[0][0], qa.choices[0][1]), 
    item.createChoice(qa.choices[1][0], qa.choices[1][1]), 
    item.createChoice(qa.choices[2][0], qa.choices[2][1]), 
    item.createChoice(qa.choices[3][0], qa.choices[3][1]), 
    ]);                 // HACK: 直打ち、きれいな書き方を思いつけず
}

