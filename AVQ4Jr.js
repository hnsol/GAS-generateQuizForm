/**
 * クイズを作成し、URLをメールで送ります
 */
function generateQuizandMail() {

  // 'config'から設定値を取得;
  var config = {};
  config = initConfig('config', config);

  // クイズを作成
  // NOTE:クイズ = フォーム + QAs = フォーム + QA1 + ... + QAn
  var formUrl = generateQuiz(config);

  // メールで通知する
  sendUrlbyMail(formUrl, config);
}


/**
 * 設定値を、設定シートから取り込みます
 * @param {string} shtName  操作対象のシートの名前
 * @param {Object} config   設定値オブジェクト
 */
function initConfig(shtName, config) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const shtConfig = ss.getSheetByName(shtName);

  return convertSht2Obj(shtConfig);
}

/**
 * シートからJSONオブジェクトを作成します
 * （1行目はヘッダ、1列目にプロパティ名、2列目にプロパティ値が入っている前提）
 * @param {Object} sheet  シートオブジェクト
 * @return {Object} obj   設定値オブジェクト
 */ 
function convertSht2Obj(sheet) {
  const array = sheet.getDataRange().getValues();
  array.shift();
  const obj = new Object();
  array.forEach( line => obj[line[0]] = line[1] );
  
  return obj;
}

/**
 * クイズフォームを作成します（フォーム + QA x n）
 * @param {Object} config           設定値オブジェクト
 * @return {string} shortenFormUrl  フォームの短縮URL
 */ 
function generateQuiz(config) {

  // テンプレートからコピーしてフォームオブジェクトを生成
  var form = copyTemplateToNewForm(config);

  // フォームのプロパティを設定
  setFormProperties(form, config);


  // 問題DBを配列に読み込み、オブジェクト化する
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pd = ss.getSheetByName(config.quizDBSht).getDataRange().getValues();
  const ph = pd.shift();
  var problems  = new Problems(ph, pd, config);

  // bodyの不適切行を削除しておく
  cleanupBody(problems);


  // QA作成のために、ランダムにn個の数字の配列を取得
  const idxRows = pickupRows(+config.nmQAItems, +problems.dataBody.length);

  // QA作成（作成する数は、idxRowsの要素数）
  idxRows.forEach( idx => {
    problems.qa = generateQA(problems, idx);
    addQAtoForm(form, problems.qa, config);
  });

  // フォーム回答ページのURL（短縮）を取得
  const publishedUrl = form.getPublishedUrl();
  const shortenFormUrl = form.shortenFormUrl(publishedUrl);

  return shortenFormUrl;
}


/**
 * テンプレートをコピーし、新しいフォームを作成します
 * @param {Object} config   設定値オブジェクト
 * @return {Object} form    新しくつくったフォームオブジェクト
 * NOTE:いくつかの項目は、GASからセットできない（たとえば成績の表示 - 送信直後）
 * そこでひながたで設定し、ひながたをコピーして設定を持ってきている。
 */
function copyTemplateToNewForm(config) {
  // フォーム名称を作成してから、フォームオブジェクトを生成
  const YMD       = getYYMMDD_(new Date());
  const NAME      = config.formTitle;
  const FILE_NAME = YMD + NAME;
  const S_FORM_ID = config.idSourceF;

  const sourceFile = DriveApp.getFileById(S_FORM_ID);
  const copiedFile = sourceFile.makeCopy();
  copiedFile.setName(FILE_NAME);

  var form = FormApp.openById(copiedFile.getId());
  form.setTitle(FILE_NAME); // ファイル名とフォームタイトルは一致させている
  
  return form;
}

/**
 * 'YYMMDD_'形式の日付Stringを得る
 * @param {Object}  dt  日付オブジェクト
 * @return {String}     'YYMMDD_'形式の日付String
 */
function getYYMMDD_(dt) {
  var YY  = dt.getFullYear().toString().slice(-2); // '21'
  var MM  = ('0' + (dt.getMonth()+1)).slice(-2);   // '03'
  var DD  = ('0' + (dt.getDate())).slice(-2);      // '05'
  return YY + MM + DD + '_';                       // '210305_'
}

/**
 * フォームのプロパティを設定します
 * @param {Object} form   フォームオブジェクト
 * @param {Object} config 設定値オブジェクト
 */
function setFormProperties(form, config) {
  const cf = config;

  // 途中コメントアウトされている行は、formオブジェクトのメソッドで設定できない項目
  // NOTE:booleanを求められる項目は、ここで変換。このやり方がベストかどうか迷っている
  form.setDescription(cf.formDscrp)                       // 説明文
    .setDestination(FormApp.DestinationType.SPREADSHEET, cf.formDstnt) // 回答記録先
    // 【全般タブ】
    .setCollectEmail(toBoolean(cf.formCMail))             // 'メールアドレスを収集する'
    // 回答のコピーを送信 OFF
    .setLimitOneResponsePerUser(toBoolean(cf.formLORPU))  // '回答を1回に制限する'
    .setAllowResponseEdits(toBoolean(cf.formAResE))       // '送信後に編集'
    .setPublishingSummary(toBoolean(cf.formPubSm))        // '概要グラフとテキストの回答を表示'
    // 【プレゼンテーションタブ】
    .setProgressBar(toBoolean(cf.formPgBar))              // '進行状況バーを表示'
    .setShuffleQuestions(toBoolean(cf.formShufQ))         // '質問の順序をシャッフルする'
    .setConfirmationMessage(cf.formCfMsg)                 // 回答後メッセージ
    // 【テストタブ】
    .setIsQuiz(toBoolean(cf.formIsQuz));                  // 'テストにする'
    // 成績の表示 - 送信直後
    // 回答者が表示できる項目 - 不正解だった質問 ON 正解 ON 点数 ON

}


/**
 * Booleanに変換
 * @param {string} string 変換する文字列
 * console.log(toBoolean('TRUE'));  // true
 * console.log(toBoolean('True'));  // true
 * console.log(toBoolean('False')); // false
 * console.log(toBoolean(123));     // false
 */
function toBoolean(string) {
  return string.toLowerCase() === 'true';
}

/**
 * 問題オブジェクトのコンストラクタ
 */
function Problems(dataHead, dataBody, config) {
  this.dataHead = dataHead;
  this.dataBody = dataBody;
  // NOTE:ここでconfigからの値を設定するのがキレイとも思えないが、他の方法を思い付いていない
  this.idx = { 
    title: config.pbidTitle,
    corAns: config.pbidCorAn,
    feedback: config.pbidFeedB,
    firstChoice: config.pbidFrstC,
    lastChoice: config.pbidLastC
  };
}

/**
 * problems.dataBodyを１行ずつ取り出し、OKなら残す、NGなら削除
 * @param {Object} problems 問題オブジェクト
 */
function cleanupBody(problems) {
  // NOTE:'+'はnumberへの変換
  const ca = +problems.idx.corAns;
  const fb = +problems.idx.feedback;

  // 正答がない場合は削除
  // NOTE:後でみたらわからなくなりそう。
  // 正答列より右側で、正答値とおなじ値をもつ列を調べ、存在すれば残している
  problems.dataBody = problems.dataBody.filter( value => {
    return value.indexOf(value[ca], ca+1) > 0;
  })

  // フィードバック欄が空白の場合、削除
  problems.dataBody = problems.dataBody.filter( value => {
    return value[fb] !== "";
  })

  // NOTE:その他のデータのエラーチェックは必要か？
}

/**
 * 重複のないN個のインデックス（行数）を取得する
 * @param {number}  numPicks 何個のインデックスを返してほしいか（N）
 * @param {number}  maxRows  インデックスの最大値
 * @return {Array}  arr      N個のインデックスを格納した配列
 * 例:idxOfRows = [ 9, 3, 5 ]　（N=3のとき）
 */
function pickupRows(numPicks, maxRows) {
  var arr = [...Array(maxRows).keys()]; // [1,2,3...,maxRows]

  var ia = arr.length; // イテレータなのでiを使って命名

  // Fisher–Yates shuffleアルゴリズム
  while (ia) {
    var ja  = Math.floor( Math.random() * ia );
    var ta  = arr[--ia];  // arrのお尻から値を取る
    arr[ia] = arr[ja];    // iaの値を、ランダム箇所の値にする 
    arr[ja] = ta;         // ランダム箇所の値をiaの値にする（この2行でスワップ）
  }

  arr.length = numPicks; // 配列の数を絞る

  return arr;
}


/**
 * Q&Aを１つ作成します
 * @param {Object}  problems  問題オブジェクト  
 * @param {Number}  idxRow    使用する行
 * @return {Object} qa        Q&Aオブジェクト
 *
 * qa.title   = '好きな動物は？'
 * qa.corAns  = 'ネコ'
 * qa.choices = [ ['イヌ', false], ['ネコ', true], ['ネズミ', false],['ヘビ', false] ]
 */
function generateQA(problems, idxRow) {
  const qa = {};
  qa.line     = problems.dataBody[idxRow];      // １行取得
  qa.title    = qa.line[problems.idx.title];    // 質問文
  qa.feedback = qa.line[problems.idx.feedback]; // フィードバック
  qa.corAns   = qa.line[problems.idx.corAns];   // 正答
  qa.choices  = [];

  // NOTE:stringをnumberに変換するため'+'付与
  var ibg = +problems.idx.firstChoice;
  var ied = +problems.idx.lastChoice;

  // 配列にpush
  for (var i=ibg; i<=ied; i++) {
    var isCorrect = (qa.line[i] == qa.corAns);
    qa.choices.push([qa.line[i] , isCorrect]);
  }

  return qa;
}

/**
 * ラジオボタン形式の質問を作成します
 * @param {Object} from   フォームオブジェクト
 * @param {Object} qa     質問文と選択肢の入っているオブジェクト
 * @param {Object} config 設定値オブジェクト
 * 
 * qa.title   = '好きな動物は？'
 * qa.choices = [ ['イヌ', false], ['ネコ', true], ['ネズミ', false],['ヘビ', false] ]
 * qa.feedback = 'よくできました！'
 */
function addQAtoForm(form, qa, config) {
  const item = form.addMultipleChoiceItem();
  item
  .setRequired(toBoolean(config.itemRqird)) // 回答の'必須'
  .setPoints(+config.itemPoint)             // 点数
  .setTitle(qa.title)                       // 質問文
  // HACK:直打ち、きれいな書き方を思いつけず      //選択肢
  .setChoices([
    item.createChoice(qa.choices[0][0], qa.choices[0][1]), 
    item.createChoice(qa.choices[1][0], qa.choices[1][1]), 
    item.createChoice(qa.choices[2][0], qa.choices[2][1]), 
    item.createChoice(qa.choices[3][0], qa.choices[3][1]), 
    ]);
  
  // NOTE:正解・不正解ともにおなじフィードバックコメントを表示させている
  item.setFeedbackForCorrect(
    FormApp.createFeedback().setText(qa.feedback).build());
  item.setFeedbackForIncorrect(
    FormApp.createFeedback().setText(qa.feedback).build());
}

/**
 * 作成されたフォームURLをメールで通知する
 * @param {string} url    フォームURL
 * @param {Object} config 設定値オブジェクト
 */
function sendUrlbyMail(url, config) {

  const recipient = config.mailRcpnt;
  const subject   = config.mailSbjct;

  let body = '';
  body += config.mailBody1 + '\n';
  body += url + '\n\n';
  body += config.mailBody2;
  
  const options = {
    name: config.mailOname,
    noReply: toBoolean(config.mailOnorp)
  }

  GmailApp.sendEmail(recipient, subject, body, options);
}


/**
 * 設定シートのボタンの動作を記述します
 * TODO: DRY原則に反しているので気持ち悪いと思っていますが、とりあえずこのまま……
 */
function buttonOnConfigSht() {

  // 開始確認（OKボタン以外は処理を中断）
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'フォーム作成の開始',
    'クイズフォームの作成を開始します。よろしいですか？',
    ui.ButtonSet.OK_CANCEL
    );
  if (response !== ui.Button.OK) return;

  // 'config'から設定値を取得;
  var config = {};
  config = initConfig('config', config);

  // クイズを作成
  // NOTE:クイズ = フォーム + QAs = フォーム + QA1 + ... + QAn
  var formUrl = generateQuiz(config);

  // 終了メッセージ
  var response = ui.alert(
    '完了しました！',
    'フォーム作成が完了しました。ご確認ください。',
    ui.ButtonSet.OK
    );

}