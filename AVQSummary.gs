// 変数の名付けにあたってのローカルルール
// arr:2次元配列 arr3:3次元配列
// ss:スプレッドシート　sht:単体のシート shts:複数のシート

/******** インタフェース ********/

/**
 * 回答(DB)を集約してシートに書き出します
 * 頻度A→シートに書き出し　のみ
 * NOTE: 実行時にパラメタを渡せないので関数を分けた
 * NOTE: 不使用が決定
 */
// function execGenerateFbSheet() {

//   // 'config'から設定値を取得し、必要なものは配列化
//   var config = setConfig();

//   // 回答を集約し、配列にして返す
//   const arrSmr = aggregateResponse(config);

//   // 回答を集計シートに書き込み
//   if (toBoolean(config.respDBwsh)) {
//     generateFbSheet(arrSmr, config);
//   };

// }

/**
 * 回答(DB)を集約してシートに書き出し、回答者毎にメールを送ります
 * 頻度B→シートに書き出し＋　人ごとに集約してメール送付
 * NOTE: 実行時にパラメタを渡せないので関数を分けた
 */
function execGenerateFbSheetandMail() {

  // 'config'から設定値を取得し、必要なものは配列化
  var config = setConfig();

  // 回答を集約し、配列にして返す
  const arrSmr = aggregateResponse(config);

  // 回答を集計シートに書き込み
  if (toBoolean(config.respDBwsh)) {
    generateFbSheet(arrSmr, config);
  };

  // 回答をメールで送付
  if (toBoolean(config.respDBsml)) {
    sendShtEachAdress(arrSmr, config);  
  }
}

/******** 主な関数 ********/

/**
 * Configを整形します
 * TODO: 初期的なエラー確認もここでやりたい
 */
function setConfig() {

  var config = fetchConfig('config');
  config.respAgHdr = config.respAgHdr.split(','); // HACK:配列化
  config.respAgRod = config.respAgRod.split(','); // HACK:配列化


  // TODO: 初期的なエラーチェックはここに入れたい
  // configシートにアクセスできなければ、エラーとみなして対処する？？

  // 回答DBにアクセスできなければ、エラーとみなして対処する

  // 集計シートにアクセスできなければ、エラーとみなして対処する

  // 問題DBにアクセスできなければ、エラーとみなして対処する

  return config;
}


/**
 * 回答DBを集約し、配列化します
 * @param {Object} config   設定値オブジェクト
 * @return {Array} arrSmr   集約済み配列
 * NOTE: ここはfunctionに切り分けないほうが見通しがいいと思われる
 */
function aggregateResponse(config) {

  // 問題DBから配列を取得
  const shtQdb  = SpreadsheetApp.openById(config.idPrblmDB);
  const arrQdb  = shtQdb.getSheetByName(config.quizDBSht).getDataRange().getValues();
  const arrQdbT = transpose2dArray(arrQdb);

  // 【A: 回答DBから配列を取得】

  // A-1 まずは配列化対象シートのオブジェクトをつくる
  const ssRes   = SpreadsheetApp.openById(config.formDstnt);
  const shtsRes = ssRes.getSheets();
  // NOTE: 集計シートは配列取得対象から取り除いておく（R:Removed）
  const shtsResR = shtsRes.filter( sht => sht.getName() !== config.respAgSna );

  // A-2 フォームタイトルを取得してオブジェクト化しておく
  // ex: objFormTitle = { フォームの回答 1: '210207_keikoチャレンジ', ... }
  const objFormTitle = {};
  shtsResR.forEach( sht => {
    const form = FormApp.openByUrl(sht.getFormUrl());
    objFormTitle[sht.getName()] = form.getTitle();
  });

  // A-3 回答DBから配列を取得し、このタイミングでシート名とフォームタイトルを右列に追加
  const arr3Res = [];
  let arr = [];
  shtsResR.forEach( sht => {
    arr = sht.getDataRange().getValues();     // 1枚のシートを2次元配列に格納
    if (arr.length > 1) {                     // シートが空の場合は配列化しない
      arr.forEach( line => {
        const shtName = sht.getName();
        line.unshift(line.length - 3);        // 設問数を左列に追加
        line.unshift(objFormTitle[shtName]);  // フォームタイトルを左列に追加
        line.unshift(shtName);                // シート名を左列に追加
      })
      arr3Res.push(arr);                      // 3次元配列に格納（シートx行x列）
    }
  });

  // console.log('A-3', arr3Res);
 

  // 【B: 配列をアウトプットに向けて変換する】

  // B-1 回答DBの各シートに対し、A1/A2/.../An/Q1/Q2.../Qnの形に変える
  var qtx = [];
  arr3Res.forEach( arr => {
    // ヘッダ行から問題文列を取得し、各行の最右列に問題文を追加
    // NOTE: 開始が[6]なのはGoogleフォームの仕様
    // NOTE: arr[0][2]に設問数を入れてあるので、+6してendを指定
    qtx = [];
    qtx = arr[0].slice(6, 6 + arr[0][2]);
    // console.log('qtx:', qtx);

    // NOTE: arr[0]には問題文が2回入っているが、後で削除するので気にしていない
    arr.forEach( line => line.push(...qtx) ); 
  });

  // console.log('B-0-1', arr3Res); // 3x2x12で狙い通り

  // B-2 Q1/A1 <RET> Q2/A1 <RET> ... Qn/Anの形に変える

  // 回答DBの各シートに対し、ヘッダ行を取り除く
  arr3Res.forEach( arr => arr.shift() );
    
  var arr3Agr = [];
  arr3Res.forEach( arr => {
    
    var arr2Agr = [];
    arr.forEach( line => {

      var lineAgr = [];
      for (var i=1; i<=line[2]; i++) {    // 設問数はline[2]にある
        lineAgr = line.slice(0,6);        // 共通情報列を取得
        lineAgr.push(line[5+i]);          // Aiを最右列に追加
        lineAgr.push(line[5+line[2]+i]);  // Qiを最右列に追加
        arr2Agr.push(lineAgr);            // 変換先配列に追加
      }
    });

    arr3Agr.push(arr2Agr);                  // 3次元配列に格納（シートx行x列）
  });

  // console.log('B-0-2', arr3Agr); // 3x3x8で狙い通り

  // B-3 回答DBの各シートに対し、各行の最右列に＜正答＞を追加
  arr3Agr.forEach( arr => {
    arr.forEach( line => {

      // 各行にある問題文を取得
      // NOTE: 設問数にかかわらず、問題文は[7]にある
      const qtext = line[7]; 

      // 問題文から問題IDを取得　ex: [No:ABCD] -> ABCD
      const qid = qtext.substring(config.respQIDBg, config.respQIDEn);

      // 問題ID→問題DB行（row）→正答。問題DBのどの列にあるかはconfigで指定
      const qrw = arrQdbT[config.pbidPbuid].indexOf(qid);
      const qca = arrQdb[qrw][config.pbidCorAn]

      // 各行の最右列に＜正答＞を追加、さらに＜マルバツ＞追加
      // NOTE: 設問数にかかわらず、回答列と正答列が定まるのでハードコーディング
      line.push(qca);
      line.push( (line[6] == line[8])? '◯' : '×' )

    });

    // console.log('arr3Agr', arr3Agr); // 3x3x10で狙い通り

  });


  // 【C: 3次元配列→2次元配列とし、アウトプットできるよう仕上げる】

  // C-1 回答DBの各シートの、ボディ行を１枚のシートにくっつける
  const arrConc = []; 
  arr3Agr.forEach( arr => arrConc.push(...arr) );

  // C-2 順番を入れ替え（configで指定している）
  const arrRes = exchangeRows(arrConc, config.respAgRod);

  // console.log('arrRes', arrRes);

  // C-3 ソート（日付→メアド）
  let sc = new Number;
  // 配列をsc列で昇順でソート（sc: Sort Column）
  sc = config.respAgRts; // 回答日付
  arrRes.sort(function(a, b){
	  if (a[sc] > b[sc]) return 1;
	  if (a[sc] < b[sc]) return -1;
	  return 0;
  });
  // 配列をsc列で降順でソート（sc: Sort Column）
  sc = config.respAgRml; // メアド
  arrRes.sort(function(a, b){
	  if (a[sc] > b[sc]) return -1;
	  if (a[sc] < b[sc]) return 1;
	  return 0;
  });

  // C-4 日付を修正 NOTE:破壊的変換であることに注意
  arrRes.forEach( line => {
    line[config.respAgRts] = 
      Utilities.formatDate(line[config.respAgRts], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  });

  // C-5 ヘッダ行を追加（文字列はconfigで指定している）
  arrRes.unshift(config.respAgHdr);

  return arrRes;

}


/**
 * 必要な行を抽出します
 * @param {Array} array     操作対象の2次元配列
 * @param {string} rowsEx   抽出する列 like [ '0', '8', '12', '13', '25' ]
 * @return {Array}          抽出後の2次元配列
 */
// function extractRows(array, rowsEx) {
function exchangeRows(array, rowsEx) {
  // 行列入れ替え
  var arrayT = transpose2dArray(array);

  // 順番の入れ替え
  var arrayCT = [];
  rowsEx.forEach( val => arrayCT.push(arrayT[val]) );

  // 行列を入れ替えてリターン
  return transpose2dArray(arrayCT);
}


/**
 * 回答を集計シートに書き込みます
 * @param {Array} arrSmr     操作対象の2次元配列
 * @param {Object} config   設定値オブジェクト
 */
function generateFbSheet(arrSmr, config) {

  // シートをIDで指定
  const ssRes   = SpreadsheetApp.openById(config.formDstnt);

  // シートに書き込み：シートが存在していることを仮定している
  // NOTE: 過去データを履歴に残しておきたいので、シート削除→新規作成は*しない*
  const shtSmr = ssRes.getSheetByName(config.respAgSna);
  shtSmr.clear();
  shtSmr
    .getRange(1, 1, arrSmr.length, arrSmr[0].length)
    .setNumberFormat('@') // 文字列であることを指定
    .setValues(arrSmr);

}

/**
 * メールで各人に送ります
 * @param {Array} array     操作対象の2次元配列（集計シートへ書き出した配列）
 * @param {Object} config   設定値オブジェクト
 * TODO: エラーハンドリングの追加
 * TODO： 送信レポートの作成
 */
function sendShtEachAdress(array, config) {

  // ヘッダは除いておく（あとで使うためここで変数に保管）
  const arrHead = array.shift();
  const YMD     = getYYMMDD_(new Date());

  // メールアドレスの配列を抽出
  const arrRcp = listupRecipient(config.mailRcpId, config.mailRcpSN , 2, config.mailRcpAp);

  // それぞれのメアドから送付対象配列を作成
  arrRcp.forEach( row => {
    
    // 必要な文字列を取得
    // '猿飛佐助 <b-ccc@ddd.co.jp>'ならば
    // fullname = '猿飛佐助', username = 'b-ccc'
    // mailaddress = 'b-ccc@ddd.co.jp'
    const reg = /(^.+)<(.+)>/;
    const fullname = reg.exec(row)[1].trim();
    const mailaddress = reg.exec(row)[2];
    const username = mailaddress.match(/(^.+)@/)[1];

    // 上記メアドに一致する行を抽出し、ヘッダを復活させる
    // const arrFltd = array.filter( line => { return line[3] === mailaddress; } );
    const arrFltd = array.filter( line => { return line[config.respAgRml] === mailaddress; } );
    arrFltd.unshift(arrHead);

    // HACK: 暫定措置、改行を取り除く　←　NOTE: データの持たせ方を再考する必要があるか？
    arrFltd.forEach( (line, idRow) => {
      line.forEach( (value, idCol) => {
        arrFltd[idRow][idCol] = value.toString().replace(/\n+/g, '');
      })
    });

    // スプレッドシート（CSV）を作成
    // blobでつくるが、ドライブに置いたりはしない
    const filename = YMD + username + '_SJIS.csv'; 
    const csv  = arrFltd.reduce((str, row) => str + '\n' + row);
    const blob = Utilities.newBlob('', MimeType.CSV, filename)
      .setDataFromString(csv, 'Shift-JIS');

    // メールで送付
    const recipient = row;
    const subject   = config.respIvMsb + '（' + username + '）';
  
    let body = '';
    body += fullname + 'さま\n\n'
    body += config.respIvMbd;
    
    const options = {
      cc: config.respIvMcc,
      noReply: toBoolean(config.mailOnorp),
      attachments: blob
    };

    // デバッグオプション：ドラフト作成までで止めることも可能
    if (toBoolean(config.respIvCdf)) {
      GmailApp.createDraft(recipient, subject, body, options);
    } else {
      GmailApp.sendEmail(recipient, subject, body, options);
    };

  });


}

