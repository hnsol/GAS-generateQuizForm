// 変数の名付けにあたってのローカルルール
// arr:2次元配列 arr3:3次元配列
// ss:スプレッドシート　sht:単体のシート shts:複数のシート

/**
 * 回答(DB)を集約してフィードバックに使用できるようにします
 * 頻度A→シートに全部書き出し　のみ
 * 頻度B→シートに書き出し＋　人ごとに集約してメール送付
 * TODO: パラメータを渡して動作を変えられるようにしておく
 */
// function firstTrial() {
function generateFbSheetandMail() {

  // 'config'から設定値を取得し、必要なものは配列化
  // TODO: エラーチェックも入れたいので、初期作業は関数化
  // 仮定していること……集計シートの存在、各シートの存在
  var config = {};
  config = initConfig('config', config);
  config.respSShHd = config.respSShHd.split(','); // HACK:配列化
  config.respSSrod = config.respSSrod.split(','); // HACK:配列化

  // 回答を集約し、配列にして返す
  const arrSmr = aggregateRespose(config);

  // 回答を集計シートに書き込み
  generateFbSheet(arrSmr, config);

  // 回答をメールで送付
  sendShtEachAdress(arrSmr, config);  

}

/**
 * 回答DBを集約し、配列化します
 * @param {Object} config   設定値オブジェクト
 * @return {Array} arrSmr   集約済み配列
 * NOTE: ここはfunctionに切り分けないほうが見通しがいいと思われる
 */
function aggregateRespose(config) {

  // transpose関数 // NOTE: 関数化したほうがいい気もするが、ラクなのでこういう使いかたを……
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));

  // 問題DBから配列を取得
  const shtQdb  = SpreadsheetApp.openById(config.idPrblmDB);
  const arrQdb  = shtQdb.getSheetByName(config.quizDBSht).getDataRange().getValues();
  const arrQdbT = transpose(arrQdb);

  // 【A: 回答DBから配列を取得】

  // A-1 まずは配列化対象シートのオブジェクトをつくる
  const ssRes   = SpreadsheetApp.openById(config.formDstnt);
  const shtsRes = ssRes.getSheets();
  // NOTE: 集計シートは配列取得対象から取り除いておく（R:Removed）
  const shtsResR = shtsRes.filter( sht => sht.getName() !== config.respSShNa );

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
    arr = sht.getDataRange().getValues(); // 1枚のシートを2次元配列に格納
    if (arr.length > 1) {                 // シートが空の場合は配列化しない（未回答は配列化する）
      arr.forEach( line => {
        const shtName = sht.getName();
        line.push(shtName);               // シート名を右列に追加 
        line.push(objFormTitle[shtName]); // フォームタイトルを右列に追加
      })
      arr3Res.push(arr);                  // 3次元配列に格納（シートx行x列）
    }
  });

 
  // 【B: 配列をアウトプットに向けて変換する】

  // B-1 回答DBの各シートに対し、各行の最右列に＜問題文＞を追加
  // NOTE: 問題文は必ずヘッダ行にある。どの列にあるかはconfigで指定している
  arr3Res.forEach( arr => {           // シートを取り出してarrに格納

    // ヘッダ行から問題文列を取得し配列化
    const qtx = arr[0].slice(+config.respChBgn, +config.respChEnd); 
    
    // arrの各行の最右列に＜問題文＞を追加
    arr.forEach( line => line.push(...qtx) ); 
  });

  // B-2 回答DBの各シートに対し、各行の最右列に＜正答＞を追加
  arr3Res.forEach( arr => {           // シートを取り出してarrに格納

    // ヘッダ行問題文列を取得し配列化
    const qid = arr[0].slice(+config.respChBgn, +config.respChEnd);

    // 問題文列から、問題IDを取得し配列化　ex: [No:ABCD] -> ABCD
    qid.forEach( (val, idx, arr) => {
      arr[idx] = val.substring(config.respQIDBg, config.respQIDEn);
    });

    // 問題ID配列→問題DB行（row）→正答の順にmap。どの列にあるかはconfigで指定
    const qca =
      qid.map( val => arrQdbT[config.pbidPbuid].indexOf(val))
        .map( row => arrQdb[row][config.pbidCorAn] )

    // arrの各行の最右列に＜正答＞を追加
    arr.forEach( line => line.push(...qca) );
  })

  // B-3 回答DBの各シートに対し、各行の右側に＜マルバツ＞を追加
  // TODO:ここはめちゃめちゃ手打ち、config化をすべきとは思う……
  arr3Res.forEach( arr => {
    arr.forEach( line => {
      line.push( (line[3] == line[11])? '◯' : '×' )
      line.push( (line[4] == line[12])? '◯' : '×' )
      line.push( (line[5] == line[13])? '◯' : '×' )
    })
  })
  

  // 【C: 3次元配列→2次元配列とし、アウトプットできるよう仕上げる】

  // C-1 回答DBの各シートに対し、ヘッダ行を取り除く
  arr3Res.forEach( arr => arr.shift() );
 
  // C-2 回答DBの各シートの、ボディ行を１枚のシートにくっつける
  const arrRes = []; 
  arr3Res.forEach( arr => arrRes.push(...arr) );

  // C-3 ボディ行のみになった配列をソート
  // TODO: ソート対象が手打ちなので、config化を行うこと
  let sc = new Number;
  // 配列をsc列で昇順でソート（sc: Sort Column）
  sc = 0; // 回答日付
  arrRes.sort(function(a, b){
	  if (a[sc] > b[sc]) return 1;
	  if (a[sc] < b[sc]) return -1;
	  return 0;
  });
  // 配列をsc列で降順でソート（sc: Sort Column）
  sc = 1; // メアド
  arrRes.sort(function(a, b){
	  if (a[sc] > b[sc]) return -1;
	  if (a[sc] < b[sc]) return 1;
	  return 0;
  });

  // C-4 日付を修正（破壊的変換であることに注意）、
  arrRes.forEach( line => {
    line[0] = Utilities.formatDate(line[0], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  });

  // C-5 ヘッダ行を追加（文字列はconfigで指定している）
  arrRes.unshift(config.respSShHd);

  // C-6 残す列を選択し、順番も入れ替え（configで指定している）
  const arrSmr = extractRows(arrRes, config.respSSrod);

  return arrSmr;

}


/**
 * 必要な行を抽出します
 * @param {Array} array     操作対象の2次元配列
 * @param {string} rowsExt  抽出する列 like [ '0', '8', '12', '13', '25' ]
 * @return {Array}          抽出後の2次元配列
 */
function extractRows(array, rowsExt) {
  // 行列入れ替え
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));
  var arrayT = transpose(array);

  // 抽出
  var arrayCT = [];
  rowsExt.forEach( val => arrayCT.push(arrayT[val]) );

  // 行列を入れ替えてリターン
  return transpose(arrayCT);
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
  const shtSmr = ssRes.getSheetByName(config.respSShNa);
  shtSmr.clear();
  shtSmr
    .getRange(1, 1, arrSmr.length, arrSmr[0].length)
    .setNumberFormat('@') // 文字列であることを指定
    .setValues(arrSmr);

}


/**
 * メールで各人に送ります
 * @param {Array} array     操作対象の2次元配列
 * @param {Object} config   設定値オブジェクト
 */
function sendShtEachAdress(array, config) {

  // メールアドレスの配列を抽出
  // （これからかく）
  // それぞれのメアドから送付対象配列を作成
  // （これからかく）

  // ダミー：あるメアドから送付対象配列を作成
  const dummyMail = 'm-iida@avergence.co.jp';
  const dummyMaCc = 'm-iida@avergence.co.jp';

  const arrHead = array.shift();
  const arrFltd = array.filter( line => {
    return line[3] === dummyMail;
    });

  arrFltd.unshift(arrHead);

  // 暫定措置、改行を取り除く　←　データの持たせ方を再考する必要がある
  arrFltd.forEach( (line, idRow) => {
    line.forEach( (value, idCol) => {
      // console.log(value, value.toString().replace(/\n+/g, '!'));
      arrFltd[idRow][idCol] = value.toString().replace(/\n+/g, '');
    })
  })

  // スプレッドシート（CSV）を作成
  // blobでつくるが、ドライブに置いたりはしない
  // TODO: ファイル名の文字列を作成すること（@より左、日付、SJISなど）
  const csv  = arrFltd.reduce((str, row) => str + '\n' + row);
  const blob = Utilities.newBlob('', MimeType.CSV, 'testdata_S-JIS.csv')
    .setDataFromString(csv, 'Shift-JIS');

  // メールで送付
  // TODO: いろいろベタ打ちなので直すこと
  const recipient = dummyMail;
  const subject   = 'test: sending csv...';

  let body = '';
  body += 'テストメールです\n';
  body += '添付でCSVを送ります';
  
  const options = {
    cc: dummyMaCc,
    noReply: toBoolean(config.mailOnorp),
    attachments: blob
  };

  GmailApp.sendEmail(recipient, subject, body, options);

}


/**
 * 新規シートをかしこく挿入します
 * @param {string} shtName  新規シートの名前
 * @return {Object}         作成した新規シートオブジェクト
 */
function smartInsSheet(shtName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // コピー先のシートがすでに存在する場合は、削除する
  var prevSht = ss.getSheetByName(shtName);
  if (prevSht !== null) ss.deleteSheet(prevSht);

  ss.insertSheet(shtName, ss.getNumSheets());

  return ss.getSheetByName(shtName);
}

