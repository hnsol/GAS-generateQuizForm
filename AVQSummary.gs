// NOTE:まったく関数化などしておらず、
// とりあえず動くコードになっている
// 集計シートのプロトタイプを見てもらうためのもの


// 変数の名付けにあたってのローカルルール
// arr:2次元配列 arr3:3次元配列
// ss:スプレッドシート　sht:単体のシート shts:複数のシート

/**
 * 回答(DB)を集約します
 * 頻度A→シートに全部書き出し　のみ
 * 頻度B→シートに書き出し＋　人ごとに集約してメール送付
 * TODO: パラメータを渡して動作を変えられるようにしておく
 */
// function firstTrial() {
function aggregateResponse() {
  // transpose関数 // NOTE:ここがいいのか？
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));

  // 'config'から設定値を取得し、必要なものは配列化
  var config = {};
  config = initConfig('config', config);
  config.respSShHd = config.respSShHd.split(','); // HACK:配列化
  config.respSSrod = config.respSSrod.split(','); // HACK:配列化

  // TODO: 配列取得系は、関数にまとめる
  // 問題DBから配列を取得
  const shtQdb  = SpreadsheetApp.openById(config.idPrblmDB);
  const arrQdb  = shtQdb.getSheetByName(config.quizDBSht).getDataRange().getValues();
  const arrQdbT = transpose(arrQdb);

  // 回答DBから配列を取得

  // ①まずは対象シートオブジェクトをつくる
  const ssRes   = SpreadsheetApp.openById(config.formDstnt);
  const shtsRes = ssRes.getSheets();
  // 集計シートは配列取得対象から取り除いておく
  const shtsResR = shtsRes.filter( sht => sht.getName() !== config.respSShNa );

  // ②フォームタイトルを取得してオブジェクト化しておく
  const objFormTitle = {};
  shtsResR.forEach( sht => {
    const form = FormApp.openByUrl(sht.getFormUrl());
    // like フォームの回答 1: '210207_keikoチャレンジ（仮称）'
    objFormTitle[sht.getName()] = form.getTitle();
  });

  // ③回答DBから配列を取得し、このタイミングでシート名とフォームタイトルを右列に追加
  const arr3Res = [];
  let arr = [];
  shtsResR.forEach( sht => {
    arr = sht.getDataRange().getValues();
    if (arr.length > 1) {
      arr.forEach( line => {
        const sn = sht.getName();
        line.push(sn);
        line.push(objFormTitle[sn]);
      })
      arr3Res.push(arr);
    }
  } );

  // TODO: ここから配列に対する変換処理になるので、関数にまとめる
  // ③は変換処理に該当するのか？要検討

  // 回答DBの各シートに対し、各行の右側に＜問題文＞を追加
  // NOTE: 問題文がヘッダのどの列にあるかは、configで設定
  arr3Res.forEach( arr => {
    const qtx = arr[0].slice(+config.respChBgn, +config.respChEnd);
    arr.forEach( line => line.push(...qtx) );
  });

  // 回答DBの各シートに対し、各行の右側に＜正答＞を追加
  arr3Res.forEach( arr => {
    // 問題IDを配列化 // NOTE:文字列は置き換えられない？ので、arrを陽に指定
    const qid = arr[0].slice(+config.respChBgn, +config.respChEnd);
    qid.forEach( (val, idx, arr) => {
      arr[idx] = val.substring(config.respQIDBg, config.respQIDEn);
    });

    // 問題ID配列→問題DB行（row）→正答の順にマッピング
    const qca =
      qid.map( val => arrQdbT[config.pbidPbuid].indexOf(val))
        .map( row => arrQdb[row][config.pbidCorAn] )

    arr.forEach( line => line.push(...qca) );
  })

  // 回答DBの各シートに対し、各行の右側に＜マルバツ＞を追加
  // HACK:ここはめちゃめちゃ手打ち
  arr3Res.forEach( arr => {
    arr.forEach( line => {
      line.push( (line[3] == line[11])? '◯' : '✕' )
      line.push( (line[4] == line[12])? '◯' : '✕' )
      line.push( (line[5] == line[13])? '◯' : '✕' )
    })
  })
  
  // 3次元配列→2次元配列
  // 回答DBの各シートに対し、ヘッダ行を取り除く
  arr3Res.forEach( arr => arr.shift() );
 
  // 回答DBの各シートの、ボディ行を１枚のシートにくっつける
  const arrRes = []; 
  arr3Res.forEach( arr => arrRes.push(...arr) );

  // ボディ行になった配列をソート
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


  // 書き出し用配列に変換
  // 日付を修正（データを壊している可能性に注意）
  arrRes.forEach( line => {
    //console.log(Utilities.formatDate(line[0], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
    line[0] = Utilities.formatDate(line[0], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  });
  // ヘッダ行追加
  arrRes.unshift(config.respSShHd);
  const arrSmr = extractRows(arrRes, config.respSSrod);


  // メールで各人に送る
  sendShtEachAdress(arrSmr, config.respSSrod, config);  

  // サマリーシートに書き出し 
  // これは残すか？要検討
  // const shtSmr = ssRes.getSheetByName(config.respSShNa);
  // shtSmr.clear();
  // shtSmr
  //   .getRange(1, 1, arrSmr.length, arrSmr[0].length)
  //   .setNumberFormat('@')
  //   .setValues(arrSmr);


  // console.log('here');

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
 * メールで各人に送ります
 */
function sendShtEachAdress(array, rowsExt, config) {

  // メールアドレスの配列を抽出
  // （これからかく）
  // それぞれのメアドから送付対象配列を作成
  // （これからかく）

  // ダミー：あるメアドから送付対象配列を作成
  const dummyMail = 'k-watabe@avergence.co.jp';
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
