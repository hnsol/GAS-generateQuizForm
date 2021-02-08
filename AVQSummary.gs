// NOTE:まったく関数化などしておらず、
// とりあえず動くコードになっている
// 集計シートのプロトタイプを見てもらうためのもの

function firstTrial() {
  // 変数の名付けにあたってのローカルルール
  // arr:2次元配列 arr3:3次元配列
  // ss:スプレッドシート　sht:単体のシート shts:複数のシート

  // transpose関数 // NOTE:ここがいいのか？
  const transpose = a => a[0].map((_, c) => a.map(r => r[c]));

  // 'config'から設定値を取得;
  var config = {};
  config = initConfig('config', config);
  config.respSShHd = config.respSShHd.split(','); // HACK:配列化
  config.respSSrod = config.respSSrod.split(','); // HACK:配列化
  // console.log(config.respSShHd);
  // console.log(config.respSSrod);

  // 問題DBから配列を取得
  const shtQdb  = SpreadsheetApp.openById(config.idPrblmDB);
  const arrQdb  = shtQdb.getSheetByName(config.quizDBSht).getDataRange().getValues();
  const arrQdbT = transpose(arrQdb);

  // 回答DBから配列を取得
  const ssRes   = SpreadsheetApp.openById(config.formDstnt);
  const shtsRes = ssRes.getSheets();
  // 集計シートは配列取得対象から取り除いておく
  const shtsResR = shtsRes.filter( sht => sht.getName() !== config.respSShNa );

  // フォームタイトルを取得してオブジェクト化しておく
  const objFormTitle = {};
  shtsResR.forEach( sht => {
    const form = FormApp.openByUrl(sht.getFormUrl());
    // like フォームの回答 1: '210207_keikoチャレンジ（仮称）'
    objFormTitle[sht.getName()] = form.getTitle();
  });

  // 回答DBから配列を取得し、このタイミングでシート名とフォームタイトルを右列に追加
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
    const qca = qid.map( val => arrQdbT[config.pbidPbuid].indexOf(val))
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


  const shtSmr = ssRes.getSheetByName(config.respSShNa);
  shtSmr.clear();
  shtSmr
    .getRange(1, 1, arrSmr.length, arrSmr[0].length)
    .setNumberFormat('@')
    .setValues(arrSmr);

  // arrSmr.forEach( line => shtSmr.appendRow(line) );


  console.log('here');

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

