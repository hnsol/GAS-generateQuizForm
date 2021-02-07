function firstTrial() {

  // 'config'から設定値を取得;
  var config = {};
  config = initConfig('config', config);

  // 回答DBから配列を取得
  const resSS   = SpreadsheetApp.openById(config.formDstnt);
  const resShts = resSS.getSheets();
  const arrRes = [];

  let arr = [];
  resShts.forEach( sht => {
    arr = sht.getDataRange().getValues();
    if (arr.length > 1) arrRes.push(arr);
  } );



  const pdbSht = SpreadsheetApp.openById(config.idPrblmDB);
  const arrQDB = pdbSht.getSheetByName(config.quizDBSht).getDataRange().getValues();
  
}
