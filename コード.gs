function copyMasterSheetWithNamesInA2() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 現在のスプレッドシートを取得
  var masterSheet = spreadsheet.getSheetByName('master'); // 'master' シートを取得
  var settingSheet = spreadsheet.getSheetByName('setting'); // 'setting' シートを取得

  if (!masterSheet) {
    Logger.log("'master' シートが見つかりません。");
    return;
  }

  if (!settingSheet) {
    Logger.log("'setting' シートが見つかりません。");
    return;
  }

  // 'setting' シートのA列からシート名を取得
  var sheetNames = settingSheet.getRange('A1:A10').getValues().flat();

  // 各シート名で 'master' シートのコピーを作成し、A2セルにシート名を入力
  sheetNames.forEach(function(name) {
    if (name) { // 名前が空でない場合のみコピーを作成
      var newSheet = masterSheet.copyTo(spreadsheet).setName(name);
      newSheet.getRange('A2').setValue(name); // 新しいシートのA2セルにシート名を設定
    }
  });
}

function insertImageInSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  // 指定されたディレクトリID
  var directoryId = '1uK98GGWlxU_MDTSJN5sDsBUVNQ-0kTe4';
  var folder = DriveApp.getFolderById(directoryId);
  
  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    
    // ディレクトリ内のすべてのファイルを走査
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      
      // ファイル名がシート名で終わるかどうかをチェック (拡張子を含む)
      if (fileName === sheetName + '.png') {
        var fileId = file.getId();
        var imageUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
        
        // 画像をシートのA4セルに貼り付け
        var cell = sheet.getRange('A4');
        try {
          sheet.insertImage(imageUrl, cell.getColumn(), cell.getRow());
        } catch (e) {
          Logger.log(e.toString());
        }
        break; // 一致するファイルが見つかったらループを終了
      }
    }
  });
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // メニュー項目を追加
  ui.createMenu('カスタムメニュー')
    .addItem('masterシートをコピーしてシート名を追加', 'copyMasterSheetWithNamesInA2')
    .addItem('シートにGoogleDriveにある画像を挿入', 'insertImageInSheet') 
    .addToUi();
}
