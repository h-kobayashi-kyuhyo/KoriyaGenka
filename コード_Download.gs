/**
 * Zaicoから在庫データを取得し、CSVにしてドライブに保存する
 */
function downloadZaicoAndSaveCSV() {
  const folderId = '1znMLCreuAmOAGdYna50T4N5Pp_Nv1fTF'; 
  
  // スクリプトプロパティからAPIトークンを取得
  const token = PropertiesService.getScriptProperties().getProperty('ZAICO_TOKEN');
  if (!token) {
    throw new Error('スクリプトプロパティに ZAICO_TOKEN が設定されていません。');
  }

  try {
    // --- 1. Zaico APIに「在庫データをちょうだい」と要求する ---
    const url = 'https://web.zaico.co.jp/api/v1/inventories/'; // 在庫一覧を取得するURL [cite: 1.1.6]
    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + token // 合鍵（APIトークン）を提示
      },
      'muteHttpExceptions': true
    };
    
    // データ取得を実行
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error('Zaicoからのデータ取得に失敗しました。詳細: ' + response.getContentText());
    }
    
    // 取得したデータをプログラム用の形式（JSON）に変換
    const jsonData = JSON.parse(response.getContentText()); //[cite: 1.1.6]
    
    // もしデータが空なら終了
    if (jsonData.length === 0) {
      Logger.log('在庫データが0件でした。');
      return;
    }

    // --- 2. データをCSVの形式に変換する ---
    let csvString = '';
    
    // 1行目（ヘッダー/項目名）を作成
    // ※今回は取得できたデータの項目をすべて自動でヘッダーにします
    const headers = Object.keys(jsonData[0]);
    csvString += headers.join(',') + '\n';
    
    // 2行目以降（実際の在庫データ）を作成
    for (let i = 0; i < jsonData.length; i++) {
      let row = [];
      for (let j = 0; j < headers.length; j++) {
        // カンマや改行が含まれているとおかしくなるので、ダブルクォーテーションで囲む
        let cellValue = jsonData[i][headers[j]];
        
        // 値が空（nullなど）の場合は空文字にする
        if (cellValue === null || cellValue === undefined) {
          cellValue = '';
        } else if (typeof cellValue === 'object') {
          // 追加項目（optional_attributes）などの場合は文字に変換 [cite: 1.1.6]
          cellValue = JSON.stringify(cellValue);
        }
        
        // エスケープ処理（文字の中の"を""に変換）
        cellValue = String(cellValue).replace(/"/g, '""');
        row.push('"' + cellValue + '"');
      }
      csvString += row.join(',') + '\n';
    }

    // --- 3. GoogleドライブにCSVファイルとして保存する ---
    const today = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');
    const fileName = today + '_zaico_inventory.csv';
    
    // 文字列をCSVファイル（Blob）に変換（文字化け防止のためBOMという印を付けます）
    const blob = Utilities.newBlob('\uFEFF' + csvString, 'text/csv', fileName);
    
    const folder = DriveApp.getFolderById(folderId);
    folder.createFile(blob);
    
    Logger.log('Zaicoの在庫データをCSVで保存しました: ' + fileName);

  } catch (e) {
    Logger.log('エラー: ' + e.message);
  }
}

/**
 * スプシのボタン（図形）に割り当てるための「確認用」の関数
 */
function confirmAndDownload() {
  const ui = SpreadsheetApp.getUi();
  
  // 画面上に「はい／いいえ」のポップアップを表示
  const response = ui.alert(
    '実行の確認', 
    '最新のCSVを読み込んでスプレッドシートへ転記します。よろしいですか？', 
    ui.ButtonSet.YES_NO
  );
  
  // 「はい」が押された場合のみ、メインの転記処理を実行する
  if (response === ui.Button.YES) {
    downloadZaicoAndSaveCSV(); // メインの転記関数を呼び出す
  } else {
    // 「いいえ」または「×ボタン」が押された場合
    ui.alert('処理をキャンセルしました。');
  }
}
