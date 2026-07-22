/**
 * 最新のCSVを読み込み、DB_zaicoシートへ差分（動数）を転記する
 */
function importCsvToSheet() {
  // --- 1. プロパティからIDを取得 ---
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SS_ID');
  const folderId = props.getProperty('FOLDER_ID');
  const doneFolderId = props.getProperty('DONE_FOLDER_ID');
  
  if (!ssId || !folderId) {
    throw new Error('スクリプトプロパティに SS_ID または FOLDER_ID が設定されていません。');
  }

  try {
    // --- 2. フォルダ内の「最新のCSV」を特定する ---
    const folder = DriveApp.getFolderById(folderId);
    const doneFolder = DriveApp.getFolderById(doneFolderId);
    const files = folder.getFilesByType(MimeType.CSV);
    
    let latestFile = null;
    let latestTime = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getLastUpdated().getTime() > latestTime) {
        latestTime = file.getLastUpdated().getTime();
        latestFile = file;
      }
    }
    
    if (!latestFile) {
      Logger.log('CSVファイルが見つかりません。');
      return;
    }

    // --- 3. ファイル名から日付（yyyy/mm/dd）を抽出 ---
    const fileName = latestFile.getName();
    // 連続する8桁の数字（yyyymmdd）を探す
    const dateMatch = fileName.match(/(\d{4})(\d{2})(\d{2})/);
    const recordDate = dateMatch ? `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}` : '日付不明';

    // --- 4. スプレッドシートから「前回の在庫データ」を読み込む ---
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName('DB_zaico');
    const lastRow = sheet.getLastRow();
    
    // 商品CDをキーにして前回の在庫数を記憶する箱
    let previousStockMap = {};
    
    if (lastRow > 1) {
      // B列(2):商品CD, E列(5):今回在庫数 を読み込む
      const existingData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const code = existingData[i][1];
        const stock = existingData[i][4];
        // 毎日実行しなくても、シートにある「一番下（最後）の在庫数」が前回値として残る
        if (code) {
          previousStockMap[code] = Number(stock);
        }
      }
    }

    // --- 5. CSVを解析して追記データを作成 ---
    // CSVの文字化け（BOM）を取り除いて解析
    const csvString = latestFile.getBlob().getDataAsString().replace(/^\uFEFF/, '');
    const csvData = Utilities.parseCsv(csvString);
    
    if (csvData.length < 2) {
      Logger.log('CSVにデータ行がありません。');
      return;
    }

    // ヘッダー行から各項目の列番号（インデックス）を取得
    const headers = csvData[0];
    const idIdx = headers.indexOf('id');
    const titleIdx = headers.indexOf('title');
    const qtyIdx = headers.indexOf('quantity');
    
    if (idIdx === -1 || titleIdx === -1 || qtyIdx === -1) {
      throw new Error('CSV内に id, title, quantity のいずれかの列が見つかりません。');
    }

    let insertData = [];

    // 2行目（データ行）からループ処理
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (row.length === 1 && row[0] === "") continue; // 空行はスキップ
      
      const itemId = row[idIdx];
      const itemTitle = row[titleIdx];
      const currentQty = Number(row[qtyIdx]) || 0;

      // 前回の在庫を取得（シートに存在しない新規商品は、今日の在庫と同じとして動数0にする）
      const prevQty = previousStockMap[itemId] !== undefined ? previousStockMap[itemId] : currentQty;
      
      // 動数（今回 - 前回）の計算
      const movement = currentQty - prevQty;

      // 配列に追加（A列:日付, B列:CD, C列:名, D列:動数, E列:今回数）
      insertData.push([
        recordDate,
        itemId,
        itemTitle,
        movement,
        currentQty
      ]);
    }

    // --- 6. スプレッドシートへ一括書き込み ---
    if (insertData.length > 0) {
      sheet.getRange(lastRow + 1, 1, insertData.length, 5).setValues(insertData);
      Logger.log(`スプレッドシートへの転記が完了しました（${fileName}）`);
     
      // --- 7. 【追加】処理が終わったCSVを「済フォルダ」へ移動 ---
      latestFile.moveTo(doneFolder);
      Logger.log('処理済みのCSVを移動しました。');
    }

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
  }
}
function confirmAndRun() {
  const ui = SpreadsheetApp.getUi();
  
  // 画面上に「はい／いいえ」のポップアップを表示
  const response = ui.alert(
    '実行の確認', 
    '最新のCSVを読み込んでスプレッドシートへ転記します。よろしいですか？', 
    ui.ButtonSet.YES_NO
  );
  
  // 「はい」が押された場合のみ、メインの転記処理を実行する
  if (response === ui.Button.YES) {
    importCsvToSheet(); // メインの転記関数を呼び出す
  } else {
    // 「いいえ」または「×ボタン」が押された場合
    ui.alert('処理をキャンセルしました。');
  }
}