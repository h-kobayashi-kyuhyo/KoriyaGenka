/**
 * 1. ログイン処理
 */
function doGet() {
  // index.html ファイルを読み込んでWebページとして出力
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('廃棄・試食 登録システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // スマホ対応用
}


function checkLogin(inputId, inputPass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_user');
  
  if (!sheet) return { success: false, message: 'ユーザー管理シート(m_user)が見つかりません。' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === inputId && String(data[i][1]) === inputPass) {
      return { success: true, staffName: data[i][2] };
    }
  }
  return { success: false, message: 'IDまたはパスワードが間違っています。' };
}

/**
 * 2. プルダウン用のマスターデータ取得（全角「ｍ」対策＆安全版）
 */
function getIngredientMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let list = [];
  
  // 原料（m_material）の処理
  const matSheet = ss.getSheetByName('m_material');
  if (matSheet) {
    const lastRow = getRealLastRow(matSheet, 1);
    if (lastRow > 1) {
      const matData = matSheet.getRange(2, 1, lastRow - 1, 10).getValues();
      matData.forEach(row => {
        if (row[0]) {
          // row[9]（J列）が空白でない場合は「isHidden: true」という目印をつける
          list.push({ cd: row[0], name: row[1], type: '原料', isHidden: row[9] !== "" });
        }
      });
    }
  }
  
  // ソース（m_source）の処理
  const srcSheet = ss.getSheetByName('m_source');
  if (srcSheet) {
    const lastRow = getRealLastRow(srcSheet, 1);
    if (lastRow > 1) {
      const srcData = srcSheet.getRange(2, 1, lastRow - 1, 6).getValues();
      srcData.forEach(row => {
        if (row[0]) {
          // row[5]（F列）が空白でない場合は「isHidden: true」という目印をつける
          list.push({ cd: row[0], name: row[1], type: 'ソース', isHidden: row[5] !== "" });
        }
      });
    }
  }
  return list;
}
/**
 * 3. ★重要★ 本当の最終行を見つける裏技関数
 */
function getRealLastRow(sheet, column) {
  const data = sheet.getRange(1, column, sheet.getMaxRows(), 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "") {
      return i + 1;
    }
  }
  return 1;
}

/**
 * 4. 次の製品CD（自動連番）を取得する関数
 */
function getNextProductCd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_product');
  let newCd = "SE0001";
  
  if (pSheet) {
    const lastRow = getRealLastRow(pSheet, 1);
    if (lastRow > 1) {
      const cdValues = pSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let maxNum = 0;
      cdValues.forEach(row => {
        const cdStr = String(row[0]);
        const numMatch = cdStr.match(/\d+/);
        if (numMatch) {
          const num = parseInt(numMatch[0], 10);
          if (num > maxNum) maxNum = num;
        }
      });
      newCd = "SE" + ("0000" + (maxNum + 1)).slice(-4);
    }
  }
  return newCd;
}

function getProductDetail(cd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_product');
  const rSheet = ss.getSheetByName('m_product_recipe'); 
  
  let parentData = null;
  const pLastRow = getRealLastRow(pSheet, 1);
  if (pLastRow >= 2) {
    // ★F列（6列目）まで取得するように変更
    const pValues = pSheet.getRange(1, 1, pLastRow, 6).getValues();
    for (let i = 1; i < pValues.length; i++) {
      if (pValues[i][0] === cd) {
        parentData = {
          cd: pValues[i][0],
          name: pValues[i][1],       // B列: 製品名
          price: pValues[i][2],      // C列: 販売価格
          displayFlag: pValues[i][5] // ★F列(6番目): 表示区分
        };
        break;
      }
    }
  }
  
  if (!parentData) return null;
  
  const ingredients = [];
  const rLastRow = getRealLastRow(rSheet, 1);
  if (rLastRow >= 2) {
    const rValues = rSheet.getRange(1, 1, rLastRow, 4).getValues();
    for (let i = 1; i < rValues.length; i++) {
      if (rValues[i][0] === cd) {
        ingredients.push({
          cd: rValues[i][1],
          amount: rValues[i][3]
        });
      }
    }
  }
  
  parentData.ingredients = ingredients;
  return parentData;
}


/**
 * 5. レシピをスプレッドシートに保存する関数
 */
function saveRecipe(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_product');
  const rSheet = ss.getSheetByName('m_product_recipe');
  
  const targetCd = data.productCd;
  
  // 1. 親データ（m_product）の保存
  if (data.mode === 'new') {
    const pNextRow = getRealLastRow(pSheet, 1) + 1;
    
    // ★D列・E列の計算式を作成
    const formulaD = `=IF(A${pNextRow}="","",SUMIFS(m_product_recipe!$E:$E,m_product_recipe!$A:$A,A${pNextRow}))`;
    const formulaE = `=IFERROR(IF(A${pNextRow}="","",D${pNextRow}/C${pNextRow}),0)`;
    
    // ★A列〜F列（6列分）を一度に書き込み
    pSheet.getRange(pNextRow, 1, 1, 6).setValues([[
      targetCd, data.productName, data.price, formulaD, formulaE, data.displayFlag
    ]]);
    
    // ★E列（5番目）を小数第1位のパーセント表示に設定
    pSheet.getRange(pNextRow, 5).setNumberFormat("0.0%");
    
  } else {
    // 【既存編集】上書き
    const lastRow = getRealLastRow(pSheet, 1);
    const cdValues = pSheet.getRange(1, 1, lastRow, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < cdValues.length; i++) {
      if (cdValues[i][0] === targetCd) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow !== -1) {
      pSheet.getRange(targetRow, 2).setValue(data.productName);  // B列: 製品名
      pSheet.getRange(targetRow, 3).setValue(data.price);        // C列: 販売価格
      pSheet.getRange(targetRow, 6).setValue(data.displayFlag);  // ★F列: 表示区分
    }
  }

  // 2. 子データ（m_product_recipe）の保存
  if (data.mode === 'edit') {
    const rLastRow = getRealLastRow(rSheet, 1);
    if (rLastRow > 1) {
      const rData = rSheet.getRange(1, 1, rLastRow, 1).getValues();
      for (let i = rData.length - 1; i >= 1; i--) {
        if (rData[i][0] === targetCd) {
          rSheet.deleteRow(i + 1);
        }
      }
    }
  }

  let rNextRow = getRealLastRow(rSheet, 1) + 1;
  const recipeData = [];
  data.ingredients.forEach(item => {
    const formulaC = `=IF(B${rNextRow}="","",IFERROR(XLOOKUP(B${rNextRow},m_source!$A:$A,m_source!$B:$B),XLOOKUP(B${rNextRow},m_material!$A:$A,m_material!$B:$B,"エラー")))`;
    const formulaE = `=IF(B${rNextRow}="","",D${rNextRow}*IFERROR(XLOOKUP(B${rNextRow},m_source!$A:$A,m_source!E:E),XLOOKUP(B${rNextRow},m_material!$A:$A,m_material!$I:$I,0)))`;
    recipeData.push([targetCd, item.cd, formulaC, item.amount, formulaE]);
    rNextRow++;
  });
  
  if (recipeData.length > 0) {
    const startRow = getRealLastRow(rSheet, 1) + 1;
    rSheet.getRange(startRow, 1, recipeData.length, 5).setValues(recipeData);
    rSheet.getRange(startRow, 5, recipeData.length, 1).setNumberFormat("0.00");
  }
  
  return { success: true, mode: data.mode, generatedCd: targetCd };
}

function getNextMaterialCd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_material');
  let newCd = "G0001";
  
  if (sheet) {
    const lastRow = getRealLastRow(sheet, 1);
    if (lastRow > 1) {
      const cdValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let maxNum = 0;
      cdValues.forEach(row => {
        const cdStr = String(row[0]);
        const numMatch = cdStr.match(/\d+/);
        if (numMatch) {
          const num = parseInt(numMatch[0], 10);
          if (num > maxNum) maxNum = num;
        }
      });
      newCd = "G" + ("0000" + (maxNum + 1)).slice(-4);
    }
  }
  return newCd;
}
/**
 * 既存の原料データを編集のために取得する関数（単位追加版）
 */
function getMaterialDetail(cd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_material');
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return null;
  
  // J列（10列目）まで取得する
  const data = sheet.getRange(1, 1, lastRow, 10).getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === cd) {
      return {
        cd: data[i][0],
        name: data[i][1],      
        kikaku: data[i][2],    
        unit: data[i][3],      
        priceEx: data[i][4],   
        supplier: data[i][5],  
        memo: data[i][6],      
        displayFlag: data[i][9] // J列: 表示区分を追加
      };
    }
  }
  return null;
}

/**
 * 原料データを保存（新規追加 or 上書き）する関数（単位追加版）
 */
function saveMaterial(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_material');
  
  if (data.mode === 'new') {
    const nextRow = getRealLastRow(sheet, 1) + 1;
    const formulaH = `=IF(E${nextRow}="","",E${nextRow}*1.08)`;
    const formulaI = `=IF(E${nextRow}="","",E${nextRow}/C${nextRow})`;
    
    const rowData = [
      data.cd, data.name, data.kikaku, data.unit, data.priceEx, 
      data.supplier, data.memo, formulaH, formulaI, data.displayFlag
    ];
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
  } else {
    const lastRow = getRealLastRow(sheet, 1);
    const cdValues = sheet.getRange(1, 1, lastRow, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < cdValues.length; i++) {
      if (cdValues[i][0] === data.cd) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow !== -1) {
      // B〜G列を上書き
      sheet.getRange(targetRow, 2, 1, 6).setValues([[
        data.name, data.kikaku, data.unit, data.priceEx, data.supplier, data.memo
      ]]);
      // J列（10列目）の表示区分を個別に上書き
      sheet.getRange(targetRow, 10).setValue(data.displayFlag);
    }
  }
  return { success: true, mode: data.mode, generatedCd: data.cd };
}
/**
 * ===================================================
 * 【ソースマスター管理用】の追加関数
 * ===================================================
 */

/**
 * 次のソースCDを取得する関数（SO0001〜で採番）
 */
function getNextSourceCd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_source');
  let newCd = "SO0001";
  
  if (sheet) {
    const lastRow = getRealLastRow(sheet, 1);
    if (lastRow > 1) {
      const cdValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let maxNum = 0;
      cdValues.forEach(row => {
        const cdStr = String(row[0]);
        const numMatch = cdStr.match(/\d+/);
        if (numMatch) {
          const num = parseInt(numMatch[0], 10);
          if (num > maxNum) maxNum = num;
        }
      });
      newCd = "SO" + ("0000" + (maxNum + 1)).slice(-4);
    }
  }
  return newCd;
}

/**
 * 既存のソースデータ（親と子レシピ）を取得する関数
 */
function getSourceDetail(cd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_source');
  const rSheet = ss.getSheetByName('m_source_recipe');
  
  let parentData = null;
  const pLastRow = getRealLastRow(pSheet, 1);
  if (pLastRow >= 2) {
    const pValues = pSheet.getRange(1, 1, pLastRow, 6).getValues();
    for (let i = 1; i < pValues.length; i++) {
      if (pValues[i][0] === cd) {
        parentData = {
          cd: pValues[i][0],
          name: pValues[i][1],
          yieldAmount: pValues[i][2], // C列: 出来上がり量
          displayFlag: pValues[i][5]  // F列: 表示区分
        };
        break;
      }
    }
  }
  
  if (!parentData) return null;
  
  // 子シートからレシピ構成を取得
  const ingredients = [];
  const rLastRow = getRealLastRow(rSheet, 1);
  if (rLastRow >= 2) {
    const rValues = rSheet.getRange(1, 1, rLastRow, 4).getValues();
    for (let i = 1; i < rValues.length; i++) {
      if (rValues[i][0] === cd) {
        ingredients.push({
          cd: rValues[i][1],      // B列: 原料CD
          amount: rValues[i][3]   // D列: 分量
        });
      }
    }
  }
  
  parentData.ingredients = ingredients;
  return parentData;
}

/**
 * ソースデータとレシピを保存（新規追加 or 上書き）する関数
 */
function saveSource(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_source');
  const rSheet = ss.getSheetByName('m_source_recipe');
  
  const targetCd = data.sourceCd;
  
  // 1. 親データ（m_source）の保存
  if (data.mode === 'new') {
    const pNextRow = getRealLastRow(pSheet, 1) + 1;
    const formulaD = `=IF(A${pNextRow}="","",SUMIFS(m_source_recipe!$E:$E,m_source_recipe!$A:$A,A${pNextRow}))`;
    const formulaE = `=IFERROR(IF(A${pNextRow}="","",D${pNextRow}/C${pNextRow}),0)`;
    
    pSheet.getRange(pNextRow, 1, 1, 6).setValues([[
      targetCd, data.sourceName, data.yieldAmount,formulaD, formulaE, data.displayFlag
    ]]);

     pSheet.getRange(pNextRow,4).setNumberFormat("0.0");
     pSheet.getRange(pNextRow,5).setNumberFormat("0.00");
  } else {
    // 上書き（B列、C列、F列のみ更新）
    const lastRow = getRealLastRow(pSheet, 1);
    const cdValues = pSheet.getRange(1, 1, lastRow, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < cdValues.length; i++) {
      if (cdValues[i][0] === targetCd) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow !== -1) {
      pSheet.getRange(targetRow, 2).setValue(data.sourceName);
      pSheet.getRange(targetRow, 3).setValue(data.yieldAmount);
      pSheet.getRange(targetRow, 6).setValue(data.displayFlag);
    }
  }

  // 2. 子データ（m_source_recipe）の保存
  if (data.mode === 'edit') {
    // 古いレシピ行を下から順に削除（行ズレ防止のため下から）
    const rLastRow = getRealLastRow(rSheet, 1);
    if (rLastRow > 1) {
      const rData = rSheet.getRange(1, 1, rLastRow, 1).getValues();
      for (let i = rData.length - 1; i >= 1; i--) {
        if (rData[i][0] === targetCd) {
          rSheet.deleteRow(i + 1);
        }
      }
    }
  }

  // 新しいレシピの追記
  let rNextRow = getRealLastRow(rSheet, 1) + 1;
  const recipeData = [];
  data.ingredients.forEach(item => {
    const formulaC = `=IF(B${rNextRow}="","",IFERROR(XLOOKUP(B${rNextRow},m_source!$A:$A,m_source!$B:$B),XLOOKUP(B${rNextRow},m_material!$A:$A,m_material!$B:$B,"エラー")))`;
    const formulaE = `=IF(B${rNextRow}="","",D${rNextRow}*IFERROR(XLOOKUP(B${rNextRow},m_source!$A:$A,m_source!E:E),XLOOKUP(B${rNextRow},m_material!$A:$A,m_material!$I:$I,0)))`;
    
    recipeData.push([targetCd, item.cd, formulaC, item.amount, formulaE]);
    rNextRow++;
  });
  
  if (recipeData.length > 0) {
    const startRow = getRealLastRow(rSheet, 1) + 1;
    rSheet.getRange(startRow, 1, recipeData.length, 5).setValues(recipeData);

    rSheet.getRange(startRow,5,recipeData.length,1).setNumberFormat("0.00");
  }
  
  return { success: true, mode: data.mode, generatedCd: targetCd };
}

/**
 * かき氷製品の編集用リストを取得する関数
 */
function getProductList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_product');
  let list = [];
  if (sheet) {
    const lastRow = getRealLastRow(sheet, 1);
    if (lastRow >= 2) {
      // ★F列（6列目）まで読み込むように変更
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      data.forEach(row => {
        // ★row[5]（F列）が空白でない場合は「非表示」
        if (row[0]) list.push({ cd: row[0], name: row[1], isHidden: row[5] !== "" });
      });
    }
  }
  return list;
}

/**
 * ===================================================
 * 【日次業務（廃棄・試食登録）用】の追加関数
 * ===================================================
 */

/**
 * 日次業務の記録を保存する関数
 */
function saveDailyRecord(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('t_daily_record');
  
  // もしシートが無ければ自動で作成し、ヘッダーをセットする親切設計
  if (!sheet) {
    sheet = ss.insertSheet('t_daily_record');
    sheet.appendRow(['システム登録日時', '対象日付', '登録者', '区分', 'マスター種別', 'アイテムCD', 'アイテム名', '数量', '理由']);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  // タイムスタンプ（現在時刻）
  const timestamp = new Date();
  
  // データをシートの最終行に追記（アペンド）
  sheet.appendRow([
    timestamp,
    data.date,
    data.staffName,
    data.category,
    data.itemType,
    data.itemCd,
    data.itemName,
    data.quantity,
    data.reason
  ]);
  
  return { success: true };
}

/**
 * ===================================================
 * 【ログイン画面用】の追加関数
 * ===================================================
 */

/**
 * m_userシートからログインIDのリストを取得する関数
 */
function getUserList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_user');
  
  if (!sheet) return [];
  
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return [];
  
  const ids = [];
  // A列（ID）を取得
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  data.forEach(row => {
    if (row[0]) {
      ids.push(row[0]);
    }
  });
  
  return ids;
}