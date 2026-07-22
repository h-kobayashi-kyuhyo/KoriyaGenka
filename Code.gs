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

/**
 * 5. レシピをスプレッドシートに保存する関数
 */
function saveRecipe(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_product');
  const rSheet = ss.getSheetByName('m_product_recipe');
  
  const targetCd = data.productCd;
  
  // 1. 親データ（かき氷製品）の書き込み
  const pNextRow = getRealLastRow(pSheet, 1) + 1;
  
  // ▼ ここで D列と E列（原価と原価率）の数式を行番号に合わせて作ります
  const formulaD = `=IF(A${pNextRow}="","",SUMIFS(m_product_recipe!$E:$E,m_product_recipe!$A:$A,A${pNextRow}))`;
  const formulaE = `=IFERROR(IF(A${pNextRow}="","",D${pNextRow}/C${pNextRow}),0)`; 
  // ※もし原価率がF列の場合は、上記の formulaE を formulaF などに変え、下の配列の末尾にもう一つ "" などを挟んで位置を調整してください。
  
  // A列〜E列の 5つのデータを一気に書き込みます
  pSheet.getRange(pNextRow, 1, 1, 5).setValues([[targetCd, data.productName, data.price, formulaD, formulaE]]);
  
  // 2. 子データ（レシピ構成）の書き込み準備
  let rNextRow = getRealLastRow(rSheet, 1) + 1;
  const recipeData = [];
  
  data.ingredients.forEach(item => {
    const formulaC = `=IF(B${rNextRow}="","",IFERROR(XLOOKUP(B${rNextRow},m_source!$A:$A,m_source!$B:$B),XLOOKUP(B${rNextRow},m_material!$A:$A,m_material!$B:$B,"エラー")))`;
    const formulaE_child = `=IF(B${rNextRow}="","",D${rNextRow}*IFERROR(XLOOKUP(B${rNextRow},m_source!$A:$A,m_source!E:E),XLOOKUP(B${rNextRow},m_material!$A:$A,m_material!$I:$I,0)))`;
    
    recipeData.push([targetCd, item.cd, formulaC, item.amount, formulaE_child]);
    rNextRow++;
  });
  
  if (recipeData.length > 0) {
    const startRow = getRealLastRow(rSheet, 1) + 1;
    rSheet.getRange(startRow, 1, recipeData.length, 5).setValues(recipeData);
  }
  
  return { success: true, generatedCd: targetCd };
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