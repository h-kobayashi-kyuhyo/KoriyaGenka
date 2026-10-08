/**
 * ===================================================
 * １．マスターデータ取得系（一覧・詳細・採番）
 * ===================================================
 */

// --- マスター一覧の取得（レシピ・日次用） ---
function getIngredientMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let list = [];
  
  // ソースデータの取得
  const sSheet = ss.getSheetByName('m_source');
  if (sSheet) {
    const sData = sSheet.getDataRange().getValues();
    for(let i = 1; i < sData.length; i++) {
      if(sData[i][0]) {
        // ソースはすべて「食品」扱い
        list.push({ cd: sData[i][0], name: sData[i][1], type: 'ソース', isHidden: sData[i][3] == 99, category: '食品' });
      }
    }
  }
  
  // 原料データの取得
  const mSheet = ss.getSheetByName('m_material');
  if (mSheet) {
    const mData = mSheet.getDataRange().getValues();
    for(let i = 1; i < mData.length; i++) {
      if(mData[i][0]) {
        // M列（12番目）の分類。空欄は「食品」として扱う
        const cat = mData[i][12] || '食品';
        list.push({ cd: mData[i][0], name: mData[i][1], type: '原料', isHidden: mData[i][9] == 99, category: cat });
      }
    }
  }
  return list;
}

// --- 製品一覧の取得 ---
function getProductList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_product');
  if (!sheet) return [];
  
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  return data.filter(row => row[0]).map(row => ({
    cd: row[0],
    name: row[1],
    isHidden: row[5] == 99
  }));
}

// --- 原料詳細の取得 ---
function getMaterialDetail(cd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_material');
  if (!sheet) return null;
  
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return null;
  
  const data = sheet.getRange(1, 1, lastRow, 13).getValues();
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
        displayFlag: data[i][9], 
        zaicoCd: data[i][10],     
        zaicoConv: data[i][11],
        category: data[i][12] || '食品'
      };
    }
  }
  return null;
}

// --- ソース詳細の取得 ---
function getSourceDetail(cd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSheet = ss.getSheetByName('m_source');
  const dSheet = ss.getSheetByName('m_source_detail');
  if (!sSheet || !dSheet) return null;
  
  const sLast = getRealLastRow(sSheet, 1);
  if (sLast < 2) return null;
  
  const sData = sSheet.getRange(1, 1, sLast, 5).getValues();
  let detail = null;
  for (let i = 1; i < sData.length; i++) {
    if (sData[i][0] === cd) {
      detail = { cd: cd, name: sData[i][1], yieldAmount: sData[i][2], displayFlag: sData[i][3], ingredients: [] };
      break;
    }
  }
  if (!detail) return null;
  
  const dLast = getRealLastRow(dSheet, 1);
  if (dLast >= 2) {
    const dData = dSheet.getRange(2, 1, dLast - 1, 3).getValues();
    dData.forEach(row => {
      if (row[0] === cd) detail.ingredients.push({ cd: row[1], amount: row[2] });
    });
  }
  return detail;
}

// --- 製品詳細の取得 ---
function getProductDetail(cd) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_product');
  const rSheet = ss.getSheetByName('m_recipe');
  if (!pSheet || !rSheet) return null;
  
  const pLast = getRealLastRow(pSheet, 1);
  if (pLast < 2) return null;
  
  const pData = pSheet.getRange(1, 1, pLast, 6).getValues();
  let detail = null;
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][0] === cd) {
      detail = { cd: cd, name: pData[i][1], price: pData[i][2], displayFlag: pData[i][5], ingredients: [] };
      break;
    }
  }
  if (!detail) return null;
  
  const rLast = getRealLastRow(rSheet, 1);
  if (rLast >= 2) {
    const rData = rSheet.getRange(2, 1, rLast - 1, 3).getValues();
    rData.forEach(row => {
      if (row[0] === cd) detail.ingredients.push({ cd: row[1], amount: row[2] });
    });
  }
  return detail;
}

// --- 採番処理 ---
function getNextProductCd() { return getNextCd_('m_product', 'P-'); }
function getNextSourceCd() { return getNextCd_('m_source', 'S-'); }
function getNextMaterialCd() { return getNextCd_('m_material', 'M-'); }

function getNextCd_(sheetName, prefix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return prefix + '001';
  
  const cds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  cds.forEach(row => {
    if (row[0] && String(row[0]).startsWith(prefix)) {
      const num = parseInt(String(row[0]).replace(prefix, ''), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  });
  return prefix + String(maxNum + 1).padStart(3, '0');
}


/**
 * ===================================================
 * ２．マスターデータ保存系
 * ===================================================
 */

// --- 原料の保存 ---
function saveMaterial(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_material');
  
  const zaicoConv = data.zaicoConv ? Number(data.zaicoConv) : 1;
  const safeZaicoCd = data.zaicoCd ? "'" + data.zaicoCd : "";
  const category = data.category || '食品';
  
  if (data.mode === 'new') {
    const nextRow = getRealLastRow(sheet, 1) + 1;
    const formulaH = `=IF(E${nextRow}="","",E${nextRow}*1.08)`;
    const formulaI = `=IF(E${nextRow}="","",E${nextRow}/C${nextRow})`;
    
    const rowData = [
      data.cd, data.name, data.kikaku, data.unit, data.priceEx, 
      data.supplier, data.memo, formulaH, formulaI, data.displayFlag, safeZaicoCd, zaicoConv, category
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
      sheet.getRange(targetRow, 2, 1, 6).setValues([[ data.name, data.kikaku, data.unit, data.priceEx, data.supplier, data.memo ]]);
      sheet.getRange(targetRow, 10).setValue(data.displayFlag);
      sheet.getRange(targetRow, 11).setValue(safeZaicoCd);
      sheet.getRange(targetRow, 12).setValue(zaicoConv); 
      sheet.getRange(targetRow, 13).setValue(category);
    }
  }
  return { success: true, mode: data.mode, generatedCd: data.cd };
}

// --- ソースの保存 ---
function saveSource(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSheet = ss.getSheetByName('m_source');
  const dSheet = ss.getSheetByName('m_source_detail');
  
  // 親データの保存
  if (data.mode === 'new') {
    const nextRow = getRealLastRow(sSheet, 1) + 1;
    sSheet.getRange(nextRow, 1, 1, 4).setValues([[ data.sourceCd, data.sourceName, data.yieldAmount, data.displayFlag ]]);
  } else {
    const sLast = getRealLastRow(sSheet, 1);
    const sCds = sSheet.getRange(1, 1, sLast, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < sCds.length; i++) {
      if (sCds[i][0] === data.sourceCd) { targetRow = i + 1; break; }
    }
    if (targetRow !== -1) {
      sSheet.getRange(targetRow, 2, 1, 3).setValues([[ data.sourceName, data.yieldAmount, data.displayFlag ]]);
    }
    // 古い構成を削除
    const dLast = getRealLastRow(dSheet, 1);
    if (dLast >= 2) {
      const dCds = dSheet.getRange(2, 1, dLast - 1, 1).getValues();
      for (let i = dCds.length - 1; i >= 0; i--) {
        if (dCds[i][0] === data.sourceCd) { dSheet.deleteRow(i + 2); }
      }
    }
  }
  
  // 構成（レシピ）の保存
  if (data.ingredients && data.ingredients.length > 0) {
    const nextD = getRealLastRow(dSheet, 1) + 1;
    const writeData = data.ingredients.map(ing => [data.sourceCd, ing.cd, ing.amount]);
    dSheet.getRange(nextD, 1, writeData.length, 3).setValues(writeData);
  }
  return { success: true, mode: data.mode };
}

// --- 製品の保存 ---
function saveRecipe(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('m_product');
  const rSheet = ss.getSheetByName('m_recipe');
  
  if (data.mode === 'new') {
    const nextRow = getRealLastRow(pSheet, 1) + 1;
    pSheet.getRange(nextRow, 1, 1, 6).setValues([[ data.productCd, data.productName, data.price, "", "", data.displayFlag ]]);
  } else {
    const pLast = getRealLastRow(pSheet, 1);
    const pCds = pSheet.getRange(1, 1, pLast, 1).getValues();
    let targetRow = -1;
    for (let i = 0; i < pCds.length; i++) {
      if (pCds[i][0] === data.productCd) { targetRow = i + 1; break; }
    }
    if (targetRow !== -1) {
      pSheet.getRange(targetRow, 2, 1, 2).setValues([[ data.productName, data.price ]]);
      pSheet.getRange(targetRow, 6).setValue(data.displayFlag);
    }
    const rLast = getRealLastRow(rSheet, 1);
    if (rLast >= 2) {
      const rCds = rSheet.getRange(2, 1, rLast - 1, 1).getValues();
      for (let i = rCds.length - 1; i >= 0; i--) {
        if (rCds[i][0] === data.productCd) { rSheet.deleteRow(i + 2); }
      }
    }
  }
  
  if (data.ingredients && data.ingredients.length > 0) {
    const nextR = getRealLastRow(rSheet, 1) + 1;
    const writeData = data.ingredients.map(ing => [data.productCd, ing.cd, ing.amount]);
    rSheet.getRange(nextR, 1, writeData.length, 3).setValues(writeData);
  }
  return { success: true, mode: data.mode };
}


/**
 * ===================================================
 * ３．業務記録・売上・レポート系
 * ===================================================
 */

// --- 日次記録の保存 ---
function saveDailyRecord(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('t_daily_record');
  const nextRow = getRealLastRow(sheet, 1) + 1;
  const key = Utilities.getUuid();
  const timestamp = new Date();
  
  const rowData = [
    key, timestamp, data.date, data.staffName, data.category,
    data.itemType, data.itemCd, data.itemName, data.quantity, data.reason
  ];
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
  return { success: true };
}

// --- 日次記録の一覧取得 ---
function getDailyRecordList(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('t_daily_record');
  if (!sheet) return [];
  
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const startObj = startDateStr ? new Date(startDateStr) : new Date(0);
  const endObj = endDateStr ? new Date(endDateStr) : new Date("2099-12-31");
  endObj.setHours(23, 59, 59, 999);
  
  let list = [];
  data.forEach(row => {
    if (row[0] && row[2]) {
      const rDate = new Date(row[2]);
      if (rDate >= startObj && rDate <= endObj) {
        const dStr = Utilities.formatDate(rDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
        list.push({
          key: row[0],
          label: `[${dStr}] ${row[4]} - ${row[7]} (${row[8]})`,
          dateObj: rDate
        });
      }
    }
  });
  list.sort((a, b) => b.dateObj - a.dateObj);
  return list;
}

// --- 日次記録の詳細取得 ---
function getDailyRecordDetail(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('t_daily_record');
  if (!sheet) return null;
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return null;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      return {
        date: data[i][2],
        staffName: data[i][3],
        category: data[i][4],
        itemType: data[i][5],
        itemCd: data[i][6],
        quantity: data[i][8],
        reason: data[i][9]
      };
    }
  }
  return null;
}

// --- 日次記録の更新 ---
function updateDailyRecord(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('t_daily_record');
  const lastRow = getRealLastRow(sheet, 1);
  const keys = sheet.getRange(1, 1, lastRow, 1).getValues();
  
  for (let i = 1; i < keys.length; i++) {
    if (keys[i][0] === data.key) {
      const targetRow = i + 1;
      const timestamp = new Date();
      sheet.getRange(targetRow, 2, 1, 9).setValues([[
        timestamp, data.date, data.staffName, data.category, 
        data.itemType, data.itemCd, data.itemName, data.quantity, data.reason
      ]]);
      return { success: true };
    }
  }
  return { success: false };
}

// --- 日次記録の削除 ---
function deleteDailyRecord(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('t_daily_record');
  const lastRow = getRealLastRow(sheet, 1);
  const keys = sheet.getRange(1, 1, lastRow, 1).getValues();
  
  for (let i = 1; i < keys.length; i++) {
    if (keys[i][0] === key) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// --- 売上CSVの取り込み ---
function uploadSalesCsv(csvText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('t_sales_record');
  
  if (!sheet) {
    sheet = ss.insertSheet('t_sales_record');
    sheet.appendRow(['システム取込日時', '取引日(月初)', '商品名', '製品CD', '数量', '1個あたり原価', '原価合計']);
    sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  const pSheet = ss.getSheetByName('m_product');
  let productMaster = {};
  if (pSheet) {
    const pLastRow = getRealLastRow(pSheet, 1);
    if (pLastRow >= 2) {
      const pData = pSheet.getRange(2, 1, pLastRow - 1, 4).getValues();
      pData.forEach(row => {
        if (row[0] && row[1]) {
          productMaster[row[1]] = { cd: row[0], cost: row[3] || 0 };
        }
      });
    }
  }

  try {
    const data = Utilities.parseCsv(csvText);
    const timestamp = new Date();
    const aggregated = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.length > 18) {
        const dateRaw = row[2];          
        const nameVal = row[12];         
        const qtyVal = Number(row[18]);  
        
        if (dateRaw && nameVal && !isNaN(qtyVal) && qtyVal !== 0) {
          const d = new Date(dateRaw);
          if (isNaN(d.getTime())) continue; 
          
          const yyyy = d.getFullYear();
          const mm = String(d.getMonth() + 1).padStart(2, '0');
          const firstDayStr = `${yyyy}/${mm}/01`; 
          
          const key = firstDayStr + "_" + nameVal;
          if (!aggregated[key]) {
            let cd = "";
            let cost = 0;
            if (productMaster[nameVal]) {
              cd = productMaster[nameVal].cd;
              cost = productMaster[nameVal].cost;
            }
            aggregated[key] = { month: firstDayStr, name: nameVal, cd: cd, unitCost: cost, qty: 0 };
          }
          aggregated[key].qty += qtyVal;
        }
      }
    }
    
    const rowsToAppend = [];
    for (const key in aggregated) {
      const item = aggregated[key];
      const unitCost = Number(item.unitCost.toFixed(1));
      const totalCost = Number((item.unitCost * item.qty).toFixed(1));
      rowsToAppend.push([ timestamp, item.month, item.name, item.cd, item.qty, unitCost, totalCost ]);
    }
    
    if (rowsToAppend.length > 0) {
      const nextRow = getRealLastRow(sheet, 1) + 1;
      sheet.getRange(nextRow, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
      sheet.getRange(nextRow, 6, rowsToAppend.length, 2).setNumberFormat("0.0");
      return { success: true, count: rowsToAppend.length };
    } else {
      return { success: false, message: "有効な売上データが見つかりませんでした。" };
    }
  } catch (e) {
    return { success: false, message: "CSVの解析に失敗しました。: " + e.message };
  }
}

// --- 棚卸レポートの取得（Zaico連携） ---
function getZaicoInventoryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zSheet = ss.getSheetByName('DB_zaico');
  const mSheet = ss.getSheetByName('m_material');
  
  if (!zSheet || !mSheet) return { success: false, message: 'DB_zaico または m_material シートが見つかりません。' };
  
  const mData = mSheet.getDataRange().getValues();
  const matMap = {};
  for (let i = 1; i < mData.length; i++) {
    const zCdStr = String(mData[i][10] || '').trim();
    if (zCdStr) {
      // 複数のZaicoCDに対応
      const zCds = zCdStr.split(/[,、]/).map(s => s.trim()).filter(s => s !== "");
      zCds.forEach(zCd => {
        matMap[zCd] = {
          sysCd: mData[i][0],
          sysName: mData[i][1],
          unit: mData[i][3],
          unitCost: Number(mData[i][8]) || 0,
          zaicoConv: Number(mData[i][11]) || 1
        };
      });
    }
  }
  
  const zData = zSheet.getDataRange().getValues();
  if (zData.length < 2) return { success: false, message: 'Zaicoデータがありません。' };
  
  let latestDate = 0;
  for (let i = 1; i < zData.length; i++) {
    if (zData[i][0]) {
      const d = new Date(zData[i][0]).getTime();
      if (d > latestDate) latestDate = d;
    }
  }
  if (latestDate === 0) return { success: false, message: 'Zaicoデータに有効な日付がありません。' };
  
  let totalValue = 0;
  const reportList = [];
  
  for (let i = 1; i < zData.length; i++) {
    const rowDate = new Date(zData[i][0]).getTime();
    if (rowDate === latestDate) {
      const zCd = String(zData[i][1] || '').trim();
      const zName = zData[i][2];
      const stockQty = Number(zData[i][4]) || 0;
      
      let matchStatus = false;
      let sysCd = "-";
      let sysName = zName;
      let unitCost = 0;
      let itemValue = 0;
      let displayQty = stockQty;
      
      if (zCd && matMap[zCd]) {
        matchStatus = true;
        sysCd = matMap[zCd].sysCd;
        sysName = matMap[zCd].sysName;
        unitCost = matMap[zCd].unitCost;
        
        const conv = matMap[zCd].zaicoConv;
        const sysQty = stockQty * conv; 
        itemValue = Math.round(unitCost * sysQty);
        totalValue += itemValue;
        displayQty = `${sysQty} ${matMap[zCd].unit} <br><span class="text-xs text-gray-400">(Zaico: ${stockQty} × ${conv})</span>`;
      }
      
      reportList.push({
        match: matchStatus, zaicoCd: zCd, sysCd: sysCd, name: sysName, 
        qtyHtml: displayQty, unitCost: unitCost.toFixed(1), itemValue: itemValue
      });
    }
  }
  
  const dateStr = Utilities.formatDate(new Date(latestDate), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return { success: true, targetDate: dateStr, totalValue: totalValue, list: reportList };
}


/**
 * ===================================================
 * ４．ユーザー管理・ログイン・共通ユーティリティ
 * ===================================================
 */

// --- Webアプリの初期画面表示 ---
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('店舗業務システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- 共通：実際の最終行を取得する関数 ---
function getRealLastRow(sheet, columnNumber) {
  const vals = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== "") return i + 1;
  }
  return 0;
}

// --- ユーザー一覧の取得 ---
function getUserList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_user');
  if (!sheet) return [];
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map(row => String(row[0])).filter(val => val !== "");
}

// --- ログインチェック ---
function checkLogin(id, pass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_user');
  if (!sheet) return { success: false, message: 'ユーザーマスタが存在しません。' };
  
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return { success: false, message: 'ユーザーが登録されていません。' };
  
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id) && String(data[i][2]) === String(pass)) {
      return { success: true, staffName: data[i][1] };
    }
  }
  return { success: false, message: 'IDまたはパスワードが間違っています。' };
}

// --- ユーザー詳細の取得 ---
function getUserDetail(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_user');
  if (!sheet) return null;
  const lastRow = getRealLastRow(sheet, 1);
  if (lastRow < 2) return null;
  
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      return { id: data[i][0], name: data[i][1], pass: data[i][2] };
    }
  }
  return null;
}

// --- ユーザーの保存 ---
function saveUser(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('m_user');
  if (!sheet) return { success: false };
  
  if (data.mode === 'new') {
    const nextRow = getRealLastRow(sheet, 1) + 1;
    sheet.getRange(nextRow, 1, 1, 3).setValues([[ data.id, data.name, data.pass ]]);
  } else {
    const lastRow = getRealLastRow(sheet, 1);
    const ids = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 2, 1, 2).setValues([[ data.name, data.pass ]]);
        break;
      }
    }
  }
  return { success: true };
}