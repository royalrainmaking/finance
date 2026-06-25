const PLAN_CONFIG = {
  '1.ฝนหลวง':     { id: '1E43wKRlEa1xKM5l_zSDdstmAKPtDgw05rwJVeRLgTYk', sheet: '1.ฝน' },
  '2.ด้านการบิน':   { id: '1lwDV6QKiJ0nx4guP6T8BhIMquCmwIqzn-zdVnjYa_hc', sheet: '2.บิน' },
  '2.1 บินสาธาฯ': { id: '1lwDV6QKiJ0nx4guP6T8BhIMquCmwIqzn-zdVnjYa_hc', sheet: '2.1 บินสาธาฯ' },
  '3.แก้ปัญหาฝุ่น':   { id: '1iOX9ISvptU2UpMbVNPTuYHbA4df11dytCHxFYKoSQxg', sheet: '3.ฝุ่น' },
  '4.บรรเทาลูกเห็บ': { id: '1SM0GYznjAKPOQT-itypsk83n7l5rx9ibWzKcLGZb4H4', sheet: '4.ลูกเห็บ' }
};

function doGet(e) {
  try {
    const action = (e && e.parameter) ? e.parameter.action : null;
    
    if (action === 'gf') {
      const template = HtmlService.createTemplateFromFile('GF_Index');
      template.initialPlan = e.parameter.plan || '';
      return template.evaluate()
          .setTitle('เครื่องมืออัพเดทข้อมูลจาก Excel ไปยัง Google Sheet')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }
    
    if (!action) {
      return HtmlService.createHtmlOutputFromFile('index')
          .setTitle('Budget Flow PRO')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }

    const planKey = e.parameter.planKey || '1.ฝนหลวง'; 
    if (action === 'getInitialData') return ContentService.createTextOutput(JSON.stringify(budget_getInitialData(planKey))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'getBuranakanData') return ContentService.createTextOutput(JSON.stringify(budget_getBuranakanData(planKey))).setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid Action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) { return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON); }
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const planKey = params.planKey || '1.ฝนหลวง';
    let res;
    if (action === 'submitReserve') res = budget_submitReserve(params.data, planKey);
    else if (action === 'submitReserveAdd') res = budget_submitReserveAdd(params.data, planKey);
    else if (action === 'submitDeduct') res = budget_submitDeduct(params.data, planKey);
    else if (action === 'submitOffset') res = budget_submitOffset(params.data, planKey);
    else if (action === 'submitDeductAdd') res = budget_submitDeductAdd(params.data, planKey);
    else if (action === 'submitUpdate') res = budget_submitUpdate(params.data, planKey);
    else if (action === 'submitCancel') res = budget_submitCancel(params.data, planKey);
    else if (action === 'submitReserveAndDeduct') res = budget_submitReserveAndDeduct(params.data, planKey);
    else if (action === 'submitUpdateDD') res = budget_submitUpdateDD(params.data, planKey);
    else if (action === 'gf_getSheetData') res = { result: getSheetData(params.data.spreadsheetId, params.data.sheetName) };
    else if (action === 'gf_updateSheetData') res = { result: updateSheetData(params.data.updates, params.data.config) };
    else if (action === 'gf_searchAmounts') res = { result: searchAmounts(params.data.excelDataStr, params.data.configStr) };
    else throw new Error("Invalid Action");
    
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) { return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON); }
}

function budget_getTargetSheet(planKey) { 
  const config = PLAN_CONFIG[planKey] || PLAN_CONFIG['1.ฝนหลวง'];
  const ss = SpreadsheetApp.openById(config.id);
  const sheet = ss.getSheetByName(config.sheet);
  if (!sheet) throw new Error("ไม่พบชีทชื่อ '" + config.sheet + "' ในระบบแผน " + planKey);
  return sheet; 
}

function budget_getBuranakanData(planKey) {
  let targetGid = 783233359; // Default
  if (planKey === '1.ฝนหลวง') targetGid = 783233359;
  if (planKey === '2.ด้านการบิน' || planKey === '2.1 บินสาธาฯ') targetGid = 2072042102;
  if (planKey === '3.แก้ปัญหาฝุ่น') targetGid = 577836523;
  if (planKey === '4.บรรเทาลูกเห็บ') targetGid = 783233359;
  let targetSheet = null;

  // Get the specific spreadsheet ID for the requested plan
  const planConfig = PLAN_CONFIG[planKey] || PLAN_CONFIG['1.ฝนหลวง'];
  
  try {
    const ss = SpreadsheetApp.openById(planConfig.id);
    const sheets = ss.getSheets();
    for (const sheet of sheets) {
      if (sheet.getSheetId() === targetGid) {
        targetSheet = sheet;
        break;
      }
    }
  } catch (e) {
    // Ignore error
  }

  if (!targetSheet) {
    return { error: 'ไม่พบชีทที่มี gid=' + targetGid + ' ในไฟล์ที่เชื่อมต่ออยู่' };
  }

  const parseSheetNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    const clean = String(val).replace(/,/g, '').trim();
    const num = Number(clean);
    return isNaN(num) ? 0 : num;
  };

  const rowNum = (planKey === '2.1 บินสาธาฯ') ? 51 : (planKey === '2.ด้านการบิน') ? 17 : 16;
  const gVal = parseSheetNum(targetSheet.getRange('G' + rowNum).getValue());
  const nVal = parseSheetNum(targetSheet.getRange('N' + rowNum).getValue());
  const lVal = parseSheetNum(targetSheet.getRange('L' + rowNum).getValue());
  
  return {
    success: true,
    data: {
      budgetAfterAdjust: gVal,
      gfTotal: nVal,
      contractRemaining: lVal,
      budgetRemaining: gVal - nVal - lVal
    }
  };
}

function budget_formatThaiDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date)) return date;
  const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  let year = date.getFullYear();
  if (year < 2400) year += 543;
  return `${date.getDate()} ${months[date.getMonth()]} ${year.toString().slice(-2)}`;
}

function budget_applyFormulas(sheet, row) {
  sheet.getRange(row, 12).setFormula(`=IF(E${row}="","",TEXT(E${row},"mmm"))`);
  const tripletCols = [16,19,22,25,28,31,34,37,40,43,46,49,52,55,58,61,64,67,70,73,76,79,82,85,88,91,94,97,100,103];
  tripletCols.forEach(c => {
    const colName = sheet.getRange(row, c-1).getA1Notation().replace(/\d+/, '');
    sheet.getRange(row, c).setFormula(`=IF($E${row}<>"",${colName}${row},"")`);
  });
  if (sheet.getName() === '2.1 บินสาธาฯ') {
    sheet.getRange(row, 104).setFormula(`=AI${row}`); // CZ
    sheet.getRange(row, 105).setFormula(`=AJ${row}`); // DA
    sheet.getRange(row, 106).setFormula(`=AK${row}`); // DB
    sheet.getRange(row, 107).setFormula(`=AL${row}`); // DC
  } else {
    const reserveRange = ["N","Q","T","W","Z","AC","AF","AI","AL","AO","AR","AU","AX","BA","BD","BG","BJ","BM","BP","BS","BV","BY","CB","CE","CH","CK","CN","CQ","CT","CW"];
    const deductRange  = ["O","R","U","X","AA","AD","AG","AJ","AM","AP","AS","AV","AY","BB","BE","BH","BK","BN","BQ","BT","BW","BZ","CC","CF","CI","CL","CO","CR","CU","CX"];
    const balanceRange = ["P","S","V","Y","AB","AE","AH","AK","AN","AQ","AT","AW","AZ","BC","BF","BI","BL","BO","BR","BU","BX","CA","CD","CG","CJ","CM","CP","CS","CV","CY"];
    sheet.getRange(row, 104).setFormula("=" + reserveRange.map(c => `${c}${row}`).join("+")); // CZ
    sheet.getRange(row, 105).setFormula("=" + deductRange.map(c => `${c}${row}`).join("+"));  // DA
    sheet.getRange(row, 106).setFormula("=" + balanceRange.map(c => `${c}${row}`).join("+")); // DB
    sheet.getRange(row, 107).setFormula(`=M${row}`); // DC
  }
}

function budget_getInitialData(planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 5) return { entries: [], todayThai: budget_formatThaiDate(new Date()) };
  const headers = sheet.getRange(2, 1, 2, lastCol).getValues();
  const data = sheet.getRange(5, 1, lastRow - 4, lastCol).getValues();
  let lastParentInfo = {};
  const entries = data.filter(r => r[0]).map(row => {
    let catCode = "", catName = "", reserveAmount = 0, deductAmount = 0, col = 0;
    for (let c = 13; c < lastCol; c += 3) {
      if (headers[0][c]) {
        const v1 = row[c];
        const v2 = row[c+1];
        const hasV1 = (v1 !== "" && v1 !== null && !isNaN(parseFloat(v1)));
        const hasV2 = (v2 !== "" && v2 !== null && !isNaN(parseFloat(v2)));
        if (hasV1 || hasV2) {
          catCode = headers[0][c] || "-";
          catName = headers[1][c] || "";
          reserveAmount = hasV1 ? parseFloat(v1) : "";
          deductAmount = hasV2 ? parseFloat(v2) : "";
          col = c + 1;
          break;
        }
      }
    }
    
    let liqRef = row[3];
    if (liqRef instanceof Date && liqRef.toISOString() === '2025-10-28T17:00:00.000Z') liqRef = '69-05-00033';
    else if (liqRef === '2025-10-28T17:00:00.000Z') liqRef = '69-05-00033';

    const id = row[0].toString();
    const isSub = id.includes(".");
    
    // Core data
    let name = row[8];
    let dept = row[9];
    let desc = row[12];
    let cat = row[10] || catCode;
    let type = (row[1] || "").toString().trim();

    // Inheritance logic: If sub-item is missing core info, pull from last parent
    if (!isSub) {
      lastParentInfo = { name, dept, desc, cat, type };
    } else {
      if (!name) name = lastParentInfo.name;
      if (!dept) dept = lastParentInfo.dept;
      if (!desc) desc = lastParentInfo.desc;
      if (!cat) cat = lastParentInfo.cat;
      if (!type) type = lastParentInfo.type;
    }

    return { 
      id: id, 
      type: type,
      date: budget_formatThaiDate(row[2]), 
      letterDateRaw: row[7], letterDateFormatted: budget_formatThaiDate(row[7]), 
      refNo: row[6], name: name, dept: dept, desc: desc, 
      catCode: cat, catName, amount: reserveAmount, amountDeduct: deductAmount, 
      col, 
      colF: budget_formatThaiDate(row[5]), 
      colE: budget_formatThaiDate(row[4]), 
      liquidateRefNo: liqRef,
      colDD: row.length > ((planKey === '2.1 บินสาธาฯ') ? 38 : 107) ? row[((planKey === '2.1 บินสาธาฯ') ? 38 : 107)] : ""
    };
  }).reverse();
  return { entries, todayThai: budget_formatThaiDate(new Date()) };
}

function budget_submitReserve(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  let max = 0; ids.forEach(r => { let v = parseInt(r[0]); if(v > max) max = v; });
  const nextId = max + 1;
  const nr = lastRow + 1;
  sheet.getRange(nr, 1).setValue(nextId);
  sheet.getRange(nr, 2).setValue(data.type === 'PO' ? 'PO' : (data.remark || ""));
  sheet.getRange(nr, 3).setValue(new Date()); 
  sheet.getRange(nr, 7).setValue(data.refNo);
  sheet.getRange(nr, 8).setValue(data.letterDate);
  sheet.getRange(nr, 9).setValue(data.name);
  sheet.getRange(nr, 10).setValue(data.dept);
  sheet.getRange(nr, 11).setValue(data.catCode);
  sheet.getRange(nr, 13).setValue(data.desc);
  if (data.col) sheet.getRange(nr, parseInt(data.col)).setValue(parseFloat(data.amount));
  budget_applyFormulas(sheet, nr);
  return { success: true, id: nextId };
}

function budget_submitDeduct(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const idData = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  
  if (data.mode === 'PO') {
    let pIdx = -1; let mSub = 0; let insAt = -1;
    const pIdStr = data.id.toString();
    for (let i = 0; i < idData.length; i++) {
        const cur = idData[i][0].toString();
        if (cur === pIdStr) { pIdx = i + 5; insAt = pIdx; }
        if (cur.startsWith(pIdStr + ".")) {
            const s = parseInt(cur.split(".")[1]); if (s > mSub) mSub = s;
            insAt = i + 5;
        }
    }
    if (pIdx === -1) throw new Error("Parent ID not found");
    
    // Check lock value on parent
    const parentLiq = sheet.getRange(pIdx, 4).getValue();
    const parentLiqStr = (parentLiq instanceof Date) ? parentLiq.toISOString() : (parentLiq ? parentLiq.toString() : "");
    if (parentLiqStr === '2025-10-28T17:00:00.000Z' || parentLiqStr === '69-05-00033') throw new Error("รายการหลักถูกตัดยอดแล้ว (Locked: 69-05-00033)");

    const nId = pIdStr + "." + (mSub + 1);
    sheet.insertRowAfter(insAt);
    const nr = insAt + 1;
    sheet.getRange(nr, 1).setValue(nId);
    sheet.getRange(nr, 2).setValue("PO");
    if (data.colF) sheet.getRange(nr, 6).setValue(data.colF);
    sheet.getRange(nr, 4).setValue(data.liquidateRefNo);
    [3, 7, 9, 10, 11, 12, 13].forEach(c => sheet.getRange(nr, c).setValue(sheet.getRange(pIdx, c).getValue()));
    sheet.getRange(nr, 8).setValue(data.liquidateLetterDate); // Set new date in Column H
    sheet.getRange(nr, 3).setValue(new Date()); 
    if (data.name) sheet.getRange(nr, 9).setValue(data.name);
    if (data.desc) sheet.getRange(nr, 13).setValue(data.desc);
    const pVals = sheet.getRange(pIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (let c = 13; c < pVals.length; c++) {
      if (sheet.getRange(2, c+1).getValue() && pVals[c] !== "" && typeof pVals[c] === 'number') {
        sheet.getRange(nr, c+2).setValue(parseFloat(data.amount)); break;
      }
    }
    budget_applyFormulas(sheet, nr);
    return { success: true, id: nId };
  } else {
    let target = -1; for (let i=0; i<idData.length; i++) if(idData[i][0] == data.id) { target = i + 5; break; }
    if (target === -1) throw new Error("ID not found");
    
    // Check lock value
    const currentLiq = sheet.getRange(target, 4).getValue();
    const currentLiqStr = (currentLiq instanceof Date) ? currentLiq.toISOString() : (currentLiq ? currentLiq.toString() : "");
    if (currentLiqStr === '2025-10-28T17:00:00.000Z' || currentLiqStr === '69-05-00033') throw new Error("รายการนี้ถูกตัดยอดแล้ว (Locked: 69-05-00033)");

    if (data.name) sheet.getRange(target, 9).setValue(data.name);
    if (data.desc) sheet.getRange(target, 13).setValue(data.desc);
    if (data.colF) sheet.getRange(target, 6).setValue(data.colF);
    sheet.getRange(target, 3).setValue(new Date()); 
    sheet.getRange(target, 4).setValue(data.liquidateRefNo);
    sheet.getRange(target, 8).setValue(data.liquidateLetterDate); // Update date in Column H
    const rowVals = sheet.getRange(target, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (let c = 13; c < rowVals.length; c++) {
      if (sheet.getRange(2, c+1).getValue() && rowVals[c] !== "" && typeof rowVals[c] === 'number') {
        sheet.getRange(target, c+2).setValue(parseFloat(data.amount)); break;
      }
    }
    budget_applyFormulas(sheet, target);
    return { success: true };
  }
}

function budget_submitOffset(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const idData = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  let pIdx = -1; let mSub = 0; let insAt = -1;
  const pIdStr = data.id.toString();
  for (let i = 0; i < idData.length; i++) {
    const cur = idData[i][0].toString();
    if (cur === pIdStr) { pIdx = i + 5; insAt = pIdx; }
    if (cur.startsWith(pIdStr + ".")) {
      const s = parseInt(cur.split(".")[1]); if (s > mSub) mSub = s;
      insAt = i + 5;
    }
  }
  if (pIdx === -1) throw new Error("Parent ID not found");
  
  // Check lock value on parent
  const parentLiq = sheet.getRange(pIdx, 4).getValue();
  const parentLiqStr = (parentLiq instanceof Date) ? parentLiq.toISOString() : (parentLiq ? parentLiq.toString() : "");
  if (parentLiqStr === '2025-10-28T17:00:00.000Z' || parentLiqStr === '69-05-00033') throw new Error("รายการหลักถูกตัดยอดแล้ว (Locked: 69-05-00033)");

  const nId = pIdStr + "." + (mSub + 1);
  sheet.insertRowAfter(insAt);
  const nr = insAt + 1;
  sheet.getRange(nr, 1).setValue(nId);
  sheet.getRange(nr, 2).setValue("หักล้างเงินยืม");
  if (data.colF) sheet.getRange(nr, 6).setValue(data.colF);
  sheet.getRange(nr, 4).setValue(data.liquidateRefNo);
  [3, 7, 9, 10, 11, 12, 13].forEach(c => sheet.getRange(nr, c).setValue(sheet.getRange(pIdx, c).getValue()));
  sheet.getRange(nr, 8).setValue(data.liquidateLetterDate); // Set new date in Column H
  sheet.getRange(nr, 3).setValue(new Date()); 
  if (data.name) sheet.getRange(nr, 9).setValue(data.name);
  if (data.desc) sheet.getRange(nr, 13).setValue(data.desc);
  const pVals = sheet.getRange(pIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let c = 13; c < pVals.length; c++) {
    if (sheet.getRange(2, c+1).getValue() && pVals[c] !== "" && typeof pVals[c] === 'number') {
      sheet.getRange(nr, c+2).setValue(Math.abs(parseFloat(data.amount || pVals[c])) * -1); break;
    }
  }
  budget_applyFormulas(sheet, nr);
  return { success: true, id: nId };
}

function budget_submitUpdate(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  let target = -1; for (let i=0; i<ids.length; i++) if(ids[i][0] == data.id) { target = i + 5; break; }
  if (target === -1) throw new Error("ID not found");
  sheet.getRange(target, 4).setValue(data.liquidateRefNo || "");
  sheet.getRange(target, 6).setValue(data.colF || "");
  sheet.getRange(target, 7).setValue(data.refNo || "");
  sheet.getRange(target, 8).setValue(data.letterDate || ""); 
  sheet.getRange(target, 9).setValue(data.name || "");
  sheet.getRange(target, 10).setValue(data.dept || "");
  sheet.getRange(target, 11).setValue(data.catCode || "");
  sheet.getRange(target, 13).setValue(data.desc || "");
  if (data.col) {
    const baseCol = parseInt(data.col);
    if (data.amount !== undefined) sheet.getRange(target, baseCol).setValue(data.amount === "" ? "" : parseFloat(data.amount));
    if (data.amountDeduct !== undefined) sheet.getRange(target, baseCol + 1).setValue(data.amountDeduct === "" ? "" : parseFloat(data.amountDeduct));
  }
  if (data.colDD !== undefined) {
    const ddColIndex = (planKey === '2.1 บินสาธาฯ') ? 39 : 108;
    sheet.getRange(target, ddColIndex).setValue(data.colDD || "");
  }
  budget_applyFormulas(sheet, target);
  return { success: true };
}

function budget_submitReserveAdd(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const idData = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  
  let pIdx = -1; let mSub = 0; let insAt = -1;
  const pIdStr = data.parentId.toString();
  for (let i = 0; i < idData.length; i++) {
    const cur = idData[i][0].toString();
    if (cur === pIdStr) { pIdx = i + 5; insAt = pIdx; }
    if (cur.startsWith(pIdStr + ".")) {
      const s = parseInt(cur.split(".")[1]); if (s > mSub) mSub = s;
      insAt = i + 5;
    }
  }
  if (pIdx === -1) throw new Error("Parent ID not found");
  const nId = pIdStr + "." + (mSub + 1);
  sheet.insertRowAfter(insAt);
  const nr = insAt + 1;
  sheet.getRange(nr, 1).setValue(nId);
  sheet.getRange(nr, 2).setValue(data.type === 'PO' ? 'PO' : ("กันเงินเพิ่ม" + (data.remark ? " " + data.remark : "")));
  sheet.getRange(nr, 3).setValue(new Date()); 
  sheet.getRange(nr, 7).setValue(data.refNo);
  sheet.getRange(nr, 8).setValue(data.letterDate);
  sheet.getRange(nr, 9).setValue(data.name);
  sheet.getRange(nr, 10).setValue(data.dept);
  sheet.getRange(nr, 11).setValue(data.catCode);
  sheet.getRange(nr, 13).setValue(data.desc);
  if (data.col) sheet.getRange(nr, parseInt(data.col)).setValue(parseFloat(data.amount));
  budget_applyFormulas(sheet, nr);
  return { success: true, id: nId };
}

function budget_submitDeductAdd(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const idData = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  
  let pIdx = -1; let mSub = 0; let insAt = -1;
  const pIdStr = data.id.toString();
  for (let i = 0; i < idData.length; i++) {
    const cur = idData[i][0].toString();
    if (cur === pIdStr) { pIdx = i + 5; insAt = pIdx; }
    if (cur.startsWith(pIdStr + ".")) {
      const s = parseInt(cur.split(".")[1]); if (s > mSub) mSub = s;
      insAt = i + 5;
    }
  }
  if (pIdx === -1) throw new Error("Parent ID not found");
  
  const parentLiq = sheet.getRange(pIdx, 4).getValue();
  const parentLiqStr = (parentLiq instanceof Date) ? parentLiq.toISOString() : (parentLiq ? parentLiq.toString() : "");
  if (parentLiqStr === '2025-10-28T17:00:00.000Z' || parentLiqStr === '69-05-00033') throw new Error("รายการหลักถูกตัดยอดแล้ว (Locked: 69-05-00033)");

  const nId = pIdStr + "." + (mSub + 1);
  sheet.insertRowAfter(insAt);
  const nr = insAt + 1;
  sheet.getRange(nr, 1).setValue(nId);
  sheet.getRange(nr, 2).setValue("ตัดยอดเพิ่ม");
  if (data.colF) sheet.getRange(nr, 6).setValue(data.colF);
  sheet.getRange(nr, 4).setValue(data.liquidateRefNo);
  [3, 7, 9, 10, 11, 12, 13].forEach(c => sheet.getRange(nr, c).setValue(sheet.getRange(pIdx, c).getValue()));
  sheet.getRange(nr, 8).setValue(data.liquidateLetterDate); // Set new date in Column H
  sheet.getRange(nr, 3).setValue(new Date()); 
  if (data.name) sheet.getRange(nr, 9).setValue(data.name);
  if (data.desc) sheet.getRange(nr, 13).setValue(data.desc);
  const pVals = sheet.getRange(pIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let c = 13; c < pVals.length; c++) {
    if (sheet.getRange(2, c+1).getValue() && pVals[c] !== "" && typeof pVals[c] === 'number') {
      sheet.getRange(nr, c+2).setValue(parseFloat(data.amount)); break;
    }
  }
  budget_applyFormulas(sheet, nr);
  return { success: true, id: nId };
}
function budget_submitCancel(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const idData = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();

  let target = -1;
  for (let i = 0; i < idData.length; i++) {
    if (idData[i][0].toString() === data.id.toString()) { target = i + 5; break; }
  }
  if (target === -1) throw new Error("ID not found");

  // Get current category and amount
  const rowVals = sheet.getRange(target, 1, 1, sheet.getLastColumn()).getValues()[0];
  let amountToCancel = 0;
  let categoryCol = -1;

  // Search from Column 14 (N) onwards
  for (let c = 13; c < rowVals.length; c++) {
    const head = sheet.getRange(2, c + 1).getValue();
    if (head && rowVals[c] !== "" && typeof rowVals[c] === 'number' && rowVals[c] !== 0) {
      amountToCancel = rowVals[c];
      categoryCol = c + 1;
      break;
    }
  }

  // Update logic:
  // 1. Column F (6): Note/System ID -> Set to "ยกเลิก"
  sheet.getRange(target, 6).setValue("ยกเลิก");

  if (categoryCol !== -1) {
    // 2. Set Category column to 0 (Cancel the reservation in that category)
    sheet.getRange(target, categoryCol).setValue(0);
    // 3. Set Column B (2) to the canceled amount
    sheet.getRange(target, 2).setValue(amountToCancel);
  }

  budget_applyFormulas(sheet, target);
  return { success: true };
}

function budget_submitReserveAndDeduct(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  let max = 0; ids.forEach(r => { let v = parseInt(r[0]); if(v > max) max = v; });
  const nextId = max + 1;
  const nr = lastRow + 1;
  
  sheet.getRange(nr, 1).setValue(nextId);
  sheet.getRange(nr, 2).setValue(data.remark || "");
  sheet.getRange(nr, 3).setValue(new Date()); 
  sheet.getRange(nr, 4).setValue(data.liquidateRefNo || "");
  sheet.getRange(nr, 6).setValue(data.colF || "");
  sheet.getRange(nr, 7).setValue(data.refNo);
  sheet.getRange(nr, 8).setValue(data.letterDate);
  sheet.getRange(nr, 9).setValue(data.name);
  sheet.getRange(nr, 10).setValue(data.dept);
  sheet.getRange(nr, 11).setValue(data.catCode);
  sheet.getRange(nr, 13).setValue(data.desc);
  
  if (data.col) {
    const baseCol = parseInt(data.col);
    const amt = parseFloat(data.amount);
    sheet.getRange(nr, baseCol).setValue(amt);
    sheet.getRange(nr, baseCol + 1).setValue(amt); // Same amount for deduction
  }
  
  budget_applyFormulas(sheet, nr);
  return { success: true, id: nextId };
}

function budget_submitUpdateDD(data, planKey) {
  const sheet = budget_getTargetSheet(planKey);
  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(5, 1, Math.max(1, lastRow - 4), 1).getValues();
  let target = -1; 
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] == data.id) { target = i + 5; break; }
  }
  if (target === -1) throw new Error("ID not found");
  
  const colIndex = (planKey === '2.1 บินสาธาฯ') ? 39 : 108;
  sheet.getRange(target, colIndex).setValue(data.colDD || "");
  
  return { success: true };
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('เครื่องมืออัพเดทข้อมูล Excel')
      .addItem('เปิดเครื่องมืออัพเดทข้อมูล', 'showDialog')
      .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('GF_Index')
      .setWidth(800)
      .setHeight(600)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'เครื่องมืออัพเดทข้อมูลจาก Excel ไปยัง Google Sheet');
}

function getSheetData(spreadsheetId, sheetName) {
  try {
    Logger.log("Getting sheet data for spreadsheetId: " + spreadsheetId + ", sheetName: " + sheetName);
    
    if (!spreadsheetId) {
      Logger.log("No spreadsheetId provided, using default");
      spreadsheetId = "1lwDV6QKiJ0nx4guP6T8BhIMquCmwIqzn-zdVnjYa_hc";
    }
    
    var ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log("No active spreadsheet, opening by ID: " + spreadsheetId);
        ss = SpreadsheetApp.openById(spreadsheetId);
      }
    } catch (e) {
      Logger.log("Error getting active spreadsheet: " + e + ". Trying to open by ID: " + spreadsheetId);
      ss = SpreadsheetApp.openById(spreadsheetId);
    }
    
    if (!ss) {
      throw new Error("Could not access any spreadsheet");
    }
    
    if (!sheetName) {
      Logger.log("No sheetName provided, using default");
      sheetName = "2.บิน";
    }
    
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      var availableSheets = ss.getSheets().map(s => s.getName()).join(", ");
      Logger.log("Sheet '" + sheetName + "' not found. Available sheets: " + availableSheets);
      
      sheet = ss.getSheets()[0];
      if (!sheet) {
        throw new Error("ไม่พบชีทใดๆ ในสเปรดชีท");
      }
      
      Logger.log("Using first sheet instead: " + sheet.getName());
    }
    
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    if (!values || values.length === 0) {
      Logger.log("No data found in sheet");
      throw new Error("ไม่พบข้อมูลในชีท");
    }
    
    Logger.log("Successfully retrieved " + values.length + " rows from sheet " + sheet.getName());
    
    return JSON.stringify(values);
  } catch(error) {
    Logger.log("Error in getSheetData: " + error);
    throw new Error("Failed to get sheet data: " + error.message);
  }
}

function updateSheetData(updates, config) {
  try {
    var updatesArray = JSON.parse(updates);
    var configObj = JSON.parse(config);
    
    var spreadsheetId = configObj.spreadsheetId || "1ZyRiEFjIDJAW5-26ZyYO_L8Ipk55ImzVtSe33vDYKms";
    var sheetName = configObj.sheetName || "GF กิจสนับสนุน";
    var updateColumnLetter = configObj.updateColumn || "C";
    
    var updateColumnIndex = 0;
    updateColumnLetter = updateColumnLetter.toUpperCase();
    for (var i = 0; i < updateColumnLetter.length; i++) {
      updateColumnIndex = updateColumnIndex * 26 + (updateColumnLetter.charCodeAt(i) - 64);
    }
    
    var ss;
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        ss = SpreadsheetApp.openById(spreadsheetId);
      }
    } catch (e) {
      ss = SpreadsheetApp.openById(spreadsheetId);
    }
    
    if (!ss) {
      throw new Error("Could not access any spreadsheet");
    }
    
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.getSheetByName("Sheet1");
      if (!sheet) {
        sheet = ss.getActiveSheet();
        if (!sheet) {
          throw new Error("Could not find any valid sheet to update");
        }
      }
    }
    
    var count = 0;
    var skippedCount = 0;
    
    Logger.log("Received " + updatesArray.length + " updates to process");
    Logger.log("Using column " + updateColumnLetter + " (index " + updateColumnIndex + ") for updates");
    
    for (var i = 0; i < updatesArray.length; i++) {
      var update = updatesArray[i];
      var cell = sheet.getRange(update.row, updateColumnIndex);
      var currentValue = cell.getValue();
      
      if (currentValue === "" || currentValue === 0 || currentValue === "0") {
        cell.setValue(update.value);
        count++;
        Logger.log("Updated row " + update.row + " with value: " + update.value + " (from column " + (update.sourceColumn || "unknown") + ")");
      } else {
        skippedCount++;
        Logger.log("Skipped row " + update.row + " because it already has value: " + currentValue);
      }
    }
    
    var message = "ดำเนินการเสร็จสิ้น อัพเดทข้อมูลแล้ว " + count + " แถว";
    if (skippedCount > 0) {
      message += " ข้ามไป " + skippedCount + " แถว เนื่องจากมีค่าอยู่แล้ว (ไม่ใช่ 0 หรือว่างเปล่า)";
    }
    
    var duplicateCount = 0;
    var matchMap = {};
    
    for (var i = 0; i < updatesArray.length; i++) {
      var matchValue = updatesArray[i].matchValue;
      if (matchValue) {
        if (!matchMap[matchValue]) {
          matchMap[matchValue] = 1;
        } else {
          matchMap[matchValue]++;
          if (matchMap[matchValue] === 2) {
            duplicateCount++;
          }
        }
      }
    }
    
    if (duplicateCount > 0) {
      message += " พบค่าซ้ำ " + duplicateCount + " รายการในคอลัมน์ " + configObj.updateColumn;
    }
    
    Logger.log(message);
    
    var columnFCount = 0;
    for (var i = 0; i < updatesArray.length; i++) {
      if (updatesArray[i].sourceColumn === configObj.alternativeColumn) {
        columnFCount++;
      }
    }
    
    return {
      success: true,
      message: message,
      updatedCount: count,
      skippedCount: skippedCount,
      duplicateCount: duplicateCount,
      columnFCount: columnFCount
    };
  } catch(error) {
    Logger.log("Error in updateSheetData: " + error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}

function compareValuesForSearch(val1, val2) {
  if (!val1 || !val2) return false;
  
  var str1 = String(val1).trim();
  var str2 = String(val2).trim();
  
  if (str1 === str2) return true;
  
  if (str1.toLowerCase() === str2.toLowerCase()) return true;
  
  if (str1.replace(/\s+/g, '') === str2.replace(/\s+/g, '')) return true;
  
  if (!isNaN(str1) && !isNaN(str2)) {
    return parseFloat(str1) === parseFloat(str2);
  }
  
  return false;
}

function searchAmounts(excelDataStr, configStr) {
  try {
    var excelData = JSON.parse(excelDataStr);
    var config = JSON.parse(configStr);
    
    Logger.log("searchAmounts called with " + excelData.length + " rows");
    
    var spreadsheetId = config.spreadsheetId;
    var sheetName = config.sheetName;
    
    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        message: "กรุณาระบุ Spreadsheet ID และชื่อชีท"
      };
    }
    
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        message: "ไม่พบชีท: " + sheetName
      };
    }
    
    var range = sheet.getDataRange();
    var sheetValues = range.getValues();
    Logger.log("Sheet has " + sheetValues.length + " rows");
    
    var matchedItems = [];
    var unmatchedItems = [];
    var totalChecked = 0;
    
    for (var i = 0; i < excelData.length; i++) {
      var excelRow = excelData[i];
      
      if (!excelRow || !Array.isArray(excelRow)) {
        continue;
      }
      
      var excelColumnA = excelRow[0];
      var excelColumnG = excelRow.length > 6 ? parseFloat(excelRow[6]) : null;
      excelColumnG = (excelColumnG && !isNaN(excelColumnG)) ? excelColumnG : null;
      
      if (!excelColumnA || excelColumnA === '') {
        continue;
      }
      
      if (excelColumnG === null) {
        continue;
      }
      
      totalChecked++;
      
      var matchingRows = [];
      var sheetSum = 0;
      
      for (var j = 10; j < sheetValues.length; j++) {
        var sheetRow = sheetValues[j];
        
        if (!sheetRow || sheetRow.length < 6) continue;
        
        var columnF = sheetRow[5];
        var columnB = sheetRow[1] || '';
        var columnDB = sheetRow.length > 104 ? parseFloat(sheetRow[104]) : 0;
        columnDB = columnDB || 0;
        
        if (compareValuesForSearch(excelColumnA, columnF)) {
          matchingRows.push({
            rowIndex: j + 1,
            columnF: columnF,
            columnB: columnB,
            columnDB: columnDB
          });
          sheetSum += columnDB;
        }
      }
      
      if (matchingRows.length === 0) {
        var item = {
          excelRowA: excelColumnA,
          excelRowNumber: i + 1,
          excelColumnCZ: excelColumnG,
          matchingRowsCount: 0,
          noMatch: true
        };
        unmatchedItems.push(item);
      } else if (excelColumnG === sheetSum) {
        var item = {
          excelRowA: excelColumnA,
          excelRowNumber: i + 1,
          excelColumnCZ: excelColumnG,
          matchingRowsCount: matchingRows.length,
          sheetSum: sheetSum,
          matchingRows: matchingRows
        };
        matchedItems.push(item);
      } else {
        var item = {
          excelRowA: excelColumnA,
          excelRowNumber: i + 1,
          excelColumnCZ: excelColumnG,
          matchingRowsCount: matchingRows.length,
          sheetSum: sheetSum,
          difference: excelColumnG - sheetSum,
          matchingRows: matchingRows
        };
        unmatchedItems.push(item);
      }
    }
    
    return {
      success: true,
      results: {
        totalChecked: totalChecked,
        matched: matchedItems.length,
        unmatched: unmatchedItems.length,
        matchedItems: matchedItems,
        unmatchedItems: unmatchedItems
      }
    };
    
  } catch(error) {
    Logger.log("Error in searchAmounts: " + error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}

function findBalance(searchValue, spreadsheetId, sheetName) {
  try {
    if (!searchValue) {
      return {
        success: false,
        message: "กรุณาระบุค่าที่ต้องการค้นหา"
      };
    }
    
    spreadsheetId = spreadsheetId || "1ZyRiEFjIDJAW5-26ZyYO_L8Ipk55ImzVtSe33vDYKms";
    sheetName = sheetName || "GF กิจสนับสนุน";
    
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error("ไม่พบชีท: " + sheetName);
    }
    
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    var results = [];
    
    for (var i = 4; i < values.length; i++) {
      if (values[i] && values[i].length > 3) {
        var columnD = values[i][3];
        
        if (compareValuesForSearch(searchValue, columnD)) {
          var balance = 0;
          if (values[i].length > 63) {
            balance = parseFloat(values[i][63]) || 0;
          }
          
          results.push({
            rowNumber: i + 1,
            columnD: columnD,
            balance: balance
          });
        }
      }
    }
    
    if (results.length === 0) {
      return {
        success: false,
        message: "ไม่พบข้อมูล",
        searchValue: searchValue
      };
    }
    
    var totalBalance = 0;
    for (var j = 0; j < results.length; j++) {
      totalBalance += results[j].balance;
    }
    
    return {
      success: true,
      searchValue: searchValue,
      matchCount: results.length,
      totalBalance: totalBalance,
      details: results
    };
    
  } catch(error) {
    Logger.log("Error in findBalance: " + error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}
