const PLAN_CONFIG = {
  '1.ฝนหลวง':     { id: '1E43wKRlEa1xKM5l_zSDdstmAKPtDgw05rwJVeRLgTYk', sheet: '1.ฝน' },
  '2.ด้านการบิน':   { id: '1lwDV6QKiJ0nx4guP6T8BhIMquCmwIqzn-zdVnjYa_hc', sheet: '2.บิน' },
  '2.1 บินสาธาฯ': { id: '1lwDV6QKiJ0nx4guP6T8BhIMquCmwIqzn-zdVnjYa_hc', sheet: '2.1 บินสาธาฯ' },
  '3.แก้ปัญหาฝุ่น':   { id: '1iOX9ISvptU2UpMbVNPTuYHbA4df11dytCHxFYKoSQxg', sheet: '3.ฝุ่น' },
  '4.บรรเทาลูกเห็บ': { id: '1SM0GYznjAKPOQT-itypsk83n7l5rx9ibWzKcLGZb4H4', sheet: '4.ลูกเห็บ' }
};

function doGet(e) {
  try {
    const action = e.parameter.action;
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


    return ContentService.createTextOutput(JSON.stringify(res || { error: 'Action not found' })).setMimeType(ContentService.MimeType.JSON);
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
