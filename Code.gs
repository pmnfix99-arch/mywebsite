/**
 * TEAM CONNECT PASSPORT - Backend (Google Apps Script)
 * ------------------------------------------------------
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet ที่จะใช้เป็นฐานข้อมูล แล้วสร้าง 3 ชีทชื่อ "Users", "Activities" และ "StepImages"
 *    - ชีท Users คอลัมน์ (แถวที่ 1 เป็นหัวตาราง):
 *      Username | Password | Role | EmployeeID | FullName | Nickname | Department | Photo
 *      (Role ใส่ค่า "user" หรือ "approver")
 *    - ชีท Activities คอลัมน์:
 *      ID | EmployeeID | FullName | ActivityKey | ActivityName | Date | TimeIn | TimeOut | Status | ApproverName | ApproveTime | Remark | PhotoUrl
 *      (Status เริ่มต้นว่างหรือ "pending", และจะถูกอัปเดตเป็น "approved" / "rejected")
 *      >> ต้องมีคอลัมน์หัวตาราง "Date" ด้วย (เก็บวันที่บันทึกกิจกรรม แยกจาก TimeIn/TimeOut)
 *      >> ถ้าต้องการแนบรูปถ่ายแต่ละกิจกรรม ต้องมีคอลัมน์หัวตาราง "PhotoUrl" ด้วย (เก็บลิงก์รูปใน Google Drive)
 *    - ชีท StepImages คอลัมน์ (ใช้เก็บ "รูปภาพตัวอย่างประจำขั้นตอน" ที่ผู้อนุมัติอัปโหลดครั้งเดียว
 *      แล้วผู้ใช้ทุกคนเห็นรูปเดียวกันตอนเปิดขั้นตอนนั้นๆ เช่นรูปตัวอย่างขั้นตอน 2-6):
 *      StepKey | ImageUrl | UpdatedBy | UpdatedAt
 *      (ไม่ต้องกรอกข้อมูลล่วงหน้า ระบบจะเพิ่ม/อัปเดตแถวให้อัตโนมัติเมื่อผู้อนุมัติอัปโหลดรูปจากหน้า approver.html)
 *    - รูปถ่ายทั้งหมด (รูปโปรไฟล์ + รูปแนบแต่ละกิจกรรม) จะถูกอัปโหลดไปเก็บใน Google Drive
 *      โฟลเดอร์ชื่อ "Team Connect Passport Photos" (สร้างอัตโนมัติถ้ายังไม่มี) แล้วเก็บแค่ลิงก์ไว้ในชีต ไม่ได้เก็บตัวรูปในชีต
 * 2. เปิดเมนู Extensions > Apps Script ในชีทนั้น แล้ววางโค้ดนี้ทับไฟล์ Code.gs
 * 3. แก้ไขค่า SHEET_ID ด้านล่างให้เป็น ID ของ Google Sheet (ดูได้จาก URL ของชีท)
 * 4. กด Deploy > New deployment > เลือกประเภท "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - กด Deploy แล้วคัดลอก Web app URL มาใส่ในไฟล์ HTML (ตัวแปร APPS_SCRIPT_URL)
 * 5. ทุกครั้งที่แก้โค้ดใหม่ ต้องกด Deploy > Manage deployments > แก้ไข (แก้เวอร์ชัน) ใหม่ทุกครั้ง
 */

const SHEET_ID = '18lRWFM9vRRMIdbr9ok7E0edPR8DOch8cIyWg-U8O-0I'; // <-- แก้เป็น Sheet ID ของคุณ
const SHEET_USERS = 'Users';
const SHEET_ACTIVITIES = 'Activities';
const SHEET_STEP_IMAGES = 'StepImages'; // ชีทใหม่: เก็บรูปภาพตัวอย่างประจำแต่ละขั้นตอน (อัปโหลดครั้งเดียวโดยผู้อนุมัติ ผู้ใช้ทุกคนเห็นรูปเดียวกัน)

function _ss() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function _sheet(name) {
  const ss = _ss();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีทชื่อ ' + name);
  return sh;
}

function _readAll(sheetName) {
  const sh = _sheet(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(row => row.join('') !== '') // skip blank rows
    .map((row, idx) => {
      const obj = { _row: idx + 2 }; // actual sheet row number (1 = header)
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function _jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

const PHOTO_FOLDER_NAME = 'Team Connect Passport Photos';

/**
 * หาโฟลเดอร์เก็บรูปใน Drive ถ้ายังไม่มีจะสร้างให้อัตโนมัติ (ทำครั้งเดียว แล้วใช้ซ้ำทุกครั้งถัดไป)
 */
function _getPhotoFolder() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

/**
 * อัปโหลดรูป (ส่งมาเป็น base64) ขึ้น Google Drive แล้วคืนลิงก์รูปกลับไป
 * ตั้งค่าการแชร์เป็น "ทุกคนที่มีลิงก์ดูได้" เพื่อให้เปิดดูรูปจากหน้าเว็บได้โดยไม่ต้องล็อกอิน Google
 * body: { imageBase64, mimeType, filename }
 */
function _uploadPhotoToDrive(body) {
  const base64Data = String(body.imageBase64 || '');
  if (!base64Data) throw new Error('ไม่พบข้อมูลรูปภาพ');

  const mimeType = body.mimeType || 'image/jpeg';
  const filename = body.filename || ('photo_' + new Date().getTime());

  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, filename);

  const folder = _getPhotoFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // ใช้ endpoint แบบ thumbnail เพื่อให้แสดงผลเป็น <img> ได้ตรงๆ (ลิงก์แบบเปิดหน้า Drive ปกติจะฝัง <img> ไม่ได้)
  const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
  return { url: url, fileId: file.getId() };
}

/**
 * อัปโหลดรูปทั่วไป (ใช้กับรูปแนบกิจกรรม step ต่างๆ) คืนแค่ลิงก์กลับไป
 * ฝั่ง client จะเอาลิงก์นี้ไปแนบต่อใน saveActivity อีกที
 */
function handleUploadPhoto(body) {
  try {
    const result = _uploadPhotoToDrive(body);
    return { ok: true, url: result.url };
  } catch (err) {
    return { ok: false, message: 'อัปโหลดรูปไม่สำเร็จ: ' + err.message };
  }
}

/**
 * อัปโหลดรูปโปรไฟล์ แล้วบันทึกลิงก์ลงคอลัมน์ Photo ในชีต Users ทันที (ใช้กับหน้าข้อมูลส่วนตัว)
 * body: { employeeId, imageBase64, mimeType, filename }
 */
function handleUpdateProfilePhoto(body) {
  const employeeId = String(body.employeeId || '').trim();
  if (!employeeId) return { ok: false, message: 'ไม่พบรหัสพนักงาน' };

  let uploadResult;
  try {
    uploadResult = _uploadPhotoToDrive(body);
  } catch (err) {
    return { ok: false, message: 'อัปโหลดรูปไม่สำเร็จ: ' + err.message };
  }

  const sh = _sheet(SHEET_USERS);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);
    const photoCol = col('Photo');
    if (photoCol === -1) {
      return { ok: false, message: 'ไม่พบคอลัมน์ "Photo" ในชีต Users' };
    }

    let targetRow = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][col('EmployeeID')]).trim() === employeeId) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) return { ok: false, message: 'ไม่พบข้อมูลพนักงาน' };

    sh.getRange(targetRow, photoCol + 1).setValue(uploadResult.url);
  } finally {
    lock.releaseLock();
  }

  return { ok: true, url: uploadResult.url, message: 'อัปเดตรูปโปรไฟล์สำเร็จ' };
}

/**
 * คืนรูปภาพตัวอย่างของทุกขั้นตอนที่เคยอัปโหลดไว้ (ให้ user.html เรียกครั้งเดียวตอนโหลดหน้า)
 * คืนกลับมาเป็น object { stepKey: imageUrl, ... } เพื่อให้ front-end ใช้งานง่าย
 * ถ้ายังไม่เคยสร้างชีท StepImages เลย ให้ถือว่ายังไม่มีรูป (ไม่ error) เพื่อไม่ให้หน้า user พังตอนยังไม่ได้ตั้งค่า
 */
function handleGetStepImages(body) {
  const ss = _ss();
  const sh = ss.getSheetByName(SHEET_STEP_IMAGES);
  if (!sh) return { ok: true, images: {} };

  const rows = _readAll(SHEET_STEP_IMAGES);
  const images = {};
  rows.forEach(r => {
    const key = String(r.StepKey || '').trim();
    if (key) images[key] = r.ImageUrl || '';
  });
  return { ok: true, images: images };
}

/**
 * ผู้อนุมัติอัปโหลดรูปภาพตัวอย่างประจำขั้นตอน (ใช้ครั้งเดียว ผู้ใช้ทุกคนเห็นรูปเดียวกัน)
 * body: { stepKey, imageBase64, mimeType, filename, updatedBy }
 * ถ้ามีแถวของ stepKey นั้นอยู่แล้วจะอัปเดตทับ ถ้ายังไม่มีจะเพิ่มแถวใหม่
 * ถ้ายังไม่เคยสร้างชีท StepImages จะสร้างให้อัตโนมัติพร้อมหัวตาราง
 */
function handleUpdateStepImage(body) {
  const stepKey = String(body.stepKey || '').trim();
  if (!stepKey) return { ok: false, message: 'ไม่พบ stepKey' };

  let uploadResult;
  try {
    uploadResult = _uploadPhotoToDrive(body);
  } catch (err) {
    return { ok: false, message: 'อัปโหลดรูปไม่สำเร็จ: ' + err.message };
  }

  const ss = _ss();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let sh = ss.getSheetByName(SHEET_STEP_IMAGES);
    if (!sh) {
      sh = ss.insertSheet(SHEET_STEP_IMAGES);
      sh.getRange(1, 1, 1, 4).setValues([['StepKey', 'ImageUrl', 'UpdatedBy', 'UpdatedAt']]);
    }

    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);

    let targetRow = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][col('StepKey')]).trim() === stepKey) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) targetRow = data.length + 1;

    const rowValues = [stepKey, uploadResult.url, body.updatedBy || '', new Date()];
    sh.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true, url: uploadResult.url, message: 'อัปเดตรูปภาพตัวอย่างสำเร็จ' };
}

/**
 * Web app entry points.
 * Front-end ส่งคำขอมาทาง POST เสมอ (เพื่อเลี่ยงปัญหา URL length และ CORS preflight)
 * โดยส่ง body เป็น JSON string ผ่าน Content-Type: text/plain;charset=utf-8
 * รูปแบบ: { action: "...", ...ข้อมูลอื่นๆ }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch (action) {
      case 'login':
        result = handleLogin(body);
        break;
      case 'getProfile':
        result = handleGetProfile(body);
        break;
      case 'getActivities':
        result = handleGetActivities(body);
        break;
      case 'saveActivity':
        result = handleSaveActivity(body);
        break;
      case 'getPendingApprovals':
        result = handleGetPendingApprovals(body);
        break;
      case 'approveActivity':
        result = handleApproveActivity(body, 'approved');
        break;
      case 'rejectActivity':
        result = handleApproveActivity(body, 'rejected');
        break;
      case 'uploadPhoto':
        result = handleUploadPhoto(body);
        break;
      case 'updateProfilePhoto':
        result = handleUpdateProfilePhoto(body);
        break;
      case 'getStepImages':
        result = handleGetStepImages(body);
        break;
      case 'updateStepImage':
        result = handleUpdateStepImage(body);
        break;
      default:
        result = { ok: false, message: 'ไม่รู้จัก action: ' + action };
    }
    return _jsonOutput(result);
  } catch (err) {
    return _jsonOutput({ ok: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
}

function doGet(e) {
  return _jsonOutput({ ok: true, message: 'TEAM CONNECT PASSPORT API พร้อมใช้งาน (ใช้ POST เท่านั้น)' });
}

/* ---------------- Handlers ---------------- */

function handleLogin(body) {
  const username = String(body.username || '').trim();
  const password = String(body.password || '').trim();
  const roleWanted = body.role; // 'user' หรือ 'approver'

  const users = _readAll(SHEET_USERS);
  const found = users.find(u =>
    String(u.Username).trim() === username &&
    String(u.Password).trim() === password
  );

  if (!found) {
    return { ok: false, message: 'Username หรือ Password ไม่ถูกต้อง' };
  }
  if (roleWanted && String(found.Role).trim() !== roleWanted) {
    return { ok: false, message: 'บัญชีนี้ไม่มีสิทธิ์เข้าใช้งานในส่วนนี้' };
  }

  return {
    ok: true,
    role: found.Role,
    employeeId: found.EmployeeID,
    fullName: found.FullName,
    nickname: found.Nickname,
    department: found.Department,
    photo: found.Photo || ''
  };
}

function handleGetProfile(body) {
  const employeeId = String(body.employeeId || '').trim();
  const users = _readAll(SHEET_USERS);
  const found = users.find(u => String(u.EmployeeID).trim() === employeeId);
  if (!found) return { ok: false, message: 'ไม่พบข้อมูลพนักงาน' };
  return {
    ok: true,
    fullName: found.FullName,
    nickname: found.Nickname,
    department: found.Department,
    photo: found.Photo || ''
  };
}

/**
 * แปลงเวลาแบบ "09:30" ให้เป็น "09.30" (ใช้จุดแทนโคลอน ตามที่ต้องการแสดงผล)
 */
function toDotTime(timeStr) {
  if (!timeStr) return '';
  return String(timeStr).trim().replace(':', '.');
}

/**
 * ป้องกันไว้เผื่อมีแถวเก่าที่บันทึกไว้ก่อนแก้บั๊ก (เคยถูก Sheets แปลงเป็น Date/time อัตโนมัติ)
 * ถ้าเจอค่าที่ยังเป็น Date object อยู่ ให้แปลงเป็นข้อความเวลาที่อ่านง่าย แทนที่จะส่ง ISO string ดิบๆ ไปให้หน้าเว็บ
 */
function normalizeLegacyTime(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Bangkok', 'HH.mm') + ' น. (รายการเก่า ไม่มีวันที่บันทึกไว้)';
  }
  return value;
}

/**
 * รองรับ 3 รูปแบบข้อมูลที่อาจเจอในคอลัมน์ Date/TimeIn/TimeOut:
 * 1) ข้อมูลใหม่ (หลังเพิ่มคอลัมน์ Date): Date="09/07/2026", TimeIn="09.30" แยกกันอยู่แล้ว -> ใช้ตรงๆ
 * 2) ข้อมูลช่วงกลาง (ก่อนมีคอลัมน์ Date แต่หลังแก้บั๊กแรก): TimeIn="09/07/2026 09.30" รวมวันที่ไว้ในช่องเดียว -> แยกออกมา
 * 3) ข้อมูลเก่าสุด (ก่อนแก้บั๊กใดๆ): TimeIn เป็น Date object ที่ไม่มีวันที่จริง -> แสดงเวลาอย่างเดียว
 */
function reconcileDateTime(a) {
  let dateVal = a.Date || '';
  let timeIn = a.TimeIn;
  let timeOut = a.TimeOut;

  if (timeIn instanceof Date || timeOut instanceof Date) {
    return { Date: dateVal || '(ไม่ทราบวันที่)', TimeIn: normalizeLegacyTime(timeIn), TimeOut: normalizeLegacyTime(timeOut) };
  }

  if (!dateVal) {
    const splitOne = value => {
      if (typeof value === 'string' && value.indexOf('/') !== -1) {
        const spaceIdx = value.indexOf(' ');
        return spaceIdx === -1 ? { date: value, time: '' } : { date: value.slice(0, spaceIdx), time: value.slice(spaceIdx + 1) };
      }
      return null;
    };
    const splitIn = splitOne(timeIn);
    const splitOut = splitOne(timeOut);
    if (splitIn) { dateVal = splitIn.date; timeIn = splitIn.time; }
    if (splitOut) { dateVal = dateVal || splitOut.date; timeOut = splitOut.time; }
  }

  return { Date: dateVal, TimeIn: timeIn, TimeOut: timeOut };
}

function handleGetActivities(body) {
  const employeeId = String(body.employeeId || '').trim();
  const all = _readAll(SHEET_ACTIVITIES);
  const mine = all
    .filter(a => String(a.EmployeeID).trim() === employeeId)
    .map(a => Object.assign({}, a, reconcileDateTime(a)));
  return { ok: true, activities: mine };
}

/**
 * บันทึก/อัปเดตกิจกรรมของพนักงาน 1 กิจกรรม (เช่น "step2_ตัดขอบ")
 * ถ้ามีอยู่แล้ว (EmployeeID + ActivityKey ตรงกัน) จะอัปเดตแถวเดิม
 * ถ้ายังไม่มีจะเพิ่มแถวใหม่ และตั้งสถานะเป็น pending เสมอ (รอผู้อนุมัติ)
 */
function handleSaveActivity(body) {
  const sh = _sheet(SHEET_ACTIVITIES);

  // ล็อกกันสองคำขอมาชนกัน (เช่นกดบันทึกซ้ำเร็วๆ) ตอนกำลังหาแถว/เขียนแถว
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);

    const employeeId = String(body.employeeId || '').trim();
    const activityKey = String(body.activityKey || '').trim();

    const dateCol = col('Date');
    if (dateCol === -1) {
      return { ok: false, message: 'ไม่พบคอลัมน์ "Date" ในชีต Activities กรุณาเพิ่มหัวคอลัมน์ Date ในชีตก่อนใช้งาน (ดูคำแนะนำด้านบนไฟล์)' };
    }

    let targetRow = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][col('EmployeeID')]).trim() === employeeId &&
          String(data[r][col('ActivityKey')]).trim() === activityKey) {
        targetRow = r + 1; // sheet row number
        break;
      }
    }

    // กันบันทึกซ้ำ: ถ้ามีอยู่แล้วและสถานะเป็น pending/approved ห้ามเขียนทับ
    // อนุญาตให้บันทึกใหม่ได้เฉพาะตอนยังไม่เคยมี หรือสถานะเป็น rejected เท่านั้น
    if (targetRow !== -1) {
      const existingStatus = String(data[targetRow - 1][col('Status')]).trim();
      if (existingStatus === 'pending' || existingStatus === 'approved') {
        return { ok: false, message: 'รายการนี้บันทึกไปแล้ว ไม่สามารถบันทึกซ้ำได้ (ยกเว้นกรณีไม่อนุมัติ)' };
      }
    }

    const isNewRow = targetRow === -1;
    if (isNewRow) targetRow = data.length + 1; // แถวใหม่ต่อท้ายข้อมูลที่มีอยู่

    // บังคับให้คอลัมน์ Date/TimeIn/TimeOut เป็น Plain Text ก่อนเขียนค่าเสมอ
    // ต้องทำก่อนเขียนค่าเท่านั้น ไม่งั้น Sheets จะยังตีความค่าที่ดูเหมือนวันที่/เวลาให้เป็น Date/time อัตโนมัติอยู่ดี
    sh.getRange(targetRow, dateCol + 1, 1, 1).setNumberFormat('@');
    sh.getRange(targetRow, col('TimeIn') + 1, 1, 1).setNumberFormat('@');
    sh.getRange(targetRow, col('TimeOut') + 1, 1, 1).setNumberFormat('@');

    const rowValues = {
      ID: isNewRow ? Utilities.getUuid() : data[targetRow - 1][col('ID')],
      EmployeeID: employeeId,
      FullName: body.fullName || '',
      ActivityKey: activityKey,
      ActivityName: body.activityName || '',
      Date: String(body.date || '').trim(),
      TimeIn: toDotTime(body.timeIn),
      TimeOut: toDotTime(body.timeOut),
      Status: 'pending',
      ApproverName: '',
      ApproveTime: '',
      Remark: body.remark || ''
    };

    // แนบลิงก์รูปถ่าย (ถ้ามีการอัปโหลดรูปมาด้วย) — ต้องมีคอลัมน์ "PhotoUrl" ในชีตก่อนถึงจะเก็บได้
    if (body.photoUrl) {
      const photoCol = col('PhotoUrl');
      if (photoCol === -1) {
        return { ok: false, message: 'ไม่พบคอลัมน์ "PhotoUrl" ในชีต Activities กรุณาเพิ่มหัวคอลัมน์ PhotoUrl ในชีตก่อนแนบรูป' };
      }
      rowValues.PhotoUrl = body.photoUrl;
    } else if (!isNewRow) {
      // ไม่มีการอัปโหลดรูปใหม่ตอนนี้ -> คงรูปเดิมไว้ (ถ้ามีคอลัมน์ PhotoUrl อยู่แล้ว)
      const photoCol = col('PhotoUrl');
      if (photoCol !== -1) rowValues.PhotoUrl = data[targetRow - 1][photoCol];
    }

    // เขียนทั้งแถวทีเดียวด้วย setValues (1 round-trip ไปยัง Sheets API แทนที่จะเป็นหลายรอบ)
    const fullRow = headers.map(h => (h in rowValues ? rowValues[h] : ''));
    sh.getRange(targetRow, 1, 1, headers.length).setValues([fullRow]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true, message: 'บันทึกกิจกรรมสำเร็จ รอการอนุมัติ' };
}

function handleGetPendingApprovals(body) {
  const all = _readAll(SHEET_ACTIVITIES);
  const pending = all
    .filter(a => String(a.Status).trim() === 'pending' || !a.Status)
    .map(a => Object.assign({}, a, reconcileDateTime(a)));
  return { ok: true, activities: pending };
}

function handleApproveActivity(body, newStatus) {
  const sh = _sheet(SHEET_ACTIVITIES);
  const id = String(body.id || '').trim();

  // ล็อกกันสองผู้อนุมัติกดพร้อมกันบนรายการเดียวกัน
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);

    let targetRow = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][col('ID')]).trim() === id) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) return { ok: false, message: 'ไม่พบรายการนี้' };

    // กันโดนอนุมัติซ้ำ (เช่นผู้อนุมัติอีกคนกดไปก่อนแล้วระหว่างที่เรารอ lock)
    const currentStatus = String(data[targetRow - 1][col('Status')]).trim();
    if (currentStatus === 'approved' || currentStatus === 'rejected') {
      return { ok: false, message: 'รายการนี้ถูกดำเนินการไปแล้วโดยผู้อนุมัติท่านอื่น' };
    }

    const statusCol = col('Status') + 1;
    const approverCol = col('ApproverName') + 1;
    const approveTimeCol = col('ApproveTime') + 1;

    // เขียน 3 ค่าทีเดียวด้วย setValues (1 round-trip แทนที่จะเรียก setValue 3 รอบ)
    // ใช้ได้เพราะ Status | ApproverName | ApproveTime อยู่ติดกันตามลำดับในชีต
    if (approverCol === statusCol + 1 && approveTimeCol === statusCol + 2) {
      sh.getRange(targetRow, statusCol, 1, 3).setValues([[newStatus, body.approverName || '', new Date()]]);
    } else {
      // fallback เผื่อมีคนย้ายลำดับคอลัมน์ในชีตภายหลัง
      sh.getRange(targetRow, statusCol).setValue(newStatus);
      sh.getRange(targetRow, approverCol).setValue(body.approverName || '');
      sh.getRange(targetRow, approveTimeCol).setValue(new Date());
    }
  } finally {
    lock.releaseLock();
  }

  return { ok: true, message: newStatus === 'approved' ? 'อนุมัติเรียบร้อย' : 'ไม่อนุมัติเรียบร้อย' };
}
