/**
 * Tiny Mail Club 🌿 (Email-based + Auto Delivery Dates + Foreign Shipping + Update Info)
 */

const COURSE_SHEET = 'แพ็กเกจ';
const REGISTRATION_SHEET = 'ลงทะเบียน';
const CONFIG_SHEET = 'บัญชีรับเงิน';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ลงทะเบียน Tiny Mail Club')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("ไม่สามารถเชื่อมต่อ Sheet ได้ กรุณาสร้าง Script จากไฟล์ Sheet ของคุณ");
  return ss;
}

// === ตรวจสอบสมาชิกเก่าจาก Email ===
function verifyMemberByEmail(email) {
  try {
    const ss = checkSpreadsheet_();
    const sheet = ss.getSheetByName(REGISTRATION_SHEET);
    if (!sheet) return { success: false, message: 'ไม่พบข้อมูลสมาชิกในระบบค่ะ กรุณาลงทะเบียนก่อน หรือเช็คความถูกต้องของอีเมลอีกครั้งนะคะ' };

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const rowEmail = String(data[i][5]).trim().toLowerCase();
      if (rowEmail === String(email).trim().toLowerCase()) {
        return {
          success: true,
          name: data[i][2],
          nickname: data[i][3],
          address: data[i][7],       
          prevEndMonth: data[i][13]  
        };
      }
    }
    return { success: false, message: 'ไม่พบข้อมูลสมาชิกในระบบค่ะ กรุณาลงทะเบียนก่อน หรือเช็คความถูกต้องของอีเมลอีกครั้งนะคะ' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// === ตั้งค่าบัญชีธนาคาร ===
function getConfig() {
  try {
    const ss = checkSpreadsheet_();
    let sheet = ss.getSheetByName(CONFIG_SHEET);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG_SHEET);
      sheet.appendRow(['ธนาคาร', 'กสิกรไทย']);
      sheet.appendRow(['เลขบัญชี', '123-4-56789-0']);
      sheet.appendRow(['ชื่อบัญชี', 'บริษัท ตัวอย่าง จำกัด']);
      sheet.getRange("A1:B1").setBackground("#e4e8d3").setFontColor("#556b2f").setFontWeight("bold");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return { bankName: '', accountNumber: '', accountName: '' };
    const data = sheet.getRange(1, 1, lastRow, 2).getValues();
    const map = {};
    data.forEach(row => { if (String(row[0]).trim()) map[String(row[0]).trim()] = String(row[1]).trim(); });

    return {
      bankName:      map['ธนาคาร']    || '',
      accountNumber: map['เลขบัญชี']  || '',
      accountName:   map['ชื่อบัญชี'] || ''
    };
  } catch (err) { return { bankName: '', accountNumber: '', accountName: '', error: err.message }; }
}

// === จัดการแพ็กเกจ ===
function getProducts() {
  try {
    const ss = checkSpreadsheet_();
    let sheet = ss.getSheetByName(COURSE_SHEET);
    
    if (!sheet) {
      sheet = ss.insertSheet(COURSE_SHEET);
      sheet.appendRow(['ชื่อแพ็กเกจ', 'ราคาฐาน (บาท)', 'จำนวนเดือน']);
      sheet.appendRow(['แพ็กเกจ 1 เดือนๆ ละ 209 บาท', 209, 1]);
      sheet.appendRow(['แพ็กเกจ 3 เดือนๆ ละ 179 บาท', 537, 3]);
      sheet.appendRow(['แพ็กเกจ 6 เดือนๆ ละ 159 บาท', 954, 6]);
      sheet.appendRow(['ทดลองระบบ 1 บาท (สำหรับทดสอบ)', 1, 1]);
      sheet.getRange("A1:C1").setBackground("#e4e8d3").setFontColor("#556b2f").setFontWeight("bold");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { products: [], error: 'ยังไม่มีข้อมูลแพ็กเกจ' };
    
    const products = sheet.getRange(2, 1, lastRow - 1, 3).getValues()
      .filter(r => r[0] && r[1] !== '')
      .map(r => {
        let m = Number(r[2]);
        if (!m || isNaN(m)) {
          const n = String(r[0]);
          if (n.includes('6 เดือน')) m = 6;
          else if (n.includes('3 เดือน')) m = 3;
          else m = 1;
        }
        return { name: String(r[0]).trim(), price: Number(r[1]), months: m };
      });
    return { products };
  } catch (err) { return { products: [], error: err.message }; }
}

// === บันทึกไฟล์รูปภาพลง Google Drive ===
function saveFileToDrive_(fileData) {
  try {
    const decoded = Utilities.base64Decode(fileData.base64);
    const blob = Utilities.newBlob(decoded, fileData.mime, 'slip_' + new Date().getTime() + '_' + fileData.name);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    throw new Error('ไม่สามารถบันทึกไฟล์รูปภาพได้: ' + err.message);
  }
}

// === บันทึกข้อมูล ===
function submitRegistration(payload) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'ระบบคิวเต็ม กรุณาลองใหม่อีกครั้งค่ะ' };

  try {
    const { 
      regType, email, product, price, fileData, shippingZone,
      deliveryLetter, deliveryShip, startMonthVal, endMonthVal
    } = payload;
    
    let finalData = { ...payload };
    const sheet = getRegistrationSheet_();

    if (regType === 'renew') {
      let found = false;
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][5]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
          finalData.name = data[i][2];
          finalData.birthday = data[i][4];
          // ดึงค่า Nickname และ Address ใหม่มาจาก payload หน้าเว็บแทนการใช้ของเดิม (ถ้ามีการอัปเดต)
          finalData.nickname = payload.nickname; 
          finalData.address = payload.address;
          finalData.source = 'ต่ออายุสมาชิก';
          found = true;
          break;
        }
      }
      if (!found) return { success: false, message: 'ไม่พบข้อมูลสมาชิกในระบบค่ะ กรุณาลงทะเบียนก่อน หรือเช็คความถูกต้องของอีเมลอีกครั้งนะคะ' };
    } else {
      if (finalData.birthday && finalData.birthday.includes('-')) {
        const parts = finalData.birthday.split('-');
        if (parts.length === 3) finalData.birthday = `${parts[2]}/${parts[1]}/${parts[0]}`;
      }
    }

    let slipUrl = '';
    if (fileData && fileData.base64) {
      slipUrl = saveFileToDrive_(fileData);
    } else {
      return { success: false, message: 'ไม่พบข้อมูลไฟล์สลิป' };
    }
    finalData.slipUrl = slipUrl;

    sheet.appendRow([
      new Date(),
      (regType === 'new' ? 'สมัครใหม่' : 'ต่ออายุ'),
      finalData.name, finalData.nickname, finalData.birthday, finalData.email, 
      finalData.shippingZone, finalData.address, finalData.source, 
      finalData.product, Number(finalData.price), 
      finalData.slipUrl, 
      finalData.startMonthVal, finalData.endMonthVal, 
      'รอตรวจสอบ' 
    ]);
    
    try {
      sendEmails_(finalData);
    } catch (e) {
      Logger.log(e);
      return { success: true, message: '🌿 สมัครเสร็จสิ้น! (ระบบส่งอีเมลยืนยันมีปัญหาเล็กน้อย)' };
    }

    return { success: true, message: '🌿 รายการสำเร็จ! ยินดีต้อนรับสู่ Tiny Mail Club ค่ะ 💌' };
  } catch (err) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function getRegistrationSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REGISTRATION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRATION_SHEET);
    sheet.appendRow([
      'วันที่ลงทะเบียน', 'ประเภท', 'ชื่อ-นามสกุล (สำหรับส่งจดหมาย)', 'ชื่อที่ให้เขียนถึง', 'วันเกิด (DD/MM/YYYY)', 'อีเมล', 
      'การจัดส่ง', 'ที่อยู่จัดส่ง', 'ช่องทางที่รู้จัก', 
      'แพ็กเกจ', 'ยอดที่แจ้งชำระ (บาท)', 'ลิงก์สลิปโอนเงิน', 
      'เล่มเริ่ม (YYYY-MM)', 'เล่มจบ (YYYY-MM)', 'สถานะ'
    ]);
    sheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#e4e8d3').setFontColor('#556b2f');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sendEmails_(p) {
  const ownerEmail = PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL') || '';
  const priceFmt = Number(p.price).toLocaleString('th-TH');
  
  if (p.email) {
    MailApp.sendEmail({
      to: p.email,
      subject: '🌿 ยืนยันข้อมูลสมาชิก Tiny Mail Club',
      htmlBody: emailUserHtml_(p, priceFmt)
    });
  }

  if (ownerEmail) {
    MailApp.sendEmail({
      to: ownerEmail,
      subject: `✨ [${p.regType === 'new' ? 'สมัครใหม่' : 'ต่ออายุ'}] — ${p.nickname}`,
      htmlBody: emailOwnerHtml_(p, priceFmt)
    });
  }
}

function emailUserHtml_(p, priceFmt) {
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;padding:24px;border:2px solid #e4e8d3;border-radius:16px;background:#f9fbf4;">' +
    '<h2 style="color:#556b2f;text-align:center;">🌿 ยินดีต้อนรับสู่คลับของเรา</h2>' +
    '<p>สวัสดีคุณ <b>' + p.nickname + '</b>,</p>' +
    '<p>เราได้รับข้อมูลการ ' + (p.regType === 'new' ? 'ลงทะเบียน' : 'ต่ออายุสมาชิก') + ' เรียบร้อยแล้วค่ะ ดีใจมากๆ ที่ได้มาร่วมเป็นส่วนหนึ่งของ Tiny Mail Club นะคะ ✨</p>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:16px;">' +
    row_('แพ็กเกจ', p.product) +
    row_('การจัดส่ง', p.shippingZone) +
    row_('ยอดที่แจ้งชำระ', priceFmt + ' บาท') +
    row_('ที่อยู่จัดส่ง', p.address) +
    '</table>' +
    '<div style="background:#e4e8d3; padding:16px; border-radius:12px; margin-top:24px; text-align:center; color:#556b2f; font-weight:bold; font-size:15px; line-height: 1.6;">' +
    'คุณจะได้รับจดหมายฉบับ' + p.deliveryLetter + '<br>(รอบแรกจัดส่งวันที่ ' + p.deliveryShip + ')<br><br>' +
    '<span style="color:#c62828; font-size: 13px;">* เราจะส่งอีเมลแจ้งเตือนให้ต่ออายุสมาชิก ภายใน 7 วันก่อนถึงรอบส่งถัดไป เพื่อจะไม่พลาดจดหมายจากเรานะคะ *</span>' +
    '</div>' +
    '<p style="margin-top:16px; text-align:center; color:#6b8e23;">ทีมงานจะทำการตรวจสอบสลิปการโอนเงิน เตรียมพบกับความน่ารักสดใสของเด็กคนนั้นในตัวคุณ + ของขวัญพิเศษในเดือนเกิดได้เลยค่ะ 💌</p>' +
    '</div>';
}

function emailOwnerHtml_(p, priceFmt) {
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:auto;padding:24px;">' +
    '<h2 style="color:#556b2f;">🌿 มีผู้' + (p.regType === 'new' ? 'สมัครใหม่' : 'ต่ออายุ') + '!</h2>' +
    '<p><b>กรุณาตรวจสอบสลิปและกดยืนยันใน Google Sheets</b></p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    row_('ชื่อที่ให้เขียนถึง', p.nickname) +
    row_('ชื่อ-นามสกุล', p.name) +
    row_('อีเมล', p.email) +
    row_('การจัดส่ง', p.shippingZone) +
    row_('ที่อยู่จัดส่ง', p.address) +
    row_('รู้จักจาก', p.source) +
    row_('แพ็กเกจ', p.product) +
    row_('ยอดที่แจ้งชำระ', priceFmt + ' บาท') +
    row_('ลิงก์ตรวจสอบสลิป', '<a href="' + p.slipUrl + '" target="_blank">คลิกเพื่อดูรูปสลิป</a>') +
    row_('เล่มเริ่ม-จบ', p.startMonthVal + ' ถึง ' + p.endMonthVal) +
    '</table></div>';
}

function row_(label, value) {
  return '<tr>' +
    '<td style="padding:12px;border:1px solid #e4e8d3;background:#f2f4e8;width:35%;color:#556b2f;"><b>' + label + '</b></td>' +
    '<td style="padding:12px;border:1px solid #e4e8d3;color:#333;">' + value + '</td>' +
    '</tr>';
}