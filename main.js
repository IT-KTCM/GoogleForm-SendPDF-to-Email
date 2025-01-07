let templateSlideId = '1wt42EkRnzplmjSV71E-R0FHVkWN69xJ2Xml5w0IzKuk'; // รหัส (ID) ของ Google Slides Template
let folderResponsePdfId = '1CjbQ9CpNdcPRWGGu5pFkQlN2fpCJDtrm'; // รหัส (ID) ของโฟลเดอร์ที่ใช้เก็บ PDF
let folderResponseSlideId = '1TB0iSwhlarG93besbpb9c7pxH3PtgIYR'; // รหัส (ID) ของโฟลเดอร์ที่ใช้เก็บ Google Slides
let sheetName = 'การตอบแบบฟอร์ม 1'; // ชื่อของ Sheet ที่เก็บข้อมูลจาก Google Form
let pdf_file_name = "รายงานการประเมินการนิเทศฝึกงาน"; // ชื่อไฟล์ PDF ที่จะทำการส่งให้ส่งรายงาน
let isSendEmail = true; // กำหนดให้ส่งอีเมลหรือไม่ (true เพื่อส่ง, false เพื่อไม่ส่ง)
let isSendLine = false; // กำหนดให้ส่ง LINE Notify หรือไม่ (true เพื่อส่ง, false เพื่อไม่ส่ง)
let email_send_default = ['dvt2563ktcm@gmail.com']; // อีเมลที่ต้องการส่งไฟล์ PDF (กรณีไม่มีอีเมลจากผู้ใช้งาน)
var email_subject = 'ขอบคุณสำหรับการประเมินการนิเทศฝึกงาน'; // หัวข้ออีเมลที่ตอบกลับ
var email_message = 'แบบฟอร์มการประเมินการนิเทศฝึกงานได้จัดส่งให้ท่านแล้ว'; // ข้อความในเนื้อหาอีเมลที่ตอบกลับ

// การแมพข้อมูลคอลัมน์ของ Google Sheet 
let index_col = {
  "ประทับเวลา": 0, "อีเมล": 1, "ครูนิเทศ": 2, "ตำแหน่งครู": 3, "ระดับการศึกษา": 4,
  "แผนกวิชา": 5, "ชื่อนักเรียน นักศึกษา": 6, "จำนวนนักเรียน นักศึกษา": 7, "ประจำสัปดาห์ที่": 8, "วันที่ทำการออกนิเทศก์": 9,
  "ชื่อสถานประกอบการ": 10, "ชื่อผู้ประสานงาน/ครูฝึก": 11, "ตำแหน่ง ผู้ประสานงาน/ครูฝึก": 12, "รูปแบบการนิเทศ": 13, "งานที่ปฏิบัติ": 14,
  "ประเมินผลการปฏิบัติงานของนักศึกษาฝึกอาชีพในสถานประกอบการ [ตรงต่อเวลา/ขยันอดทน]": 15,
  "ประเมินผลการปฏิบัติงานของนักศึกษาฝึกอาชีพในสถานประกอบการ [การปฏิบัติตามคำสั่ง คำแนะนำ ของครูฝึก]": 16,
  "ประเมินผลการปฏิบัติงานของนักศึกษาฝึกอาชีพในสถานประกอบการ [การปฏิบัติตามกฎระเบียบของสถานประกอบการ]": 17,
  "ประเมินผลการปฏิบัติงานของนักศึกษาฝึกอาชีพในสถานประกอบการ [กิริยามารยาทสภาพเรียบร้อย]": 18,
  "ประเมินผลการปฏิบัติงานของนักศึกษาฝึกอาชีพในสถานประกอบการ [บำรุงรักษาเครื่องมือเครื่องใช้ และทรัพย์สินขององค์กร]": 19,
  "ประเมินผลการปฏิบัติงานของนักศึกษาฝึกอาชีพในสถานประกอบการ [เรียนรู้และพัฒนาตนเองอยู่เสมอ]": 20,
  "นักศึกษาฝึกอาชีพมีปัญหาหรือไม่": 21, "ปัญหาที่พบ": 22, "การแก้ปัญหา": 23,
  "รูปภาพการนิเทศ1": 24, "รูปภาพการนิเทศ2": 25, "ข้อเสนอแนะ": 26,
  "create_pdf_status": 27, "send_email_status": 28
};


let colEmail = index_col["อีเมล"]; // ตัวแปรนี้ใช้เพื่อระบุคอลัมน์ที่เก็บอีเมลใน Google Sheets (ดึงมาจาก index_col ซึ่งแมปคอลัมน์ทั้งหมดของ Google Sheets)
let colEmailStatus = index_col["send_email_status"] || ""; // ใช้ระบุคอลัมน์ที่เก็บสถานะการส่งอีเมล (send_email_status) หากคอลัมน์ไม่มีข้อมูลจะกำหนดค่าเป็น ""
let colLineStatus = index_col["send_line_status"] || ""; // ระบุคอลัมน์ที่เก็บสถานะการส่งข้อความ LINE Notify (send_line_status)
let colEmailStatusName = "AC"; // ชื่อคอลัมน์ใน Google Sheets ที่ใช้เก็บสถานะการส่งอีเมล (กำหนดไว้ว่าเป็น AC)
let colLineStatusName = "O"; // ชื่อคอลัมน์ใน Google Sheets ที่ใช้เก็บสถานะการส่งข้อความ LINE Notify (กำหนดไว้ว่าเป็น O)


let colName = index_col["ชื่อสถานประกอบการ"] + "_" + index_col["ครูนิเทศ"] || "";; // คอลัมน์ที่จะนำมาตั้งเป็นชื่อไฟล์ pdf
let colPdfStatus = index_col["create_pdf_status"] || ""; // ระบุคอลัมน์ที่เก็บสถานะการสร้าง PDF (create_pdf_status)
let colPdfStatusName = "AB"; // ชื่อคอลัมน์ใน Google Sheets ที่เก็บสถานะการสร้าง PDF (กำหนดไว้ว่าเป็น AB)

// กำหนดข้อมูลคอลัมน์ที่เกี่ยวข้องกับรูปภาพ โดยใช้ index_col เพื่อแมปชื่อรูปภาพ 
// เช่น รูปภาพการนิเทศ1 และ รูปภาพการนิเทศ2 เข้ากับ Placeholder ที่จะใช้ใน 
// Google Slides ({{รูปภาพการนิเทศ1}}, {{รูปภาพการนิเทศ2}})
let colAllImage = [
  { [index_col['รูปภาพการนิเทศ1']]: '{{รูปภาพการนิเทศ1}}' },
  { [index_col['รูปภาพการนิเทศ2']]: '{{รูปภาพการนิเทศ2}}' },
];

// ใช้สำหรับเก็บข้อมูลคอลัมน์ที่เกี่ยวข้องกับ Checkboxes ซึ่งยังไม่ได้กำหนดในโค้ดนี้ (ตัวอย่างถูกคอมเมนต์ไว้)
let index_col_checkbox = [];
// let index_col_checkbox = [
//   { [index_col['เพศ']] : [{'ชาย':'H'},{'หญิง':'I'}] },
//   ];


// ใช้สำหรับจัดการคอลัมน์ที่มี Checkboxes หลายตัวใน 1 คอลัมน์ (ตัวอย่างถูกคอมเมนต์ไว้)
let index_col_multi_checkbox = [];
// let index_col_multi_checkbox = [
//   { [index_col['ทักษะ [การพูด]']] : [{'ไทย':'J'},{'อังกฤษ':'K'}] },
//   { [index_col['ทักษะ [การฟัง]']] : [{'ไทย':'L'},{'อังกฤษ':'M'}] },
//   ];  
let tokensV2 = ['']; // เก็บ Token ที่ใช้สำหรับส่งข้อความผ่าน LINE Notify (ปัจจุบันยังว่าง)


// Init
let newSlideName = 'New_FormToSlidePDF_'; // ชื่อ Template ของไฟล์ Google Slides ที่จะสร้างใหม่
let sent_status = 'SENT'; // ค่า "SENT" ใช้เพื่อตรวจสอบสถานะว่าการส่งอีเมลหรือการสร้าง PDF เสร็จสมบูรณ์แล้ว
let ss, sheet, lastRow, lastCol, range, values;
let data_name;
let newSlide, newSlideId, presentation, all_shape;
let titleName;
let exportPdf, pdf_name_full;
let email_send = [];
let filePath;
let isPdfOnly = false;


function formToSlidePdfLine() {
  /*######################### Editable2 Start #########################*/
  function formatMsgToLine() {
    return ฟอร์มการประเมินการนิเทศฝึกงาน


${ filePath };
  }
  /*#########################  Editable2 End  #########################*/
  initSpreadSheet().then(async function () {
    formatTitle();
    if (isSendEmail == false && isSendLine == false) {
      isPdfOnly = true;
    }
    for (let i = 1; i < lastRow; i++) {
      let numRow = i + 1;
      clearVal();
      let cur_data = values[i];
      data_name = cur_data[colName];
      try {
        data_name = data_name.replace(/\s/g, '');
      } catch (e) { }


      if (isPdfOnly) {
        let status = cur_data[colPdfStatus];
        if (status == sent_status) {
          continue;
        }
      } else {
        let emailStatus = cur_data[colEmailStatus];
        let lineStatus = cur_data[colLineStatus];
        if ((!isSendEmail || (isSendEmail && emailStatus == sent_status)) && (!isSendLine || (isSendLine && lineStatus == sent_status))) {
          continue;
        }
      }
      await duplicateSlide().then(async function () {
        await updateCheckboxCol(cur_data, numRow).then(async function () {
          values = range.getValues();
          cur_data = values[i];
          await updateSlideData(cur_data).then(async function () {
            presentation.saveAndClose();
            await createPdf().then(async function () {
              removeTempSlide();
              if (!isPdfOnly) {
                let cur_email = cur_data[colEmail];
                let emailStatus = cur_data[colEmailStatus];
                let lineStatus = cur_data[colLineStatus];
                if (validateEmail(cur_email)) {
                  email_send.push(cur_email);
                }
                console.log(email_send);
                if (isSendEmail && emailStatus != sent_status) {
                  for (let j = 0; j < email_send.length; j++) {
                    if (validateEmail(email_send[j])) {
                      await sendEmailWithAttachment(email_send[j]).then(function () {
                        if (j == email_send.length - 1) {
                          updateStatusSent(numRow, 'email');
                        }
                      });
                    }
                  }
                }
                if (isSendLine && lineStatus != sent_status) {
                  sendLineNotify(formatMsgToLine());
                  updateStatusSent(numRow, 'line');
                }
              } else {
                updateStatusSent(numRow, 'pdf');
              }
            });
          });
        });
      });
    }
    console.log('Program completed');
  });
}


async function formatDateToSlide(date) {
  // แปลงวันที่เป็นรูปแบบ dd mmmm yyyy เช่น 17 October 2024
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd MMMM yyyy');
}


async function sendLineNotify(all_message_send) {
  return new Promise(function (resolve) {
    for (let k = 0; k < tokensV2.length; k++) {
      let formData = {
        'message': all_message_send,
      }
      let options = {
        "method": "post",
        "payload": formData,
        "headers": { "Authorization": "Bearer " + tokensV2[k] }
      };
      UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
    }


    resolve();
    console.log('sendLineNotify completed');
  });
}


function clearVal() {
  data_name = '';
  newSlide = newSlideId = presentation = '';
  exportPdf = pdf_name_full = '';
  email_send = isSendEmail ? [...email_send_default] : [];
  all_shape = '';
  console.log('clearVal completed');
}


async function initSpreadSheet() {
  return new Promise(function (resolve) {
    ss = SpreadsheetApp.getActive();
    sheet = ss.getSheetByName(sheetName);
    lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    range = sheet.getDataRange();
    values = range.getValues();
    resolve();
    console.log('initSpreadSheet completed');
  });
}


function formatTitle() {
  titleName = values[0];
  titleName.forEach(function (item, index) {
    titleName[index] = '{{' + item + '}}';
  });
  console.log('formatTitle completed');
}


async function duplicateSlide() {
  return new Promise(function (resolve) {
    let templateSlide = DriveApp.getFileById(templateSlideId);
    let templateResponseFolder = DriveApp.getFolderById(folderResponseSlideId);
    newSlide = templateSlide.makeCopy(newSlideName.concat(data_name), templateResponseFolder);
    resolve();
    console.log('duplicateSlide completed');
  });
}


async function updateCheckboxCol(cur_data, numRow) {
  return new Promise(function (resolve) {
    index_col_checkbox.forEach(function (item) {
      Object.keys(item).forEach(function (key) {
        var cur_checkbox_val = cur_data[key];
        item[key].forEach(function (item_ele) {
          Object.keys(item_ele).forEach(function (key_item_ele) {
            if (key_item_ele === cur_checkbox_val) {
              sheet.getRange(item_ele[key_item_ele].concat(numRow)).setValue('✓');
            }
          })
        })
      });
    });
    index_col_multi_checkbox.forEach(function (item) {
      Object.keys(item).forEach(function (key) {
        let cur_multi_checkbox_val = cur_data[key];
        let cur_multi_checkbox_val_arr = cur_multi_checkbox_val.split(", ");
        for (let i = 0; i < cur_multi_checkbox_val_arr.length; i++) {
          let cur_val = cur_multi_checkbox_val_arr[i];
          item[key].forEach(function (item_ele) {
            Object.keys(item_ele).forEach(function (key_item_ele) {
              if (key_item_ele === cur_val) {
                sheet.getRange(item_ele[key_item_ele].concat(numRow)).setValue('✓');
              }
            })
          })
        }
      });
    });
    resolve();
    console.log('updateCheckboxCol completed');
  });
}


async function updateSlideData(cur_data) {
  return new Promise(function (resolve) {
    // Init
    newSlideId = newSlide.getId();
    presentation = SlidesApp.openById(newSlideId);
    let slide = presentation.getSlides()[0];
    all_shape = slide.getShapes();
    titleName.forEach(async function (item, index) {
      let isColImg = false;
      colAllImage.forEach(async function (img_item) {
        Object.keys(img_item).forEach(async function (key) {
          if (item === img_item[key]) {
            all_shape.forEach(async function (s) {
              if (s.getText().asString().includes(img_item[key])) {
                let cur_img_url = cur_data[key];
                let imageFileId = getIdFromUrl(cur_img_url)
                if (imageFileId) {
                  isColImg = true;
                  let image = DriveApp.getFileById(imageFileId).getBlob();
                  await replaceImage(s, image).then(async function () {
                    console.log('replaceImage completed')
                  });
                }
              }
            });
          }
        });
      });
      if (!isColImg) {
        let templateVariable = item;
        let replaceValue = cur_data[index];
        presentation.replaceAllText(templateVariable, replaceValue);


        if (templateVariable.includes('วันที่ทำการออกนิเทศก์')) {
          replaceValue = formatDateToSlide(new Date(replaceValue));
        }


        presentation.replaceAllText(templateVariable, replaceValue);
      }
    })
    resolve();
    console.log('updateSlideData completed');
  });
}


async function replaceImage(s, image) {
  let res;
  return new Promise(function (resolve) {
    try {
      res = s.replaceWithImage(image);
    } catch (e) {
      console.log("error:", e)
    }
    if (res) {
      console.log('resolve');
      resolve();
    }
  })
}
async function createPdf() {
  return new Promise(function (resolve, reject) {
    let pdf = DriveApp.getFileById(newSlideId).getBlob().getAs("application/pdf");
    pdf_name_full = pdf_file_name + data_name + '.pdf';
    pdf.setName(pdf_name_full);
    exportPdf = DriveApp.getFolderById(folderResponsePdfId).createFile(pdf);
    filePath = exportPdf.getUrl();
    if (exportPdf) {
      resolve();
      console.log('Create PDF completed');
    } else {
      reject();
      console.log('Create PDF error');
    }
  });
}


async function sendEmailWithAttachment(email) {
  return new Promise(function (resolve, reject) {
    let file = DriveApp.getFolderById(folderResponsePdfId).getFilesByName(pdf_name_full);
    if (!file.hasNext()) {
      console.error("Could not open file " + pdf_name_full);
      return;
    }
    try {
      MailApp.sendEmail({
        to: email,
        subject: email_subject,
        htmlBody: email_message,
        attachments: [file.next().getAs(MimeType.PDF)]
      });
      resolve();
      console.log('sendEmailWithAttachment completed')
    } catch (e) {
      reject();
      console.log("sendEmailWithAttachment error with email (" + email + "). " + e);
    }
  });
}


function removeTempSlide() {
  try {
    DriveApp.getFileById(newSlideId).setTrashed(true);
    console.log('removeTempSlide completed');
  } catch (e) {
    console.log('removeTempSlide error')
  }
}


function updateStatusSent(numRow, mode) {
  if (mode == 'email') {
    sheet.getRange(colEmailStatusName.concat(numRow)).setValue(sent_status);
  } else if (mode == 'line') {
    sheet.getRange(colLineStatusName.concat(numRow)).setValue(sent_status);
  } else if (mode == 'both') {
    sheet.getRange(colEmailStatusName.concat(numRow)).setValue(sent_status);
    sheet.getRange(colLineStatusName.concat(numRow)).setValue(sent_status);
  } else if (mode == 'pdf') {
    sheet.getRange(colPdfStatusName.concat(numRow)).setValue(sent_status);
  }
  console.log('updateStatusSent completed');
}


function formatUrlImg(url) {
  let new_url = '';
  let start_url = 'https://drive.google.com/uc?id=';
  new_url = start_url + getIdFromUrl(url);
  return new_url;
}


function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}


function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  if (!re.test(email)) {
    return false;
  } else {
    return true;
  }
}


function generateTitle() {
  let result = "";
  initSpreadSheet();
  let title = values[0];
  console.log("title:", title);
  for (let i = 0; i < title.length; i++) {
    result += "${title[i]}":${ i };
    if (i != title.length - 1) {
      result += ','
    } else {
      result = let index_col = { ${ result }
    };;
  }
}
console.log(result);
}


function customFormatDate(date, mode, format) {
  let _timezone = "";
  if (mode == "date") {
    _timezone = "GMT+7";
  } else if (mode == "time") {
    _timezone = "GMT+6:43";
  } else {
    _timezone = timezone;
  }
  return Utilities.formatDate(date, _timezone, format);
}
