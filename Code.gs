function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Deen Connect - Video System v3.2(Normal)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getLatestMeetingStatusInfo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings");
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  
  for (var i = data.length - 1; i >= 1; i--) {
    var status = String(data[i][4]).trim();
    if (status === "WAITING" || status === "ACTIVE") {
      return {
        room: String(data[i][0]).trim(),
        secretRoom: String(data[i][1]).trim(),
        pass: String(data[i][3]).trim(),
        status: status
      };
    }
  }
  return null;
}

function checkHostLogin(u, p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hosts");
  if (!sheet) return false;
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) return true;
  }
  return false;
}

function checkOwnerLogin(u, p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Owner");
  if (!sheet) return false;
  
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) return (String(data[1][0]).trim() === String(u).trim() && String(data[1][1]).trim() === String(p).trim());
  return false;
}

function updateOwnerCredentials(currentU, currentP, newU, newP) {
  if (!checkOwnerLogin(currentU, currentP)) {
    return { success: false, msg: "বর্তমান Owner ইউজারনেম বা পাসওয়ার্ড ভুল!" };
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Owner");
  if (!sheet) return { success: false, msg: "Owner Sheet পাওয়া যায়নি!" };

  sheet.getRange(2, 1).setValue(String(newU).trim());
  sheet.getRange(2, 2).setValue(String(newP).trim());

  return { success: true, msg: "Owner তথ্য সফলভাবে পরিবর্তন করা হয়েছে!" };
}

function createMeeting(roomName, pass, hostUser, autoStart) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Meetings");
    sheet.appendRow(["Room", "SecretRoom", "Host", "Password", "Status", "CreatedAt"]);
  }
  
  var secretRoom = "DeenConnect_" + Math.random().toString(36).substring(2, 9);
  var status = autoStart ? "ACTIVE" : "WAITING";
  var now = new Date();
  
  sheet.appendRow([String(roomName).trim(), secretRoom, String(hostUser).trim(), String(pass).trim(), status, now]);
  
  return { success: true, secretRoom: secretRoom };
}

function getActiveMeetingByName(roomName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings");
  if (!sheet) return { status: "INACTIVE" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var sheetRoom = String(data[i][0]).trim().toLowerCase();
    var inputRoom = String(roomName).trim().toLowerCase();
    var status = String(data[i][4]).trim();
    
    if (sheetRoom === inputRoom && (status === "ACTIVE" || status === "WAITING")) {
      return {
        room: String(data[i][0]).trim(),
        secretRoom: String(data[i][1]).trim(),
        pass: String(data[i][3]).trim(),
        status: status
      };
    }
  }
  return { status: "INACTIVE" };
}

function hostStartMeeting(roomName, u, p) {
  if (!checkHostLogin(u, p)) {
    return { success: false, msg: "ভুল হোস্ট লগইন!" };
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings");
  var data = sheet.getDataRange().getValues();
  
  for (var i = data.length - 1; i >= 1; i--) {
    var sheetRoom = String(data[i][0]).trim().toLowerCase();
    var inputRoom = String(roomName).trim().toLowerCase();
    var status = String(data[i][4]).trim();
    
    if (sheetRoom === inputRoom && (status === "WAITING" || status === "ACTIVE")) {
      sheet.getRange(i + 1, 5).setValue("ACTIVE");
      return { success: true, secretRoom: String(data[i][1]).trim() };
    }
  }
  return { success: false, msg: "সক্রিয় মিটিং পাওয়া যায়নি" };
}

function getOwnerFullMeetingDetails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings");
  if (!sheet) return null;
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;
  
  for (var i = data.length - 1; i >= 1; i--) {
    var status = String(data[i][4]).trim();
    if (status === "ACTIVE" || status === "WAITING") {
      var startTime = new Date(data[i][5]);
      var now = new Date();
      var durationMinutes = Math.floor((now - startTime) / (1000 * 60));

      return {
        rowIndex: i + 1,
        roomName: String(data[i][0]).trim(),
        secretRoom: String(data[i][1]).trim(),
        hostName: String(data[i][2]).trim(),
        password: String(data[i][3]).trim() || "None",
        status: status,
        isHostPresent: (status === "ACTIVE"),
        createdAt: startTime.toLocaleTimeString(),
        duration: durationMinutes + " minute(s)"
      };
    }
  }
  return null;
}

function deleteMeetingByOwner() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meetings");
  if (!sheet) return { success: false, msg: "Sheet পাওয়া যায়নি" };
  
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = data.length - 1; i >= 1; i--) {
    var status = String(data[i][4]).trim();
    if (status === "ACTIVE" || status === "WAITING") {
      sheet.getRange(i + 1, 5).setValue("INACTIVE");
      found = true;
    }
  }
  if (found) return { success: true, msg: "মিটিং সফলভাবে বন্ধ/ডিলিট করা হয়েছে!" };
  return { success: false, msg: "কোন সক্রিয় মিটিং পাওয়া যায়নি" };
}

function getHostList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hosts");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push({ rowIndex: i + 1, username: String(data[i][0]).trim() });
    }
  }
  return list;
}

function addHostByOwner(u, p) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hosts");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Hosts");
    sheet.appendRow(["Username", "Password"]);
  }
  sheet.appendRow([String(u).trim(), String(p).trim()]);
  return { success: true, msg: "নতুন হোস্ট সফলভাবে যোগ করা হয়েছে!" };
}

function deleteHostByOwner(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hosts");
  if (!sheet) return { success: false, msg: "Sheet পাওয়া যায়নি" };
  
  sheet.deleteRow(rowIndex);
  return { success: true, msg: "হোস্ট সফলভাবে রিমুভ করা হয়েছে!" };
}
