var TELEGRAM_TOKEN = '8481633989:AAEVeyE1n_LRkxeOfVqwsVKjoJR3X63kQOY';
var SHEET_USERS = 'Users';
var SHEET_LOGS = 'Logs';
var SESSION_PROP = PropertiesService.getScriptProperties();

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var chat_id = data.message.chat.id;
    var text = data.message.text.trim();

    if (text == "/start" || text == "/help") {
      sendMessage(chat_id, "Commands:\n/register\n/log\n/status\n/delelast");
      return;
    }

    if (text == "/register") {
      sendMessage(chat_id, "Please enter your Name and IC last 4 digits separated by a comma. Example:\nJohn,1234");
      setSession(chat_id, {step: "register"});
      return;
    }

    if (text == "/log") {
      if (!isUserRegistered(chat_id)) {
        sendMessage(chat_id, "You must /register first.");
        return;
      }
      sendMessage(chat_id, "Enter drive type: DCP/POL/OTHERS");
      setSession(chat_id, {step: "drive_type"});
      return;
    }

    if (text == "/status") {
      if (!isUserRegistered(chat_id)) {
        sendMessage(chat_id, "You must /register first.");
        return;
      }
      sendStatus(chat_id);
      return;
    }

    if (text == "/delelast") {
      if (!isUserRegistered(chat_id)) {
        sendMessage(chat_id, "You must /register first.");
        return;
      }
      deleteLastLog(chat_id);
      return;
    }

    // Handle session steps
    var session = getSession(chat_id);
    if (!session) {
      sendMessage(chat_id, "Invalid command. Type /help for options.");
      return;
    }

    switch(session.step) {
      case "register":
        handleRegister(chat_id, text);
        break;
      case "drive_type":
        handleDriveType(chat_id, text);
        break;
      case "drive_details":
        handleDriveDetails(chat_id, text);
        break;
      default:
        sendMessage(chat_id, "Invalid state. Type /help for options.");
    }

  } catch(err) {
    Logger.log(err);
    sendMessage(chat_id, "Error occurred. Please try again.");
  }
}

/* ------------------ SESSION FUNCTIONS ------------------ */
function setSession(chat_id, data) {
  SESSION_PROP.setProperty(chat_id.toString(), JSON.stringify(data));
}
function getSession(chat_id) {
  var val = SESSION_PROP.getProperty(chat_id.toString());
  return val ? JSON.parse(val) : null;
}
function clearSession(chat_id) {
  SESSION_PROP.deleteProperty(chat_id.toString());
}

/* ------------------ USER FUNCTIONS ------------------ */
function isUserRegistered(chat_id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == chat_id) return true;
  }
  return false;
}
function getNameByChatId(chatId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = ss.getDataRange().getValues(); // 2D array

  for (var i = 1; i < data.length; i++) { // start at 1 to skip header
    if (data[i][0] == chatId) {  // column A = index 0
      return data[i][1]; // return name (column B)
    }
  }

  return null; // no user found
}

function handleRegister(chat_id, text) {
  var parts = text.split(",");
  if (parts.length != 2) {
    sendMessage(chat_id, "Invalid format. Example: John,1234");
    return;
  }
  var name = parts[0].trim();
  var ic = parts[1].trim();
  if (!/^\d{4}$/.test(ic)) {
    sendMessage(chat_id, "IC last 4 digits must be numeric.");
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  sheet.appendRow([chat_id, name, ic, 0, ""]); // cumulative mileage=0, lastHA=""
  sendMessage(chat_id, "Registered successfully!");
  clearSession(chat_id);
}

/* ------------------ DRIVE LOG FUNCTIONS ------------------ */
function handleDriveType(chat_id, text) {
  text = text.toUpperCase();
  if (text != "DCP" && text != "POL" && text != "OTHERS") {
    sendMessage(chat_id, "Invalid drive type. Enter DCP or POL or OTHERS");
    return;
  }
  var session = {step:"drive_details", driveType:text};
  setSession(chat_id, session);
  if (text == "DCP") {
    sendMessage(chat_id, "Enter details: Date(YYYY-MM-DD),odoStart,odoEnd,VehicleNo,VehicleClass\nExample: 2025-11-16,1000,1050,36589,LUV");
  } 
  else if (text == "OTHERS") {
    sendMessage(chat_id, "Enter details: Date(YYYY-MM-DD),odoStart,odoEnd,VehicleNo,VehicleClass,reason\nExample: 2025-11-16,1000,1050,36589,LUV,VSwap");
  } 
  else {
    sendMessage(chat_id, "Enter details: Date(YYYY-MM-DD),odoStart,odoMid,odoEnd,VehicleNo,VehicleClass,FuelFilled\nExample: 2025-11-16,1000,1025,1050,36589,LUV,50");
  }
}

function handleDriveDetails(chat_id, text) {
  var session = getSession(chat_id);
  if (!session) {
    sendMessage(chat_id, "Session expired. Please start again.");
    return;
  }

  var type = session.driveType;
  var parts = text.split(",");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOGS);
  var name = getNameByChatId(chat_id);
  var timestamp = new Date();

  // Define drive schemas
  const schemas = {
    "DCP": ["date","odoStart","odoEnd","vehicleNo","vehicleClass"],
    "POL": ["date","odoStart","odoMid","odoEnd","vehicleNo","vehicleClass","fuelFilled"],
    "OTHERS": ["date","odoStart","odoEnd","vehicleNo","vehicleClass","reason"]
  };

  if (!schemas[type] || parts.length != schemas[type].length) {
    sendMessage(chat_id, `Invalid format for ${type}. Please follow the instructions.`);
    return;
  }

  // Parse fields
  let data = {};
  schemas[type].forEach((field, i) => {
    data[field] = parts[i].trim();
  });

  // Convert numeric fields
  if (data.odoStart) data.odoStart = Number(data.odoStart);
  if (data.odoEnd) data.odoEnd = Number(data.odoEnd);
  if (data.odoMid) data.odoMid = Number(data.odoMid);
  if (data.fuelFilled) data.fuelFilled = Number(data.fuelFilled);

  // Validate odometer readings
  if (isNaN(data.odoStart) || isNaN(data.odoEnd) || data.odoEnd < data.odoStart) {
    sendMessage(chat_id, "Invalid odometer readings.");
    return;
  }
  if (type == "POL" && (isNaN(data.odoMid) || data.odoMid < data.odoStart || data.odoMid > data.odoEnd || isNaN(data.fuelFilled))) {
    sendMessage(chat_id, "Invalid odometer readings or fuel.");
    return;
  }

  // Calculate mileage
  let mileage = data.odoEnd - data.odoStart;

  // Prepare row for logging
  let row;
  if (type == "DCP") {
    row = [chat_id, timestamp, data.date, type, data.odoStart, "", data.odoEnd, data.vehicleNo, data.vehicleClass, "", mileage, name];
  } else if (type == "POL") {
    row = [chat_id, timestamp, data.date, type, data.odoStart, data.odoMid, data.odoEnd, data.vehicleNo, data.vehicleClass, data.fuelFilled, mileage, name];
  } else { // OTHERS
    row = [chat_id, timestamp, data.date, data.reason, data.odoStart, "", data.odoEnd, data.vehicleNo, data.vehicleClass, "", mileage, name];
  }

  // Append to sheet
  sheet.appendRow(row);

  // Update cumulative mileage & HA
  updateCumulativeMileage(chat_id, mileage);
  updateDriverHA(chat_id);

  // Calculate HA & JIT messages
  const ha = calculateHA(chat_id);
  const msgHA =
    `🪪 Driver HA Status\n` +
    `📅 HA Expiry: ${ha.haExpiry}\n` +
    `🚗 Remaining KM to reset cycle: ${ha.remaining} km`;

  sendMessage(chat_id, `Drive logged. Mileage for this drive: ${mileage} km\nCumulative mileage: ${getCumulativeMileage(chat_id)} km\n${msgHA}`);

  const jit = checkJITExpiry(chat_id, data.vehicleClass);
  if (jit.jitRequired) {
    sendMessage(chat_id, `⚠️ JIT required for ${data.vehicleClass}\nLast driven: ${jit.lastDriven || "Never. Please do JIT before your drive."}`);
  } else {
    const lastDriven = new Date(data.date);
    lastDriven.setDate(lastDriven.getDate() + 9);
    const formattedDate = Utilities.formatDate(lastDriven, Session.getScriptTimeZone(), "EEE dd MMM yyyy");
    sendMessage(chat_id, `⚠️ JIT for ${data.vehicleClass} expires at: ${formattedDate}`);
  }

  clearSession(chat_id);
}

/* ------------------ CUMULATIVE MILEAGE ------------------ */
function getCumulativeMileage(chat_id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == chat_id) return data[i][3];
  }
  return 0;
}

function updateCumulativeMileage(chat_id, added) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == chat_id) {
      var newTotal = Number(data[i][3]) + added;
      sheet.getRange(i+1,4).setValue(newTotal);
      return;
    }
  }
}


/* ------------------ DRIVER HA and JIT CALCULATION ------------------ */
function calculateHA(chat_id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LOGS");
  const data = sheet.getDataRange().getValues().slice(1); // skip header

  // 1️⃣ Filter logs for the user and map date + mileage
  const userLogs = data
    .filter(r => String(r[0]) === String(chat_id))
    .map(r => ({
      date: new Date(r[2]),   // DATE column
      mileage: Number(r[10])  // MILEAGE column
    }));

  if (userLogs.length === 0) {
    return { haExpiry: "", remaining: 50 };
  }

  // 2️⃣ Sort logs ascending by date
  userLogs.sort((a, b) => a.date - b.date);

  let accumulated = 0;       // track partial mileage toward next cycle
  let haExpiry = null;       // last completed cycle date + 3 months

  for (let i = 0; i < userLogs.length; i++) {
    accumulated += userLogs[i].mileage;

    if (accumulated >= 50) {
      // Cycle completed
      haExpiry = new Date(userLogs[i].date);
      haExpiry.setMonth(haExpiry.getMonth() + 3);

      // Only subtract 50 for next cycle tracking (extra over 50 counts toward next)
      accumulated -= 50;
    }
  }

  // Remaining km to complete next cycle
  const remainingKm = Math.max(50 - accumulated, 0);

  return {
    haExpiry: haExpiry
      ? Utilities.formatDate(haExpiry, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : "",
    remaining: remainingKm
  };
}

/* ------------------ UPDATE USERS SHEET ------------------ */
function updateDriverHA(chat_id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USERS");
  const data = sheet.getDataRange().getValues();
  const ha = calculateHA(chat_id);

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(chat_id)) {
      // Column E = HA expiry, Column F = remaining KM
      sheet.getRange(i + 1, 5).setValue(ha.haExpiry);
      sheet.getRange(i + 1, 6).setValue(ha.remaining);
      return;
    }
  }
}

function checkJITExpiry(chat_id, vehicleClass) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_LOGS);

  const data = sheet.getDataRange().getValues().slice(1);

  // 1️⃣ Filter logs for user + vehicle class
  const logs = data
    .filter(r => r[0] == chat_id && r[8] == vehicleClass)
    .map(r => new Date(r[2]));

  // 2️⃣ If never driven this vehicle → JIT needed
  if (logs.length === 0) {
    return {
      jitRequired: true,
      reason: "Never driven this vehicle class"
    };
  }

  // 3️⃣ Find most recent drive
  logs.sort((a, b) => b - a);
  const lastDrivenDate = logs[0];

  // 4️⃣ Calculate days since last drive
  const today = new Date();
  const diffMs = today - lastDrivenDate;
  const diffDays = diffMs / (1000 * 60 * 60 * 24);

  // 5️⃣ Check JIT rule
  if (diffDays > 10) {
    return {
      jitRequired: true,
      lastDriven: Utilities.formatDate(
        lastDrivenDate,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd"
      )
    };
  }

  // 6️⃣ Still valid
  return {
    jitRequired: false,
    daysRemaining: Math.floor(10 - diffDays)
  };
}



/* ------------------ DELETE LAST LOG ------------------ */
function deleteLastLog(chat_id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOGS);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length-1; i>0; i--) {
    if (data[i][0] == chat_id) {
      var mileage = Number(data[i][10]);
      sheet.deleteRow(i+1);
      // Update cumulative mileage
      updateCumulativeMileage(chat_id, -mileage);
      sendMessage(chat_id, `Last log deleted. Updated cumulative mileage: ${getCumulativeMileage(chat_id)} km`);
      return;
    }
  }
  sendMessage(chat_id, "No logs to delete.");
}

/* ------------------ STATUS ------------------ */
function sendStatus(chat_id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USERS);
  const data = sheet.getDataRange().getValues().slice(1); // skip header

  // Find the driver row
  let haExpiry = "";
  let remainingKm = "";
  let driverFound = false;

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(chat_id)) {
      // Assuming Column E = HA expiry, Column F = remaining KM
      haExpiry = data[i][4];   // E = 5th column
      remainingKm = data[i][5]; // F = 6th column
      driverFound = true;
      break;
    }
  }

  if (!driverFound) {
    sendMessage(chat_id, `No HA status found for this driver.`);
    return;
  }

  // Construct message
  const msg = 
    `🪪 Driver HA Status\n` +
    `📅 HA Expiry: ${haExpiry}\n` +
    `🚗 Remaining KM to reset cycle: ${remainingKm} km`;

  sendMessage(chat_id, msg);
}

/* ------------------ TELEGRAM SEND MESSAGE ------------------ */
function sendMessage(chat_id, text) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  var payload = { chat_id: chat_id, text: text };
  var options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload) };
  UrlFetchApp.fetch(url, options);
}
