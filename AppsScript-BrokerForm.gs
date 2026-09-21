/**
 * Dark Star Dispatch — Broker Form Backend
 * -----------------------------------------
 * Ye script Google Sheet ke "Extensions > Apps Script" mein paste karni hai
 * (usi Google Sheet mein jisme "Loads" tab hai — Dark Star Dispatch — Data).
 *
 * Ye kya karti hai:
 *  1) Web form se POST request receive karti hai
 *  2) "Loads" tab ke header row se column match karti hai (jo bhi order/naam ho)
 *  3) Agar koi field ka column nahi milta, khud naya column bana deti hai
 *  4) Ek Load ID auto-generate karti hai (format: DS-YYYYMMDD-NNN)
 *  5) Naya row Sheet mein add (append) kar deti hai
 *
 * SETUP STEPS:
 *  1) Google Sheet kholein (Dark Star Dispatch — Data)
 *  2) Extensions > Apps Script
 *  3) Is file ka poora code paste karein (purana Code.gs content replace kar dein)
 *  4) Upar "Deploy" > "New deployment" > type: "Web app"
 *  5) Execute as: "Me"   |   Who has access: "Anyone"
 *  6) Deploy karein, jo URL mile wo copy karein
 *  7) Us URL ko index.html mein ENDPOINT_URL wali line mein paste karein
 */

const SHEET_NAME = "Loads";           // Sheet tab ka naam — apni Sheet ke mutabik badal lein agar different ho
const DEFAULT_STATUS = "New";         // Naye submissions ko ye default Status milega

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonResponse({ success: false, error: "Sheet tab '" + SHEET_NAME + "' nahi mila." });
    }

    // Header row padhna (row 1)
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Har field ko sahi column mein map karna — naam se match, case-insensitive
    const fieldToColumn = {
      "Load ID": generateLoadId(sheet, headers),
      "Broker": data.broker || "",
      "Origin": data.origin || "",
      "Destination": data.destination || "",
      "Rate": data.rate || "",
      "Pickup DateTime": data.pickup || "",
      "Delivery DateTime": "",                          // form se nahi aata, khali rahega
      "Payload Weight": data.weight || "",
      "Vehicle Type": data.vehicle || "",
      "Tailgate Required": data.tailgate || "",
      "Assigned Driver": "",                             // dispatcher baad mein assign karega
      "Assigned Truck": "",                               // dispatcher baad mein assign karega
      "Status": DEFAULT_STATUS,
      "Contact Name": data.contact || "",
      "Contact Email": data.email || "",
      "Contact Phone": data.phone || "",
      "MC/DOT Number": data.mc || "",
      "Notes": data.notes || ""
    };

    // Agar koi header missing hai to naya column add karo
    Object.keys(fieldToColumn).forEach(function (fieldName) {
      if (headers.indexOf(fieldName) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(fieldName);
        headers.push(fieldName);
      }
    });

    // Naya row banayein, headers ke sahi order mein values bharein
    const newRow = headers.map(function (h) {
      return fieldToColumn.hasOwnProperty(h) ? fieldToColumn[h] : "";
    });

    sheet.appendRow(newRow);

    return jsonResponse({ success: true, loadId: fieldToColumn["Load ID"] });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function generateLoadId(sheet, headers) {
  const today = new Date();
  const y = today.getFullYear();
  const m = ("0" + (today.getMonth() + 1)).slice(-2);
  const d = ("0" + today.getDate()).slice(-2);
  const datePart = "" + y + m + d;

  const idColIndex = headers.indexOf("Load ID");
  let countToday = 0;

  if (idColIndex !== -1 && sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, idColIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(function (row) {
      if (row[0] && row[0].toString().indexOf("DS-" + datePart) === 0) {
        countToday++;
      }
    });
  }

  const seq = ("00" + (countToday + 1)).slice(-3);
  return "DS-" + datePart + "-" + seq;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Optional: is function ko manually run kar ke test kar sakte hain
 * (Apps Script editor mein "doGet" select kar ke Run dabayein)
 */
function doGet() {
  return ContentService.createTextOutput("Dark Star Dispatch broker-form endpoint is live.");
}
