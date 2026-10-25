function runDataPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ดึงข้อมูลจากแท็บต่างๆ
  var rawStockSheet = ss.getSheetByName("Raw_Stock");
  var productsSheet = ss.getSheetByName("Products");
  var supplierSheet = ss.getSheetByName("Supplier_SLA");
  var cleanedSheet = ss.getSheetByName("Cleaned_Data");
  var quarantineSheet = ss.getSheetByName("Quarantine");
  
  var stockData = rawStockSheet.getDataRange().getValues();
  var productData = productsSheet.getDataRange().getValues();
  var supplierData = supplierSheet.getDataRange().getValues();
  
  // 2. สร้างพจนานุกรม (Dictionary) ไว้เชื่อมตาราง Products และ Stock 
  var productDict = {};
  for (var i = 1; i < productData.length; i++) {
    var pId = String(productData[i][0]).trim().toUpperCase();
    productDict[pId] = { name: productData[i][1], category: productData[i][2], price: productData[i][3] };
  }
  
  var stockDict = {};
  for (var i = 1; i < stockData.length; i++) {
    var sId = String(stockData[i][0]).trim().toUpperCase();
    stockDict[sId] = { qty: stockData[i][2] };
  }
  
  // 3. เตรียมหัวตารางสำหรับแท็บ Cleaned_Data และ Quarantine
  var cleanedRows = [["SKU_ID", "Product Name", "Category", "Unit Price", "Available Qty", "Supplier Name", "Daily Burn Rate", "Lead Time (Days)", "Days Left", "Risk Status"]];
  var quarantineRows = [["SKU_ID", "Supplier Name", "Burn Rate", "Lead Time", "Contract Date", "Contact No", "Quarantine Reason (สาเหตุที่กักกัน)"]];
  
  // 4. กระบวนการตรวจสอบและคัดแยกข้อมูล (Data Cleaning & Validation)
  for (var i = 1; i < supplierData.length; i++) {
    var row = supplierData[i];
    if (String(row[0]).trim() === "") continue; // ข้ามบรรทัดว่าง
    
    // ปัญหาที่ 1: Format Error (รหัสมีช่องว่างหรือตัวพิมพ์เล็ก) -> โค้ดจะคลีนเป็นตัวพิมพ์ใหญ่อัตโนมัติ
    var cleanSku = String(row[0]).trim().toUpperCase();
    var supplierName = row[1];
    var burnRate = parseFloat(row[2]);
    var leadTime = parseInt(row[3]);
    var contractDate = String(row[4]);
    var contactNo = String(row[5]).trim();
    
    var reasons = [];
    
    // ปัญหาที่ 2: Business Rule (อัตราการขาย หรือ เวลาจัดส่ง ติดลบ)
    if (burnRate <= 0 || leadTime <= 0) {
      reasons.push("ค่าติดลบ/ศูนย์ผิดกฎธุรกิจ (Burn Rate/Lead Time <= 0)");
    }
    
    // ปัญหาที่ 3: Ambiguous/Unverified Data (วันที่กำกวมมีสแลช / หรือ ไม่มีเบอร์ติดต่อ)
    if (contractDate.indexOf("/") !== -1) {
      reasons.push("วันที่กำกวม (รูปแบบไม่ใช่ YYYY-MM-DD)");
    }
    if (contactNo === "") {
      reasons.push("ไม่มีเบอร์ติดต่อยืนยันซัพพลายเออร์");
    }
    
    // 5. แยกข้อมูลลงแท็บที่ถูกต้อง
    if (reasons.length > 0) {
      // ถ้าพบปัญหา -> ส่งเข้าห้องกักกัน (Quarantine)
      quarantineRows.push([cleanSku, supplierName, burnRate, leadTime, contractDate, contactNo, reasons.join(", ")]);
    } else {
      // ถ้าข้อมูลคลีน -> นำไป Join กับตาราง Products และ Stock
      var pInfo = productDict[cleanSku] || {name: "Unknown", category: "Unknown", price: 0};
      var sInfo = stockDict[cleanSku] || {qty: 0};
      
      var daysLeft = (sInfo.qty / burnRate).toFixed(1);
      
     // ประเมินความเสี่ยง (Risk Assessment) - ฉบับปรับสมดุล
      var safetyStockDays = 2; // วันเผื่อฉุกเฉิน (ใช้สำหรับแจ้งเตือน MEDIUM)
      var minQtyThreshold = 2; // จำนวนชิ้นขั้นต่ำ (ใช้สำหรับแจ้งเตือน CRITICAL ทันที)
      
      var riskStatus = "✅ SAFE (ปลอดภัย)";
      
      // เงื่อนไขที่ 1: แดงวิกฤต (CRITICAL) = วันเหลือน้อยกว่าหรือเท่ากับเวลาส่งของเป๊ะๆ หรือ ของเหลือน้อยกว่า 2 ชิ้น
      if (daysLeft <= leadTime || sInfo.qty <= minQtyThreshold) {
        riskStatus = "🔥 CRITICAL (เสี่ยงของขาด ต้องสั่งด่วน!)";
      } 
      // เงื่อนไขที่ 2: เหลืองเตือนภัย (MEDIUM) = ยังไม่วิกฤต แต่เข้าสู่ช่วง Safety Stock (เริ่มนับถอยหลัง)
      else if (daysLeft <= (leadTime + safetyStockDays)) {
        riskStatus = "🟡 MEDIUM (ความเสี่ยงปานกลาง เตรียมสั่งซื้อ)";
      }
      
      cleanedRows.push([cleanSku, pInfo.name, pInfo.category, pInfo.price, sInfo.qty, supplierName, burnRate, leadTime, daysLeft, riskStatus]);
    }
  }
  
  // 6. พิมพ์ข้อมูลลงใน Google Sheets
  cleanedSheet.clearContents();
  quarantineSheet.clearContents();
  
  if (cleanedRows.length > 1) {
    cleanedSheet.getRange(1, 1, cleanedRows.length, cleanedRows[0].length).setValues(cleanedRows);
  }
  if (quarantineRows.length > 1) {
    quarantineSheet.getRange(1, 1, quarantineRows.length, quarantineRows[0].length).setValues(quarantineRows);
  }
  
  SpreadsheetApp.getUi().alert("✅ Data Pipeline ทำงานเสร็จสิ้น! ข้อมูลถูกคัดแยกไปที่แท็บ Cleaned_Data และ Quarantine แล้วครับ");
}
//-------------------------------------------------------------------------------------------------------------------------------------------------


// เอาไว้ใช้ "ใน" Google Sheets (จำลองการยิง API ให้อาจารย์ดูง่ายๆ โดยไม่ต้องออกไปเข้าเว็บอื่น)
function showAPITesterDialog() {
  var html = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px;">
      <label for="skuTest"><b>ระบุรหัสสินค้าที่ต้องการทดสอบ API:</b></label><br>
      <input type="text" id="skuTest" style="width: 100%; padding: 5px; margin-bottom: 10px;" placeholder="เช่น SKU-0001, SKU-0003, SKU-9999"><br>
      
      <button onclick="testAPI()" style="padding: 8px 15px; background-color: #2196F3; color: white; border: none; cursor: pointer;">ยิง API ทดสอบ</button>
      <button onclick="google.script.host.close()" style="padding: 8px 15px; background-color: #f44336; color: white; border: none; cursor: pointer; margin-left: 10px;">ปิด</button>
      
      <div id="resultBox" style="margin-top: 15px; background-color: #333; color: #00FF00; padding: 10px; border-radius: 5px; font-family: monospace; white-space: pre-wrap; min-height: 100px; display: none;"></div>
    </div>
    <script>
      function testAPI() {
        var sku = document.getElementById('skuTest').value;
        var resultBox = document.getElementById('resultBox');
        
        resultBox.style.display = 'block';
        resultBox.innerText = '⏳ กำลังยิง API...';
        
        google.script.run
          .withSuccessHandler(function(responseJSON) {
             var jsonObj = JSON.parse(responseJSON);
             resultBox.innerText = JSON.stringify(jsonObj, null, 2);
          })
          .simulateAPIRequest(sku);
      }
    </script>
  `)
  .setWidth(450)
  .setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'ทดสอบ REST API แบบจำลอง');
}

// ฟังก์ชันการรับ Request 
function simulateAPIRequest(sku) {
  var skuSearch = String(sku).trim().toUpperCase();
  
  if (!skuSearch) {
    return JSON.stringify({"status": "error", "message": "API Requirements: กรุณาระบุรหัสสินค้า"});
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cleanedSheet = ss.getSheetByName("Cleaned_Data");
  var data = cleanedSheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === skuSearch) {
      return JSON.stringify({
        "status": "success",
        "data": {
          "sku_id": data[i][0],
          "product_name": data[i][1],
          "available_qty": data[i][4],
          "supplier_name": data[i][5],
          "days_left": data[i][8],
          "risk_status": data[i][9]
        }
      });
    }
  }
  
  return JSON.stringify({"status": "not_found", "message": "ไม่พบข้อมูลสินค้ารหัส " + skuSearch + " หรือข้อมูลซัพพลายเออร์ถูกระงับ (Quarantined)"});
}


// อัปเดตฟังก์ชันเมนู (เอาไปแทนที่ฟังก์ชัน onOpen เดิม)
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 ระบบจัดการสต็อก BaanDee')
      .addItem('1. รันระบบคัดกรองข้อมูล (Run Pipeline)', 'runDataPipeline')
      .addSeparator()
      .addItem('2. สรุปรายการต้องสั่งด่วน (Dashboard)', 'showCriticalSummary')
      .addSeparator()
      .addItem('3. ออกใบสั่งซื้อ (Create PO)', 'showCreatePODialog')
      .addItem('4. รับสินค้าเข้าคลัง (Receive PO)', 'showReceivePODialog') // บรรทัดที่เพิ่งเพิ่มเข้ามา
      .addSeparator()
      .addItem('5. ทดสอบ API (API Tester)', 'showAPITesterDialog')
      .addToUi();
}


// --- ส่วนที่เพิ่มใหม่: ระบบสรุปรายการสั่งซื้อด่วน (Summary Dashboard) ---

function showCriticalSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cleaned_Data");
  var data = sheet.getDataRange().getValues();
  var criticalItems = [];
  
  // วนลูปเช็กข้อมูลทั้งหมด (รองรับเป็นพันๆ แถว)
  for(var i=1; i<data.length; i++) {
    if(String(data[i][9]).indexOf("CRITICAL") !== -1) {
      criticalItems.push(data[i]);
    }
  }
  
  // ถ้าไม่มีของต้องสั่งเลย
  if(criticalItems.length === 0) {
    SpreadsheetApp.getUi().alert("🎉 ยอดเยี่ยม! วันนี้ไม่มีสินค้าที่เสี่ยงขาดสต็อกเลยครับ (สต็อกปลอดภัยทั้งหมด)");
    return;
  }
  
  // สร้างหน้าต่าง HTML สรุปรายการ
  var html = '<div style="font-family: sans-serif; padding: 10px;">';
  html += '<h3 style="color: #d32f2f;">🔥 พบสินค้าต้องสั่งด่วน ' + criticalItems.length + ' รายการ</h3>';
  html += '<p>กรุณานำรหัสเหล่านี้ไปสร้างใบสั่งซื้อ (Create PO):</p>';
  html += '<table border="1" style="border-collapse: collapse; width: 100%; text-align: left; margin-bottom: 15px;">';
  html += '<tr style="background-color: #f2f2f2;"><th>รหัสสินค้า</th><th>ชื่อสินค้า</th><th>เวลาเหลือ</th></tr>';
  
  for(var j=0; j<criticalItems.length; j++) {
    html += '<tr>';
    html += '<td><b>' + criticalItems[j][0] + '</b></td>';
    html += '<td>' + criticalItems[j][1] + '</td>';
    html += '<td style="color: red; font-weight: bold;">' + criticalItems[j][8] + ' วัน</td>';
    html += '</tr>';
  }
  html += '</table>';
  html += '<button onclick="google.script.host.close()" style="padding: 8px 15px; background-color: #2196F3; color: white; border: none; cursor: pointer; width: 100%;">รับทราบ & ปิดหน้าต่าง</button>';
  html += '</div>';
  
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 สรุปรายงาน Dashboard');
}
//--------------------------------------------------------------------------------------------------------------------------------------------------

//สร้าง "หน้าต่าง Pop-up สำหรับกรอกข้อมูลใบสั่งซื้อ"
function showCreatePODialog() {
  var html = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px;">
      <label for="skuId"><b>รหัสสินค้า (SKU_ID):</b></label><br>
      <input type="text" id="skuId" style="width: 100%; padding: 5px; margin-bottom: 10px;" placeholder="เช่น SKU-0001"><br>
      
      <label for="orderQty"><b>จำนวนที่ต้องการสั่ง (ชิ้น):</b></label><br>
      <input type="number" id="orderQty" style="width: 100%; padding: 5px; margin-bottom: 10px;" min="1" value="10"><br>
      
      <label for="note"><b>หมายเหตุ (ตัวเลือก):</b></label><br>
      <input type="text" id="note" style="width: 100%; padding: 5px; margin-bottom: 15px;" placeholder="เช่น สั่งด่วนของใกล้หมด"><br>
      
      <button onclick="submitForm()" style="padding: 8px 15px; background-color: #4CAF50; color: white; border: none; cursor: pointer;">บันทึกใบสั่งซื้อ</button>
      <button onclick="google.script.host.close()" style="padding: 8px 15px; background-color: #f44336; color: white; border: none; cursor: pointer; margin-left: 10px;">ยกเลิก</button>
    </div>
    <script>
      function submitForm() {
        var sku = document.getElementById('skuId').value;
        var qty = document.getElementById('orderQty').value;
        var note = document.getElementById('note').value;
        
        if (!sku || !qty) {
          alert('กรุณากรอกรหัสสินค้าและจำนวนให้ครบถ้วน');
          return;
        }
        
        google.script.run
          .withSuccessHandler(function(msg) {
             alert(msg);
             google.script.host.close();
          })
          .processPurchaseOrder(sku, qty, note);
      }
    </script>
  `)
  .setWidth(350)
  .setHeight(250);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'ออกใบสั่งซื้อ (Create Purchase Order)');
}
//ดึงข้อมูลราคาสินค้ามาคำนวณยอดรวม (Total Cost) แล้วบันทึกข้อมูลทั้งหมดลงในแท็บ Purchase_Orders
function processPurchaseOrder(sku, qtyStr, note) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cleanedSheet = ss.getSheetByName("Cleaned_Data");
  var poSheet = ss.getSheetByName("Purchase_Orders");
  
  var skuClean = String(sku).trim().toUpperCase();
  var orderQty = parseInt(qtyStr);
  
  if (isNaN(orderQty) || orderQty <= 0) {
    return "❌ ข้อผิดพลาด: จำนวนสั่งซื้อต้องมากกว่า 0";
  }
  
  var cleanedData = cleanedSheet.getDataRange().getValues();
  var foundItem = null;
  
  for (var i = 1; i < cleanedData.length; i++) {
    if (String(cleanedData[i][0]).toUpperCase() === skuClean) {
      foundItem = cleanedData[i];
      break;
    }
  }
  
  if (!foundItem) {
    return "❌ ข้อผิดพลาด: ไม่พบรหัสสินค้า " + skuClean + " ในระบบ หรือสินค้าอาจถูกกักกัน (Quarantined) อยู่";
  }
  
  var unitPrice = parseFloat(foundItem[3]);
  var supplierName = foundItem[5];
  var totalCost = unitPrice * orderQty;
  
  if (poSheet.getLastRow() === 0) {
    poSheet.appendRow(["PO_ID", "SKU_ID", "Supplier Name", "Order Qty", "Total Cost (THB)", "Status", "Reviewer Note", "Timestamp"]);
    poSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#e0e0e0");
  }
  
  var timestamp = new Date();
  var timestampStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var poId = "PO-" + timestamp.getTime().toString().substring(5);
  var status = "Approved";
  
  poSheet.appendRow([poId, skuClean, supplierName, orderQty, totalCost, status, note, timestampStr]);
  
  return "✅ บันทึกใบสั่งซื้อ " + poId + " สำเร็จ! (ยอดรวม " + totalCost.toLocaleString() + " บาท)";
}

// --- ระบบรับสินค้าเข้าคลัง (Goods Receipt) ---
function showReceivePODialog() {
  var html = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px;">
      <label for="poId"><b>ระบุรหัสใบสั่งซื้อ (PO_ID) ที่มาส่งมอบ:</b></label><br>
      <input type="text" id="poId" style="width: 100%; padding: 5px; margin-bottom: 15px;" placeholder="เช่น PO-25697265"><br>
      
      <button onclick="submitReceive()" style="padding: 8px 15px; background-color: #FF9800; color: white; border: none; cursor: pointer;">📦 รับของเข้าสต็อก</button>
      <button onclick="google.script.host.close()" style="padding: 8px 15px; background-color: #f44336; color: white; border: none; cursor: pointer; margin-left: 10px;">ยกเลิก</button>
    </div>
    <script>
      function submitReceive() {
        var poId = document.getElementById('poId').value;
        if (!poId) {
          alert('กรุณากรอกรหัส PO_ID');
          return;
        }
        
        google.script.run
          .withSuccessHandler(function(msg) {
             alert(msg);
             google.script.host.close();
          })
          .processReceivePO(poId);
      }
    </script>
  `)
  .setWidth(350)
  .setHeight(150);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'รับสินค้าเข้าคลัง (Receive Goods)');
}

function processReceivePO(poId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var poSheet = ss.getSheetByName("Purchase_Orders");
  var stockSheet = ss.getSheetByName("Raw_Stock");
  
  var poData = poSheet.getDataRange().getValues();
  var foundRow = -1;
  var sku = "";
  var qty = 0;
  var status = "";
  
  // 1. ค้นหาใบสั่งซื้อจาก PO_ID
  for (var i = 1; i < poData.length; i++) {
    if (String(poData[i][0]).trim() === poId.trim()) {
      foundRow = i + 1; // ตำแหน่งแถวในชีต (บวก 1 เพราะ index เริ่มที่ 0)
      sku = String(poData[i][1]).trim();
      qty = parseInt(poData[i][3]);
      status = String(poData[i][5]);
      break;
    }
  }
  
  if (foundRow === -1) {
    return "❌ ข้อผิดพลาด: ไม่พบรหัสใบสั่งซื้อ " + poId;
  }
  
  if (status.indexOf("Received") !== -1) {
    return "⚠️ แจ้งเตือน: ใบสั่งซื้อนี้ถูกรับเข้าคลังไปแล้ว ไม่สามารถรับซ้ำได้ครับ!";
  }
  
  // 2. ไปอัปเดตจำนวนในแท็บ Raw_Stock
  var stockData = stockSheet.getDataRange().getValues();
  var stockUpdated = false;
  
  for (var j = 1; j < stockData.length; j++) {
    if (String(stockData[j][0]).trim().toUpperCase() === sku.toUpperCase()) {
      var currentQty = parseInt(stockData[j][2]) || 0; // คอลัมน์ C คือ available_qty
      var newQty = currentQty + qty;
      
      // อัปเดตจำนวนใหม่ในคอลัมน์ C
      stockSheet.getRange(j + 1, 3).setValue(newQty); 
      
      // อัปเดตเวลาล่าสุด (as_of) ในคอลัมน์ D อัตโนมัติ
      var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
      stockSheet.getRange(j + 1, 4).setValue(timestampStr); 
      
      stockUpdated = true;
      break;
    }
  }
  
  if (!stockUpdated) {
    return "❌ ข้อผิดพลาด: ไม่พบสินค้ารหัส " + sku + " ในแท็บ Raw_Stock";
  }
  
  // 3. เปลี่ยนสถานะใบสั่งซื้อเป็น Received
  poSheet.getRange(foundRow, 6).setValue("Received (รับเข้าคลังแล้ว)");
  
  return "✅ อัปเดตสำเร็จ! รับสินค้า " + sku + " จำนวน " + qty + " ชิ้น เข้าสู่ Raw_Stock เรียบร้อยแล้วครับ";
}
//--------------------------------------------------------------------------------------------------
// --- ฟังก์ชันหลักสำหรับ Web App และ API ---
function doGet(e) {
  // 1. ถ้ามีการแนบ parameter ?sku= มา (แปลว่าระบบภายนอกกำลังยิง API มา)
  if (e.parameter.sku) {
    var jsonString = simulateAPIRequest(e.parameter.sku);
    return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. เข้ามาหน้าเว็บ (คนใช้งาน)(Web UI)
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;600&display=swap" rel="stylesheet">
      <style>
        body { font-family: 'Prompt', sans-serif; background-color: #f4f7f6; padding: 20px; display: flex; justify-content: center; align-items: center; min-height: 90vh; margin: 0; }
        .container { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); max-width: 450px; width: 100%; text-align: center; }
        h2 { color: #2c3e50; margin-bottom: 5px; }
        p { color: #7f8c8d; margin-bottom: 25px; font-size: 14px; }
        input[type="text"] { width: calc(100% - 24px); padding: 12px; margin-bottom: 15px; border: 1.5px solid #bdc3c7; border-radius: 8px; font-size: 16px; transition: 0.3s; font-family: 'Prompt', sans-serif;}
        input[type="text"]:focus { border-color: #3498db; outline: none; }
        button { width: 100%; padding: 12px; background-color: #3498db; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: 0.3s; font-family: 'Prompt', sans-serif;}
        button:hover { background-color: #2980b9; }
        #result { margin-top: 20px; text-align: left; padding: 15px; border-radius: 8px; display: none; background-color: #f8f9fa; border-left: 5px solid #3498db; line-height: 1.6; font-size: 15px;}
        .error-box { border-left-color: #e74c3c !important; color: #c0392b; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📦 BaanDee Stock Portal</h2>
        <p>ระบบตรวจสอบสถานะสินค้าคงคลัง (Real-time)</p>
        <input type="text" id="skuInput" placeholder="พิมพ์รหัสสินค้าที่นี่ (เช่น SKU-0001)">
        <button onclick="searchData()" id="searchBtn">🔍 ค้นหาสินค้า</button>
        
        <div id="result"></div>
      </div>
      
      <script>
        function searchData() {
          var sku = document.getElementById('skuInput').value;
          var resultDiv = document.getElementById('result');
          var btn = document.getElementById('searchBtn');
          
          if(!sku) {
            alert('กรุณากรอกรหัสสินค้าก่อนค้นหาครับ');
            return;
          }
          
          btn.innerText = '⏳ กำลังค้นหา...';
          btn.disabled = true;
          resultDiv.style.display = 'block';
          resultDiv.className = '';
          resultDiv.innerHTML = 'กำลังดึงข้อมูลจากระบบ...';
          
          google.script.run
            .withSuccessHandler(function(res) {
              btn.innerText = '🔍 ค้นหาสินค้า';
              btn.disabled = false;
              
              var obj = JSON.parse(res);
              if (obj.status === 'success') {
                resultDiv.style.borderLeftColor = '#2ecc71';
                resultDiv.innerHTML = 
                  "<b>🏷️ รหัสสินค้า:</b> " + obj.data.sku_id + "<br>" +
                  "<b>📦 ชื่อสินค้า:</b> " + obj.data.product_name + "<br>" +
                  "<b>🏢 ผู้จัดจำหน่าย:</b> " + obj.data.supplier_name + "<br>" +
                  "<b>🔢 จำนวนคงเหลือ:</b> " + obj.data.available_qty + " ชิ้น<br>" +
                  "<b>⚠️ สถานะ:</b> " + obj.data.risk_status;
              } else {
                resultDiv.classList.add('error-box');
                resultDiv.innerHTML = "❌ <b>ไม่พบข้อมูล:</b><br>" + obj.message;
              }
            })
            .simulateAPIRequest(sku);
        }
      </script>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html)
    .setTitle('BaanDee Stock Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}