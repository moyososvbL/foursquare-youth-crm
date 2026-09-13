/**
 * Main automation engine for Foursquare Youth CRM.
 * Handles daily member database checks, automated HTML birthday cards via email,
 * direct WhatsApp birthday messages, individual service reminder emails,
 * and group WhatsApp service broadcasts.
 */
function sendYouthAutomations() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!activeSpreadsheet) {
    Logger.log("Error: Script is not bound to a spreadsheet.");
    return;
  }
  
  var sheet = activeSpreadsheet.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  
  // 1. Clean column headers to remove hidden line breaks or trailing spaces
  var headers = data[0].map(function(header) {
    return String(header).trim().replace(/\s+/g, ' '); 
  });
  
  var nameIdx = headers.indexOf("Full Name");
  var emailIdx = headers.indexOf("Email Address");
  var dobIdx = headers.indexOf("Date of Birth");
  var phoneIdx = headers.indexOf("Phone Number");

  // Error trap to make sure critical database columns exist
  if (nameIdx === -1 || emailIdx === -1 || dobIdx === -1 || phoneIdx === -1) {
    Logger.log("Error: Columns Missing! Check header names.");
    return;
  }

  var today = new Date();
  var todayMonth = today.getMonth() + 1; 
  var todayDate = today.getDate();
  var dayOfWeek = today.getDay(); // 0 = Sunday, 2 = Tuesday

  var isFirstSunday = (dayOfWeek === 0 && todayDate <= 7);
  var isServiceDay = false;
  var emailSubject = "";
  var emailHtmlTemplate = "";
  var whatsappTemplate = ""; 

  // Hosted church logo URL used across all email design templates
  var churchLogoUrl = "https://i.postimg.cc/ncCsT6WR/images-(1).jpg"; 

  // 2. Define Tuesday Prayer Service Logic
  if (dayOfWeek === 2) {
    isServiceDay = true;
    emailSubject = "🔥 Reminder: Youth Power Prayer Service Tonight!";
    
    whatsappTemplate = "🔥 *Youth Power Prayer Service Tonight!*\n\nJoin us tonight for our weekly Youth Prayer Service.\n⏰ Time: 8:00 PM – 9:00 PM\n📍 Venue: WhatsApp Online Call\n\nCome expectant. See you there!";
    
    emailHtmlTemplate = `
      <div style="font-family: Arial, sans-serif; max-width: 520px; margin: 0 auto; padding: 30px; border: 1px solid #eaeaea; border-radius: 16px; background-color: #ffffff; box-shadow: 0 4px 12px rgba(0,0,0,0.03);">
        <div style="text-align: center; margin-bottom: 25px;">
          <img src="${churchLogoUrl}" alt="Foursquare Logo" style="width: 70px; height: auto; margin-bottom: 15px; border-radius: 8px;">
          <h2 style="color: #e74c3c; margin: 0; font-size: 24px;">Youth Power Prayer Service 🔥</h2>
        </div>
        
        <p style="font-size: 16px; color: #333333; line-height: 1.5;">
          Hi <strong>{Name}</strong> 👋,
        </p>
        
        <p style="font-size: 16px; color: #333333; line-height: 1.5;">
          Join us tonight for our weekly Youth Prayer Service as we connect and seek God together!
        </p>
        
        <div style="background-color: #fcf3f2; padding: 20px; border-radius: 12px; margin: 25px 0; border-left: 4px solid #e74c3c;">
          <p style="margin: 8px 0; color: #333333; font-size: 15px;">⏰ <strong>Time:</strong> 8:00 PM – 9:00 PM (1 Hour)</p>
          <p style="margin: 8px 0; color: #333333; font-size: 15px;">📍 <strong>Venue:</strong> WhatsApp Online Call</p>
        </div>
        
        <p style="font-size: 16px; color: #333333; line-height: 1.5;">
          Come expectant. See you there!
        </p>
        
        <hr style="border: none; border-top: 1px solid #eeeeee; margin: 30px 0;">
        <p style="font-size: 13px; color: #888888; text-align: center; margin: 0;">
          Foursquare Gospel Youth Church • Priceless Youths
        </p>
      </div>
    `;
  } 
  // 3. Define Sunday Service Logic (Excludes the first Sunday of the month)
  else if (dayOfWeek === 0 && !isFirstSunday) {
    isServiceDay = true;
    emailSubject = "⛪ Sunday Youth Service is Live This Morning!";
    
    whatsappTemplate = "⛪ *Sunday Youth Service is Live!*\n\nHappy Sunday! Our Youth Service is holding this morning, and we can't wait to fellowship with you.\n⏰ Time: 7:00 AM – 8:30 AM\n📍 Venue: Youth Auditorium\n\nCome with a friend!";
    
    emailHtmlTemplate = `
      <div style="font-family: Arial, sans-serif; max-width: 520px; margin: 0 auto; padding: 30px; border: 1px solid #eaeaea; border-radius: 16px; background-color: #ffffff; box-shadow: 0 4px 12px rgba(0,0,0,0.03);">
        <div style="text-align: center; margin-bottom: 25px;">
          <img src="${churchLogoUrl}" alt="Foursquare Logo" style="width: 70px; height: auto; margin-bottom: 15px; border-radius: 8px;">
          <h2 style="color: #2980b9; margin: 0; font-size: 24px;">Sunday Youth Service ⛪</h2>
        </div>
        
        <p style="font-size: 16px; color: #333333; line-height: 1.5;">
          Hi <strong>{Name}</strong> 👋,
        </p>
        
        <p style="font-size: 16px; color: #333333; line-height: 1.5;">
          Happy Sunday! Our Youth Service is holding this morning, and we can't wait to fellowship in God's presence with you.
        </p>
        
        <div style="background-color: #f4f8fb; padding: 20px; border-radius: 12px; margin: 25px 0; border-left: 4px solid #2980b9;">
          <p style="margin: 8px 0; color: #333333; font-size: 15px;">⏰ <strong>Time:</strong> 7:00 AM – 8:30 AM</p>
          <p style="margin: 8px 0; color: #333333; font-size: 15px;">📍 <strong>Venue:</strong> Youth Auditorium</p>
        </div>
        
        <p style="font-size: 16px; color: #333333; line-height: 1.5;">
          Come with a friend and be blessed!
        </p>
        
        <hr style="border: none; border-top: 1px solid #eeeeee; margin: 30px 0;">
        <p style="font-size: 13px; color: #888888; text-align: center; margin: 0;">
          Foursquare Gospel Youth Church • Priceless Youths
        </p>
      </div>
    `;
  }

  // 4. Iterate Through Member Database Row by Row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var memberName = row[nameIdx];
    var memberEmail = row[emailIdx];
    var dobValue = row[dobIdx];
    var memberPhone = row[phoneIdx];

    // Skip rows missing a valid email address
    if (!memberEmail) continue;

    // --- A. Individual Birthday Check Engine ---
    if (dobValue) {
      var birthMonth, birthDate;
      
      if (dobValue instanceof Date) {
        birthMonth = dobValue.getMonth() + 1;
        birthDate = dobValue.getDate();
      } else {
        var parsedDate = new Date(dobValue);
        if (!isNaN(parsedDate.getTime())) {
          birthMonth = parsedDate.getMonth() + 1;
          birthDate = parsedDate.getDate();
        }
      }

      // If today matches the member's birthday, dispatch celebration messages
      if (todayMonth === birthMonth && todayDate === birthDate) {
        var bdaySubject = "🎉 Happy Birthday from the Youth Ministry! 🎂";
        var bdayHtml = `
          <div style="font-family: Arial, sans-serif; max-width: 520px; margin: 0 auto; padding: 30px; border: 1px solid #eaeaea; border-radius: 16px; background-color: #ffffff; box-shadow: 0 4px 12px rgba(0,0,0,0.03);">
            <div style="text-align: center; margin-bottom: 25px;">
              <img src="${churchLogoUrl}" alt="Foursquare Logo" style="width: 70px; height: auto; margin-bottom: 15px; border-radius: 8px;">
              <h1 style="color: #f39c12; margin: 0; font-size: 26px;">🎉 Happy Birthday! 🎉</h1>
            </div>
            
            <p style="font-size: 16px; color: #333333; line-height: 1.5;">
              Dear <strong>${memberName}</strong> 🌟,
            </p>
            
            <p style="font-size: 16px; color: #333333; line-height: 1.5;">
              The entire Foursquare Gospel Youth Church family wishes you a very Happy Birthday today! We pray that God grants you supernatural wisdom, boundless strength, and opening doors in this new year of your life.
            </p>
            
            <p style="font-size: 16px; color: #333333; line-height: 1.5;">
              Have a beautiful, joy-filled celebration!
            </p>
            
            <hr style="border: none; border-top: 1px solid #eeeeee; margin: 30px 0;">
            <p style="font-size: 13px; color: #888888; text-align: center; margin: 0;">
              With love and blessings,<br>
              <strong>Youth Ministry Team</strong>
            </p>
          </div>
        `;
        
        // Send birthday email
        MailApp.sendEmail({
          to: memberEmail,
          subject: bdaySubject,
          htmlBody: bdayHtml
        });
        Logger.log("Success! Birthday HTML email sent to: " + memberEmail);

        // Send direct WhatsApp birthday message if phone number exists
        if (memberPhone) {
          var whatsappBdayText = "🎉 Happy Birthday " + memberName + "! \n\nThe Foursquare Gospel Youth Church wishes you a beautiful day filled with God's blessings, love, and prosperity.";
          sendWhatsApp(memberPhone, whatsappBdayText);
        }
      }
    }

    // --- B. Individual Service Reminder Email Engine ---
    if (isServiceDay) {
      var personalizedHtml = emailHtmlTemplate.replace("{Name}", memberName);
      
      MailApp.sendEmail({
        to: memberEmail,
        subject: emailSubject,
        htmlBody: personalizedHtml
      });
      Logger.log("Success! Service HTML reminder sent to: " + memberEmail);
    }
  }

  // 5. Group WhatsApp Service Broadcast (Fires only once per service day)
  if (isServiceDay) {
    var groupId = "GROUP ID FROM ultramsg.com"; 
    sendWhatsApp(groupId, whatsappTemplate);
    Logger.log("Success! Service reminder broadcasted to the Priceless Youths WhatsApp Group.");
  }
}

/**
 * Secure Helper Function for UltraMsg API Integration.
 * Pulls instance ID and token securely from Google Apps Script Script Properties 
 * instead of exposing hardcoded credentials on public repositories.
 */
function sendWhatsApp(phoneNumberOrGroupId, textMessage) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var instanceId = scriptProperties.getProperty('INSTANCE_ID');
  var token = scriptProperties.getProperty('TOKEN');
  
  var apiUrl = "https://api.ultramsg.com/" + instanceId + "/messages/chat"; 

  var payload = {
    "token": token,
    "to": phoneNumberOrGroupId,
    "body": textMessage
  };

  var options = {
    "method": "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload": payload,
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log("WhatsApp API Response: " + response.getContentText());
  } catch (error) {
    Logger.log("WhatsApp routing failed | Error: " + error);
  }
}
