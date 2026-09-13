# ⛪ Foursquare Youth Church Automated CRM

An automated church management and communication engine built using the Google Workspace ecosystem and the UltraMsg WhatsApp API. Designed to streamline administrative tasks, pastoral care, and member engagement for the Priceless Youths (FUTA Chapter).

## 🚀 Features

- **Automated Data Pipeline:** Captures member data instantly via Google Forms linked to a Google Sheets database.
- **Smart Service Scheduling:** Automatically detects service days (Tuesdays and Sundays—excluding the first Sunday of the month) and routes reminders.
- **Personalized Birthday Engine:** Scans the database daily to send customized HTML birthday cards via email and direct messages via WhatsApp.
- **WhatsApp Group Integration:** Broadcasts service reminders cleanly to specific WhatsApp group IDs without hitting individual spam limits.
- **Secure Configuration:** Uses Google Apps Script `PropertiesService` to keep sensitive API tokens and instance IDs safe from source control.

## 🛠️ Tech Stack

- **Language:** JavaScript (Google Apps Script)
- **Database & Input:** Google Sheets, Google Forms
- **Communication APIs:** Gmail App (`MailApp`), UltraMsg WhatsApp API (`UrlFetchApp`)
- **Version Control:** Git & GitHub

## ⚙️ Setup & Deployment

1. Bind the script to your Google Sheet containing the member records (`Full Name`, `Email Address`, `Date of Birth`, `Phone Number`).
2. Add your UltraMsg credentials into the Google Apps Script **Project Settings > Script Properties**:
   - `INSTANCE_ID`: Your UltraMsg instance ID
   - `TOKEN`: Your API token
3. Set up a daily time-driven trigger in Apps Script to execute the `sendYouthAutomations` function automatically.
