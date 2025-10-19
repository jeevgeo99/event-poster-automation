🪄 Google Slides Poster Generator for Toastmasters (District 91)

This Google Apps Script automatically generates a branded event poster in Google Slides whenever a new response is submitted via a Google Form.
It uses a template, replaces placeholders with real event data, and emails the final poster (PDF) to the organizer.

🚀 Features

✅ Automatically creates posters from Google Form responses
✅ Uses your own Google Slides template
✅ Emails posters to submitters automatically
✅ Fully customizable placeholders and email template

🧩 How It Works
-> A user fills out your Google Form (with fields like Club, Event Type, Date, Venue, etc.).
-> The script reads the last form response from your linked Google Sheet.
-> It makes a copy of your Google Slides template.
-> It replaces placeholders (like <<Club>>, <<Date>>, <<Venue>>, etc.) with the real data.
-> It sends a PDF copy of the poster to the email address provided.

⚙️ Setup
-> Open your Google Sheet linked to the form.
-> Go to Extensions → Apps Script.
-> Paste the code from CODE.gs
