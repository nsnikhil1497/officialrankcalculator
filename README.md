# 🥇 4th Grade Rank Calculator

यह एक सर्वरलेस (Serverless) वेब एप्लीकेशन है जो उम्मीदवारों को उनके मार्क्स के आधार पर तुरंत रैंक प्रेडिक्शन प्रदान करता है। यह टूल **Google Apps Script** और **Google Sheets** का उपयोग करके बनाया गया है।

---

## 🚀 Live Demo

**टूल का सीधा उपयोग यहाँ करें (Click to use the live tool):**
[https://tinyurl.com/fourthgraderanktool] 

---

## ✨ Features

* **प्रोफेशनल इंटरफ़ेस:** मॉडर्न, प्रीमियम, और इंग्लिश UI/UX।
* **रियल-टाइम प्रेडिक्शन:** स्कोर सबमिट करने पर तुरंत Overall, Category (General, EWS, OBC, SC, ST), और Shift-wise (1-6) रैंक दिखाता है।
* **सटीक मार्किंग:** सही मार्किंग स्कीम के अनुसार Raw Score की गणना: **(+1.66) / (-0.55)**
* **रैंक अपडेट:** 'Check Current Rank' टैब के माध्यम से Roll No. डालकर अपनी वर्तमान रैंक कभी भी चेक करें।
* **डेटा हैंडलिंग:** डुप्लीकेट एंट्री की जाँच करता है।

---

## ⚙️ Technologies Used

* **Backend & Serverless:** Google Apps Script (JavaScript)
* **Database:** Google Sheets (Using the 'Final' sheet)
* **Frontend:** HTML5, CSS3, JavaScript

---

## 💡 How to set up (For Developers)

1.  **Create Google Sheet:** A Google Sheet named **"Final"** is required for the database.
2.  **Apps Script:** Copy the content of `Code.gs` and `Index.html` into a new Google Apps Script project bound to the "Final" sheet.
3.  **Deployment:** Deploy the script as a Web App with access set to **'Anyone'**.
