# Rojnamcha Ledger - Project Files & Architecture Guide

Is guide me yeh bataya gaya hai ki is project ki kaunsi files **Netlify Frontend** ki hain, kaunsi **Google Sheets Backend** ki hain, aur **Firebase Backend** kaise kaam karta hai.

---

## 🖥️ 1. Netlify Frontend (Client-Side)
Yeh files web browser me user interface (UI) aur screens ko dikhane aur control karne ke liye hain.

* **`index.html` (Main Frontend Entry Point):**
  * Yeh Netlify par deploy hone wali **sabse important** single compiled file hai.
  * Isme poore system ka HTML, CSS styling, aur JavaScript logic ek hi jagah bundle hai.
  * Netlify is `index.html` ko read karke user ke browser par screen open karta hai.
* **`CSS.html`:**
  * Isme design, animations, dark theme aur layouts ki styling code hai. (Yeh `index.html` me compile hoti hai).
* **`JS.html`:**
  * Isme saari client-side logic hai. Jaise buttons ke click handlers, local calculations, aur database (Sheets/Firebase) se data send/fetch karne ka code. (Yeh `index.html` me compile hoti hai).

---

## 📊 2. Google Sheets Backend (Server-Side)
Agar aap **Google Sheets** ko database ki tarah chalana chahte hain, to yeh files Google Apps Script par backend banati hain.

* **`Code.js` (Backend API Routes):**
  * Yeh Google Apps Script ka backend controller hai.
  * Isme sheet me data write karne, read karne, date filter reports banane, aur database initialization ke saare functions likhe hain.
* **`EmailTemplate.html`:**
  * Jab aap reports email karte hain, to email ka design aur format kaisa hoga, uska code isme hai.
* **`appsscript.json`:**
  * Google Apps Script project ki internal configuration setting file hai.

---

## ⚡ 3. Firebase Backend (Serverless Cloud Database)
Firebase backend me aapke computer ya local folder me **koi alag file nahi hoti**. 

Firebase **Serverless** hai, yani iska backend seedhe Google ke Cloud servers par chalta hai.
* **Database Engine:** Firestore Database (NoSQL) aapke Firebase Console (online) me store hota hai.
* **Connection Logic:** `index.html` (jo `JS.html` se compiled hai) ke andar Firebase Web SDK (compat CDN) daala gaya hai, jo browser se **seedhe** Firebase Database me read/write karta hai.
* **Firestore Collections (Tables):**
  * `users` (Users and Approvals)
  * `dropdown_options` (Receipt/Expense options)
  * `company_profile` (GSTIN, Logo, Details)
  * `rojnamcha` (Daily summary sheet data)
  * `rojnamcha_details` (Ledger receipts and expenses list)

---

## ⚙️ 4. Developer Settings & Config Files
Yeh files development tools (Git aur Clasp) ke liye hain.

* **`.clasp.json`:** Apps Script ko aapke local folder se connect rakhne ki script ID setting.
* **`.claspignore`:** Clasp ko kaunsi files server par upload nahi karni chahiye (jaise Netlify ki `index.html` ya backup folders).
* **`.gitignore`:** GitHub par push karte waqt local test/backup files ko ignore karne ke liye.
