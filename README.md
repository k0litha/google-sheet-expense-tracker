# 📝 Expense Tracker (Google Forms + Sheets + Apps Script)

A fully fresh, fully editable, and **beautifully styled** expense tracking system built with Google Forms, Google Sheets, and Apps Script.

This project lets you:

✅ Log expenses via a Google Form  
✅ Automatically separate data month-by-month  
✅ Get a live dashboard with charts and budget tracking  
✅ Highlight overspending in real-time

---

## 📂 Project Structure

```
Expense-Tracker/
│
├── README.md           # This file
└── src/
    └── trigger.js      # Google Apps Script automation
```

---

## 🚀 Features

- 📦 Google Form for entering expenses
- 📊 Google Sheet Dashboard with:
  - Total spend vs. budget status
  - Pie chart of expenses by category
  - Bar chart of monthly expenses
  - Top 5 expenses list
- 📅 Automatic creation of monthly sheets
- 🔥 Triggers to run automation on every form submission

---

## 🛠 Setup

### 1️⃣ Create the Google Form

- Go to [Google Forms](https://forms.google.com) and create a new form.
- Add the following fields:

| Field Label      | Type         | Options / Validation                                                       |
| ---------------- | ------------ | -------------------------------------------------------------------------- |
| **Price**        | Short Answer | Number validation                                                          |
| **Name of Item** | Short Answer | Required                                                                   |
| **Importance**   | Dropdown     | Essential, Non-Essential, Somewhat                                         |
| **Category**     | Dropdown     | Food, Transport, Bills, Shopping, Entertainment, Utilities, Health, Others |
| **Payment Type** | Dropdown     | Cash, Card, Bank Transfer                                                  |

- Link the form to a Google Sheet (Responses tab → Google Sheets icon).

---

### 2️⃣ Google Sheet Setup

Once the linked sheet is created:

#### 📑 Tabs to Create:

- `Dashboard` – For charts and budget tracking
- `Config` – To manage categories, payment types, and monthly budget

Rename the auto-generated `Form Responses 1` tab to `Master`.

---

### 3️⃣ Add Config Sheet

Create a tab named `Config` and add:

| A                | B                                                                   |
| ---------------- | ------------------------------------------------------------------- |
| **Setting**      | **Values**                                                          |
| Categories       | Food,Transport,Bills,Shopping,Entertainment,Utilities,Health,Others |
| PaymentTypes     | Cash,Card,Bank Transfer                                             |
| ImportanceLevels | Essential,Non-Essential,Somewhat                                    |
| MonthlyBudget    | 100000                                                              |

---

### 4️⃣ Add Apps Script Automation

- Open your Google Sheet
- Go to `Extensions → Apps Script`
- Replace any starter code with [`src/trigger.js`](src/trigger.js)

---

### 5️⃣ Trigger Setup

- In Apps Script, click the clock icon (⏰) on the left → **Triggers**
- Add a new trigger:
  - Function: `onFormSubmit`
  - Event: From spreadsheet → On form submit

---

## ✨ Live Features

| Feature           | Description                            |
| ----------------- | -------------------------------------- |
| 📅 Monthly Sheets | Automatically creates sheets per month |
| 📊 Dashboard      | Live total, budget, charts, top spends |
| 🔥 Overspending   | Alerts when you go over your budget    |

---

## 📄 Apps Script

All automation logic is in [`src/trigger.js`](src/trigger.js).

---

## 👀 Preview

## 📌 License

MIT © 2025
