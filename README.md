# 💰 Finance Analysis Platform For Financial Reporting and Budget Tracking

### 📊 A Complete AI-Assisted Personal Finance Management Platform

Manage your income, expenses, budgets, investments, financial goals, wallets, bills, and get AI-driven financial insights — all through one connected dashboard.

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Flask](https://img.shields.io/badge/Flask-3.x-black)
![SQLite](https://img.shields.io/badge/Database-SQLite-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Live-brightgreen)

### 🌐 Live Demo

**[🚀 Open Finance Analysis Platform](https://smart-finance-insights-dvjw.onrender.com)**

---

## 📖 Overview

**Finance Analysis Platform For Financial Reporting and Budget Tracking** is a full-stack personal finance management system built using **Python, Flask, SQLite, HTML, CSS, and Chart.js**. It began as a simple expense tracker and grew — across four development milestones — into a 16-module platform combining budgeting, investing, goal planning, portfolio analytics, and a conversational AI assistant.

The platform combines:

- 💰 Financial Management
- 📊 Data Analytics & Reporting
- 🤖 Conversational AI Assistance
- 📈 Data Visualization
- 🔐 Secure Authentication
- ☁️ Live Cloud Deployment

into a single connected web application.

---

## ✨ Key Features

### 💵 Income & Expense Management
- Set and update monthly income
- Add, edit, delete, and search/filter expenses by category or date
- Category-wise expense tracking and monthly summaries
- Recent transactions view with animated summary counters

### 📊 Interactive Dashboard
- Real-time income, expenses, savings, and budget-used stat cards
- Pie chart (expenses by category) and bar chart (income vs. expenses vs. savings)
- Time-based greeting and smart, spending-pattern notifications
- One-click download of a full financial report (CSV, Excel, or Word)

### 🎯 Budget Planning
- Category-wise monthly budgets
- Automatic **On Track / Over Budget** status per category
- Budget utilization percentage and remaining-amount tracking

### 💼 Investment Portfolio
- Track multiple asset types (Stocks, Mutual Funds, Fixed Deposits, Gold, Bonds, Real Estate)
- Automatic profit/loss and return % calculation per holding
- Asset allocation breakdown and overall portfolio ROI

### 🎯 Financial Goals
- Set target amount, saved amount, target date, and priority
- Automatic **monthly savings required**, **days remaining**, and **completion percentage**
- **AI-suggested monthly savings**, splitting leftover income across incomplete goals
- Celebration banner when a goal reaches 100%

### 📈 Portfolio Analytics
- Combined investment + goal analytics dashboard
- Top and lowest performing investments
- Portfolio risk rating (Low / Medium / High) derived from overall ROI

### 💳 Wallets & Accounts
- Track balances separately across Bank Account, Cash, Credit Card, and UPI
- Combined total balance across all wallets

### 📅 Bills & Financial Calendar
- Add one-time or recurring bills with due dates
- Visual monthly calendar showing bills by day
- Urgency-based reminders (overdue / due soon / upcoming)

### 🔄 Subscription Tracker
- Track recurring subscriptions (Netflix, Spotify, etc.)
- Automatic monthly and projected yearly cost totals

### 🏆 Savings Challenges & Milestones
- Create custom savings challenges with daily check-ins and progress tracking
- Automatic **no-spend streak** counter
- Auto-detected financial milestones (first ₹100 saved, goals completed, budget discipline, streaks)

### 🧠 Spending Personality
- Classifies the user as **Saver**, **Balanced**, or **Frequent Spender**
- Based on savings rate, discretionary spending share, and budget adherence
- Includes personalized reasoning and improvement tips

### 📈 What-If Financial Simulator
- Interactive sliders to project future savings under different monthly contributions
- Live comparison chart: current savings path vs. projected path

### 👨‍👩‍👧 Shared Expenses & Bill Splitter
- Split a shared expense across multiple participants and track who has settled
- Standalone bill-splitter calculator with equal or custom percentage splits and tip handling

### 🗂️ Advanced Financial Reports
- Dedicated report pages: Monthly Expense, Budget Utilization, Investment Performance, Goal Progress
- Print-friendly layouts
- Export to **PDF** (ReportLab) and **Excel** (openpyxl), in addition to full-app CSV/Excel/Word export

### 🤖 JARVIS — AI Financial Assistant
A built-in conversational assistant that answers natural-language questions about the user's own financial data.

- Understands questions about expenses, budget, investments, goals, savings, and overall financial health
- Uses a custom rule-based NLP intent detector (Python `re`) — no external AI API required
- Every answer is generated live from the user's real database records
- Maintains conversation history per session, with a clear-chat option

### 👤 User Management & Security
- Registration and login with **hashed passwords** (Werkzeug)
- Secure, HTTP-only, same-site session cookies with automatic expiry
- Server-side input validation on all forms
- Per-user data isolation — every query and edit/delete route is scoped to the logged-in user

### 🎨 UI/UX
- Persistent left-sidebar navigation with a profile block
- Gradient-colored stat cards, animated number counters, and hover effects
- Loading spinner and dark-mode toggle
- Currency-symbol preference (Settings page)

---

## 🛠️ Technologies Used

| Category | Technologies |
|---|---|
| Programming Language | Python 3.x |
| Web Framework | Flask 3.x |
| Frontend | HTML5, CSS3, Jinja2, JavaScript |
| Database | SQLite (13+ relational tables) |
| Data Visualization | Chart.js |
| Reporting / Export | openpyxl (Excel), python-docx (Word), ReportLab (PDF), csv |
| Security | Werkzeug password hashing, Flask secure sessions |
| Testing | pytest |
| AI Assistant | Custom rule-based NLP (Python `re`) — JARVIS |
| Package Management | pip |
| Version Control | Git, GitHub |
| Deployment | Gunicorn, Render (cloud hosting) |

---

## 📁 Project Structure

```
smart_finance/
│
├── app.py                       # Main Flask application (all routes)
├── finance.db                   # SQLite database (auto-created, gitignored)
├── requirements.txt             # Python dependencies
├── Procfile                     # Deployment start command (Gunicorn)
├── test_app.py                  # Automated pytest test suite
├── .gitignore
├── LICENSE
├── README.md
│
├── templates/
│   ├── login.html / register.html
│   ├── dashboard.html
│   ├── expenses.html / edit_expense.html
│   ├── budget.html / edit_budget.html
│   ├── investments.html / edit_investment.html
│   ├── goals.html / edit_goal.html
│   ├── analytics.html
│   ├── wallets.html / edit_wallet.html
│   ├── bills.html
│   ├── subscriptions.html / edit_subscription.html
│   ├── challenges.html
│   ├── spending_personality.html
│   ├── simulator.html
│   ├── shared_expenses.html
│   ├── bill_splitter.html
│   ├── reports.html
│   ├── report_expense.html / report_budget.html / report_investment.html / report_goals.html
│   ├── jarvis.html
│   ├── profile.html / settings.html
│   └── ai_analysis.html
│
└── static/
    └── style.css                 # Complete CSS styling incl. sidebar & themes
```

---

## 🔄 Application Workflow

1. 👤 User registers or logs in securely (hashed password, protected session).
2. 📊 The dashboard shows a live overview — income, expenses, savings, and budget usage.
3. 💵 User logs categorized expenses and sets category budgets.
4. 💼 User adds investments and financial goals; the app computes returns and AI-suggested monthly savings.
5. 📈 User explores Analytics, Reports, and the Spending Personality profile.
6. 🤖 User asks **JARVIS** natural-language questions about any part of their finances.
7. 📄 User exports reports (CSV, Excel, Word, or PDF) or prints them directly.

---

## 🤖 JARVIS — How It Works

```
User Message → Intent Detection (regex-based NLP) → Live Query Against
Expenses / Budget / Investments / Goals → Natural-Language Reply → Stored in Session History
```

JARVIS runs entirely on the application's own data — there is no external AI API, no API key, and no per-request cost. Every response is directly traceable to a real database query, which also makes its behavior easy to explain and debug.

---

## 🔐 Security & Privacy

- 🔑 Passwords are hashed with Werkzeug's `generate_password_hash` / `check_password_hash` — never stored in plain text.
- 🍪 Sessions are `HttpOnly` and `SameSite=Lax`, with a 2-hour automatic expiry.
- ✅ Server-side validation on registration and expense forms (required fields, valid email format, positive amounts).
- 🛡️ All SQL queries use parameterized statements — no string-concatenated SQL.
- 🔒 Every edit/delete route filters by both record ID **and** the logged-in `user_id`, so one account can never modify another's data.
- 🚫 The database file and Python cache are excluded from version control via `.gitignore`.

---

## 🚀 Deployment

The application is deployed on **Render** using **Gunicorn** as the production WSGI server.

The deployed application includes every module: expense/budget/investment/goal tracking, wallets, bills, subscriptions, challenges, spending personality, the What-If simulator, shared expenses, the bill splitter, advanced reports with PDF/Excel export, and the JARVIS assistant.

**Live URL:** [smart-finance-insights-dvjw.onrender.com](https://smart-finance-insights-dvjw.onrender.com)

> ⚠️ **Note:** Render's free tier uses an ephemeral filesystem, so the SQLite database resets on redeploy or after periods of inactivity. This is expected for a free-tier demo deployment; a production version would migrate to a persistent database such as PostgreSQL.

---

## 💻 Installation (Run Locally)

### 1️⃣ Clone the Repository
```bash
git clone https://github.com/meghanabodrapallim-coder/Finance-Analysis-Platform-For-Financial-Reporting-and-Budget-Tracking.git
```

### 2️⃣ Navigate to the Project
```bash
cd Finance-Analysis-Platform-For-Financial-Reporting-and-Budget-Tracking
```

### 3️⃣ Create a Virtual Environment
```bash
python -m venv venv
```

**Windows**
```bash
venv\Scripts\activate
```

**macOS / Linux**
```bash
source venv/bin/activate
```

### 4️⃣ Install Dependencies
```bash
pip install -r requirements.txt
```

### 5️⃣ Run the Application
```bash
python app.py
```

The application will be available at `http://127.0.0.1:5000`.

---

## 🗃️ Database

The application uses **SQLite**, with 13+ tables managing:

👤 Users · 💵 Income · 💸 Expenses · 🎯 Budgets · 💼 Investments · 🏆 Goals · 💳 Wallets · 📅 Bills · 🔄 Subscriptions · 🎯 Challenges & Challenge Logs · 🔔 Notifications · ⚙️ Settings

---

## 🧪 Testing

The project includes an automated **pytest** suite (`test_app.py`) covering:

- ✅ Registration and login (including wrong-password rejection)
- ✅ Dashboard access control (redirects unauthenticated users)
- ✅ Expense creation and deletion
- ✅ Budget creation and over-budget status detection
- ✅ Investment creation and profit calculation
- ✅ Goal creation, progress percentage, and deletion
- ✅ Logout

Run the suite with:
```bash
pytest test_app.py -v
```

---

## ⚠️ Limitations

- The platform does not connect directly to real bank accounts or payment providers.
- Currency preference changes the displayed symbol only — it does not perform live exchange-rate conversion.
- JARVIS uses rule-based intent matching rather than a full language model, so it understands a defined set of financial question patterns.
- SQLite on the free-tier deployment is not persistent across restarts.

---

## 🔮 Future Enhancements

- 🧾 OCR-based receipt scanning for automatic expense entry
- 🎙️ Voice-based expense logging
- 💱 Live multi-currency conversion with real exchange rates
- 🗺️ Expense location mapping
- 🐘 Migration to PostgreSQL for persistent production storage
- 🧠 Upgrading JARVIS to a full language-model-powered assistant

---

## 📚 Learning Outcomes

This project demonstrates practical, hands-on experience with:

- 🐍 Python programming and full-stack Flask web development
- 🗄️ Relational database design across 13+ interconnected SQLite tables
- 🎨 Frontend development with Jinja2 templating and Chart.js
- 🔐 Authentication security best practices (hashing, session hardening, input validation)
- 🤖 Building a rule-based NLP intent detector for a conversational assistant
- 📄 Multi-format report generation (PDF, Excel, Word, CSV)
- 🧪 Automated testing with pytest
- ☁️ End-to-end deployment: Git, GitHub, Gunicorn, and Render — including diagnosing and fixing a production-only bug

---

## 👩‍💻 Developer

**Meghana Bodrapalli**
🎓 Student & Developer — B.Tech, Artificial Intelligence & Data Science

💻 Python • Flask • SQLite • Chart.js

🔗 **GitHub:** [github.com/meghanabodrapallim-coder](https://github.com/meghanabodrapallim-coder)
🌐 **Live Project:** [Open Finance Analysis Platform](https://smart-finance-insights-dvjw.onrender.com)

---

## ⭐ Support

If you found this project useful, please consider giving it a ⭐ star on GitHub.

---

## 📜 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

**MIT License Permissions:** ✅ Commercial use · ✅ Modification · ✅ Distribution · ✅ Private use
**Conditions:** 📜 Preserve the copyright notice · 📜 Preserve the license notice

---

<p align="center">
💙 Made as part of the Infosys Springboard Virtual Internship — Smart Finance Insights
</p>