# Yield Dashboard

A client-side production analytics dashboard built to extract, analyze, and visualize manufacturing yield and defect data directly from structured Excel reports.

This project demonstrates real-world data parsing, dynamic UI rendering, ranking algorithms, and chart visualization using pure JavaScript — without a backend.

---

## 🧠 Problem It Solves

Manufacturing teams often rely on manually reviewing Excel production reports to:

- Monitor process yield
- Identify recurring defects
- Track weekly defect trends
- Compare performance across processes

This dashboard automates that workflow by transforming structured Excel reports into an interactive analytics interface.

---
## 📌 Project Scope

This system is designed to support manufacturing performance monitoring and yield analysis across multiple operational units.

The dashboard can be utilized by:

- **Production Department** – Monitor daily yield performance, track process efficiency, and identify output gaps.
- **Quality Assurance (QA) Department** – Analyze defect trends, evaluate NG% rates, and monitor quality stability.
- **Engineering Department** – Perform root cause analysis, track process improvement initiatives, and evaluate process yield performance.

The tool is intended for internal operational analysis and decision support within manufacturing environments.

---

## 🚀 Deployment

This dashboard was deployed for internal use within the manufacturing environment.

It was used by:
- Production Department
- QA Department
- Engineering Department

The system was integrated into the local network environment and utilized for operational monitoring and yield tracking.

---

## ✨ Key Capabilities

### 📊 Process Analytics
- Automatically detects process blocks
- Extracts:
  - Target Yield
  - Actual Yield
  - Total Input
  - Total Output
  - Rejected Quantity

### 🏆 Defect Intelligence
- Weekly Top 5 defects (Week 1–5)
- Overall Top 10 defects across all weeks
- Automatic NG% computation
- Ranked defect aggregation

### 📈 Trend Visualization
- Interactive defect trend line chart
- Dynamic Top 1–10 switching
- Real-time chart updates (no reload)

### 📂 Smart Excel Parsing
- SheetJS-powered Excel parsing
- Dynamic sheet switching
- Structured block detection
- Normalized label matching

---

## 🛠 Tech Stack

- **Frontend:** HTML5, CSS3
- **Logic:** Vanilla JavaScript (ES6)
- **Excel Parsing:** SheetJS
- **Visualization:** Chart.js
- **File Access:** File System Access API

No frameworks. No backend. Fully client-side.

---

## 🏗 Architecture Overview

Excel File
↓
SheetJS Parsing
↓
Process Block Detection
↓
Yield & Metric Extraction
↓
Defect Aggregation Engine
↓
UI Rendering + Chart.js Visualization


### Core Concepts Demonstrated

- Structured data parsing from semi-tabular Excel formats
- Dynamic DOM rendering
- Ranking & sorting algorithms
- Defensive string normalization
- State management without frameworks
- Chart lifecycle management
- Event rebinding and UI state synchronization

---

## 📁 Project Structure
```bash
yield-dashboard/
│
├── index.html # Application layout
├── styles.css # UI styling
├── script.js # Data parsing & business logic
├── setup.vbs # Windows launcher helper
└── README.md
```

---

## 🚀 How To Run

1. Open `index.html` in a Chromium-based browser.
2. Click **Open File**.
3. Select a structured production Excel file.
4. Choose a sheet.
5. Click a process to view full analytics.

---

## 📊 Excel Structure Assumptions

The system expects structured production reports containing:

- `Process` blocks
- `Yield%` row
- `Total Input`
- `Total Output`
- `Total Quantity Rejected`
- `Target Yield`
- `MACHINING DEFECT` section
- Weekly defect columns (BO → BS)

The parsing engine relies on consistent column positioning and naming conventions.

---

## 🔒 Design Decisions

- Fully client-side to ensure:
  - Data privacy
  - Zero server dependency
  - Offline capability

- No framework used intentionally to:
  - Demonstrate core JavaScript competency
  - Maintain lightweight performance
  - Show direct DOM control

---

## ⚠ Limitations

- Strongly coupled to Excel structure
- Requires Chromium browser (File System Access API)
- Large Excel files may impact performance

---

## 📌 Future Improvements

- Abstract Excel column mappings
- Add schema validation
- Convert to Electron for desktop deployment
- Add PDF export
- Add process comparison view
- Add authentication for production environments

---

## 👨‍💻 Author

Built as an internal manufacturing analytics tool and as a demonstration of:

- Real-world data processing
- UI state management
- Business logic implementation
- Visualization integration

---

## 📷 Preview

Screenshots and demo preview will be added in a future update.

---

## 📜 License

Internal / Private Use


