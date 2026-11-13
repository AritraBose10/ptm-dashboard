# PTM Student Performance & Attendance Analytics Dashboard

An interactive, production-ready **Parent-Teacher Meeting (PTM)** analytics dashboard built with **Streamlit** and **Plotly**.

---

## 🚀 Quick Start

Run the application using:

```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

---

## 🔄 Dynamic & 100% Reusable (Zero Hardcoding)

This dashboard is designed to adapt to **any batch, semester, or academic year** with different Excel sheets in similar formats:

1. **Auto-Discovery**: If no files are manually uploaded, it automatically detects and loads any matching Excel sheets located in the directory.
2. **File Uploader in Sidebar**: You can drag and drop new **Marks Sheets** or **Attendance Sheets** directly in the sidebar anytime.
3. **Dynamic Subject Discovery**: Automatically identifies subject columns and course codes without relying on fixed lists.
4. **Dynamic Max Marks**: Scans header cells for `"Out of <N>"` (e.g., 1000, 500, 100) or automatically scales based on detected components.
5. **Dynamic Section Extraction**: Dynamically reads class sections from sheet tabs (e.g. `AIML 2A`, `CC 2A`, `DS 2B`, etc.).
6. **Dynamic Header & Metadata Mapping**: Intelligently identifies student IDs, student names, phone numbers, and lecture dates regardless of column positions.
7. **Customizable Policy Thresholds**: Easily adjust the **Debarment Attendance Cutoff** (e.g., 75%) and **Passing Marks Cutoff** (e.g., 40%) via sidebar sliders.

---

## 🎯 Features Included

### 1. Student PTM Dossier (Single-Student View)
* **Student Overview Card**: Roll/Registration ID, Name, Section, Stream, and **Parent Contact Number**.
* **Key KPI Cards**: Overall Attendance %, Grand Total Marks, Section Rank, and Batch Rank.
* **Faculty Talking Points**: Auto-generated smart bullet points identifying the student's best-performing subject, area of concern, and attendance standing.
* **Attendance Recovery Simulator ("What-If" Calculator)**:
  * Shows exact number of **consecutive classes** needed to reach 75% attendance.
  * Shows safe buffer classes if the student is already in good standing.
* **Attendance Behavioral Analytics**:
  * **Day-of-Week Attendance Chart**: Spotlights which day of the week the student misses classes most frequently.
  * **Monthly Progression Trend**: Identifies if attendance is improving, stable, or declining over the semester.
* **Subject Performance vs. Section Average**: Comparative Plotly grouped bar chart.
* **Detailed Component Breakdown**: End-Sem (80), Internal Assessment (20), Assignment (7), Quiz/Viva (8), and Practical scores.
* **Parent Communication Utilities**:
  * **1-Click WhatsApp Summary**: Formatted PTM message ready to send directly to the parent's phone.
  * **Meeting Remarks**: Text area for faculty to record agreed action items.

### 2. Section & Cohort Analytics
* Class average marks, class average attendance, and defaulter rates.
* Attendance and marks distribution histograms.
* Attendance vs. Academic Performance correlation scatter plot with trendline.

### 3. Defaulter Action List
* Prioritized table of all students below 75% attendance or failing scores.
* Includes parent contact numbers for bulk follow-up.
* **1-Click CSV Download** for department reporting.


---

## How the data is read (rules applied)

* **The attendance workbook is the source of truth** for attendance %. The `Attendance % till <date>` column of the marks sheet is only a fallback for students missing from the workbook; the date in that header is used as the **cut-off date** for counting lectures.
* From the attendance workbook only **lectures that actually happened** are counted: columns after the cut-off date, and columns where nobody was present (cancelled / not yet held), are ignored. Cells such as `No class`, `GROUP A`, `Session` are not counted for that student.
* Roll numbers are matched on their digits, so `261011041004 (lateral)` or IDs with hidden spaces still link to the marks sheet.
* Students with **no marks** show as unavailable (never as 0), and are excluded from marks averages and ranks. Students marked **DROP OUT** are flagged and left out of the intervention list.
* If the marks-sheet % differs from the workbook count by more than 5 points, the student's dossier shows both; the workbook figure is used.
* The sidebar **Data health** panel summarises matches, missing records and discrepancies for the loaded files.
* Branding: the university logo lives in `assets/logo_header.png` (source: `assets/logo.pdf`).

---

## Mobile layout

The same app adapts to phones automatically (screens narrower than ~760px):

* A **bottom tab bar** (Overview, Student, Directory, Subjects, Actions) replaces the top tabs.
* **Filters** collapse into a panel (closed by default on phones), laid out two per row.
* The **Directory** and **Actions** lists switch from wide tables to **student cards** with *Open dossier* and *WhatsApp* buttons.
* The student dossier has **WhatsApp / Call parent** buttons right under the profile card.
* Buttons and inputs get larger touch targets; charts no longer capture one-finger scrolling.
* **Policy rules** (attendance requirement and pass mark) live in a panel at the top of the page, not the sidebar.

---

## Deploying (Streamlit Community Cloud)

1. Push this repository to GitHub (the `.gitignore` keeps all Excel files and PTM records out of it).
2. Go to <https://share.streamlit.io> → **Create app** → pick this repo, branch `main`, main file `app.py`.
3. In the app's **Settings → Sharing**, restrict viewing to specific email addresses (the app handles student data).
4. Staff open the app, upload the **marks** and **attendance** workbooks from the sidebar, and work as usual.

PTM records (status, remarks, faculty name, dates) are saved on the server in `ptm_records.json` and shared by all faculty, so only give access to faculty. Excel workbooks are never stored: they live only in the uploader for that session.

**Important:** on free Streamlit Community Cloud the server's files are wiped when the app restarts or is redeployed. Use **Export PTM log (CSV)** in the sidebar regularly as a backup and **Restore PTM log** to reload it. For permanent storage, connect a database or Google Sheet (not set up yet).
