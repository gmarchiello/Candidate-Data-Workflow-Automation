# Candidate Data Workflow Automation with Python & Excel

![Python](https://img.shields.io/badge/python-3.10%2B-blue.svg)  
![License](https://img.shields.io/badge/license-MIT-lightgrey)  
[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/gmarchiello/pdf_form_filler/blob/main/pdf_form_filler_colab.ipynb)

---

## 🔍 About This Project
This repository contains a **mockup version** of a tool I originally developed at *Instituto Cervantes* to improve the workflow for updating **DELE candidate data**.

Previously, updating candidate information required **manually filling PDF forms** and sending them via JIRA to colleagues in Madrid — a process that was **time-consuming, repetitive, and prone to errors**, especially during peak exam periods.

To optimize this workflow, I developed an **end-to-end automation pipeline**: it handles everything from **data aggregation to PDF generation and ready-to-send JIRA messages**, reducing manual effort and errors while improving speed and flexibility.

---

## 🚀 Try It Online
You can test a **simplified demo version** directly in Google Colab — no installation needed:

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/gmarchiello/pdf_form_filler/blob/main/pdf_form_filler_colab.ipynb)

---

## ⚡ Key Features
- **Excel / Power Query integration** — consolidates candidate data and allows flexible filtering.  
- **Automated PDF filling (Python)** — fills text fields and checkboxes based on Excel input.  
- **Output management** — timestamped folders, safe filenames (`"UNKNOWN"` if missing), and batch-ready email summaries.  
- **Logging & error handling** — shows success ✔️, missing text ⚠️, or missing checkboxes ❗; continues processing without stopping.  
- **Flexible workflow** — fully automated for email changes; manual filtering supported for other updates.

---

## 🛠️ Handling Missing Data
- **Missing Name / Surname** — PDF field left empty; filename uses `UNKNOWN` (e.g. `UNKNOWN_John_change_request.pdf`).  
- **Missing text fields** — PDF left blank; a warning is logged.  
- **Missing checkboxes** — PDF still generated; an alert is logged.  
- **Success cases** — rows logged with ✔️.

This approach favors **speed and flexibility**, letting you quickly fix edge cases manually rather than stopping the whole run.

---

## 📊 Impact
- ⏱️ **Time savings:** ~93% per candidate *(manual ~7.5 min → automated ~0.5 min)*  
  *Estimated from average processing times per candidate.*  
- ✅ **Error reduction:** automated validation reduces mistakes  
- 📈 **Scalability:** handles high-volume sessions without bottlenecks  
- 💡 **Flexibility:** continues running even with incomplete data.

---

## 📂 Project Structure
```

project_root/
├─ src/
│  ├─ main.py          # Main execution pipeline
│  ├─ config.py        # Configuration and mappings
│  ├─ pdf_filler.py    # PDF form handling
│  └─ utils.py         # Helper functions
├─ input/
│  ├─ templates/       # PDF templates
│  └─ change_request_data/ # Excel source file
├─ output/             # Generated PDFs (auto-created)
└─ requirements.txt    # Project dependencies

````

---

## 📦 Requirements
- Python 3.10+  
- `pandas`  
- `openpyxl`  
- `pdfrw`  
- `pytz`  

Install all dependencies:

```bash
pip install -r requirements.txt
````

---

## 🚀 Usage

1. Place the Excel input file in `input/change_request_data/`
2. Place the PDF template in `input/templates/`
3. Run the script:

```bash
python src/main.py
```

4. Check `/output/` for generated PDFs.
5. Review console logs for email-ready summaries.

---

## 🧑‍💻 Skills Demonstrated

* Python scripting & automation
* Data pipelines (Excel → Python → PDF)
* Error handling & logging
* Workflow optimization & process improvement
* Practical use of `pandas`, `openpyxl`, `pdfrw`
