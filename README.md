# Investment Banking Pitch Kit

**Automated Comparable Company Analysis (Comps) Tool**

## The Problem & Solution
Investment Banking analysts spend hours manually formatting data in PowerPoint. This tool automates the **Comparable Company Analysis** workflow, turning raw financial data into client-ready slides in **<2 seconds**.

## Key Features
* **Automated Valuation:** Calculates **EV/EBITDA** & **EV/Revenue** multiples instantly.
* **Smart Formatting:** Applies industry-standard formatting (e.g., $100.5M, 12.5x) automatically.
* **Editable Output:** Generates native .pptx tables (via python-pptx), not static screenshots.

## Tech Stack
* **Python** (Pandas for financial logic)
* **python-pptx** (Slide automation)

## Quick Start
1. Clone the repository
   `git clone https://github.com/Mateusz-Olejniczak/ib_pitch_kit.git`
2. Install dependencies
   `pip install -r requirements.txt`
3. Run the script
   `python generate_pitch.py`

**Output:** US_Software_Comps_v1.pptx generated automatically.

---
