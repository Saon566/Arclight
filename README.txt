ARCLIGHT STREAMLIT DASHBOARD STARTER
===================================

Files
-----
app.py           Streamlit dashboard, charts, scenario controls, worksheet explorer
model_bridge.py Copies the Excel workbook, applies scenario inputs, asks LibreOffice to recalculate
requirements.txt Python dependencies

Setup (Windows / VS Code)
-------------------------
1. Place these files in the same folder as:
   Vyomaa_P3_P4_Capital_Raise_Model_v8(3).xlsx

2. Install dependencies in the VS Code terminal:
   python -m pip install -r requirements.txt

3. Check LibreOffice path:
   Test-Path "C:\Program Files\LibreOffice\program\soffice.exe"

4. Start app:
   python -m streamlit run app.py

Scenario inputs currently mapped to the workbook
------------------------------------------------
P3 inputs use column C and P4 inputs use column D:
- row 10: hard capex
- row 14: specific yield
- row 16: PPA tariff
- row 18: all-in opex
- row 28: maximum gearing (custom financing case)
- row 30: custom all-in interest rate
- C103: financing case selector (0 = custom)

The original workbook is not modified. Each scenario is applied to a copy and saved under
./recalculated/. Verify that LibreOffice recalculates all workbook formulas as expected.
Macros, Excel-only features, data tables, circular references and external links may need
separate testing.

IMPORTANT
---------
This is a development starter, not an investment-grade verified portal. Before sharing:
- Compare key outputs with the same scenario recalculated in Microsoft Excel.
- Review the Checks worksheet; require all checks to pass.
- Verify that the scenario assumptions match the intended investment case.
- Add authentication and hosting security before exposing confidential financial data.
- Do not rely on an unprotected URL as security.
