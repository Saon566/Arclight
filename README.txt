ARCLIGHT STREAMLIT CLOUD PATCH

Files:
- app.py: the app.py you uploaded, preserved as-is
- model_bridge.py: patched to locate LibreOffice on Linux and Windows
- requirements.txt: Python dependencies
- packages.txt: installs LibreOffice Calc on Streamlit Community Cloud

Deployment:
1. Open the exact GitHub repository and branch connected to your Streamlit app.
2. Replace app.py and model_bridge.py with these files.
3. Ensure requirements.txt and packages.txt are at the same directory level as app.py.
4. Ensure the Excel workbook named exactly:
   Vyomaa_P3_P4_Capital_Raise_Model_v8(3).xlsx
   is available in that same directory (do not commit confidential workbook to a public repository).
5. Commit the changes. Wait for Streamlit to rebuild, then reboot the app if needed.
6. If it still fails, open Streamlit Cloud app settings/logs and copy the first full traceback.

Important:
- This patch addresses the exact Windows-only LibreOffice lookup in the uploaded model_bridge.py.
- The app's financial cell mappings and model results have not been validated against the Excel workbook.
- Keep the workbook and investor data private; use an appropriately secured repository/app.
