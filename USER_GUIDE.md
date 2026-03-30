# Skoda License Extractor

Extract license details and debit duties from Bill of Entry (BOE) PDFs into a clean, formatted Excel file with one click.

## Quick Start
* **Requirements:** Windows OS (No Python installation needed if using the compiled version).
* **Installation:** Simply locate the `boe_extractor_gui.exe` (or run the script) in your project folder.

## How to Use (Step-by-Step)
1. **Launch the App:** Double-click the application icon to open the Skoda License Extractor.
2. **Upload BOEs:** Click the **"Upload BOE PDFs"** button and select one or multiple PDF files from your computer.
3. **Monitor Progress:** The app will automatically start processing each file. You can see the extracted **BE Number**, **Date**, **License No**, and **Debit Duty** appearing in the table.
4. **Generate Report:** Once processing is finished, the **"Generate Excel"** button will turn red. Click it to save your final report.
5. **Reset:** Use the **"Clear All"** button if you need to start a fresh batch.

## Common Issues
* **Security Warning:** If Windows or internal security prevents the app from starting, click *"More Info"* and then *"Run Anyway"*. This is common for internal tools.
* **Missing Table:** If a BOE does not contain the *"F. LICENCE DETAILS"* section, the tool will mark it as "No License Table Found" and skip it.
* **File Permissions:** Ensure your target Excel file is not open in another program when you click "Generate Excel".

## Contact
For support or feature requests, please contact the IT Team.
