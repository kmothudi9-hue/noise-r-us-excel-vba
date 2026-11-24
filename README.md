# Noise-R-Us: Excel VBA rental project

An early academic project by Kagiso Mothudi, revisited with AI assistance. The original work and its mistakes are preserved in `learning-journey/`. A replacement VBA module proposes a complete rental and invoice workflow.

**Status:** Code reviewed but **not compiled or tested in Microsoft Excel**. This is a learning project, not a production billing system. The assignment instructions and real prices were not found; all rates and any condition charge must be supplied by the operator. The condition-charge rule is an assumption, not a verified policy.

## What it does

- Collects a sound-system group, days, delivery choice and initial condition.
- On return, asks for the condition and agreed rates, then calculates rental, delivery, condition charges and total.
- Validates inputs and lets the user cancel before sheet data changes.

## Get started in Excel for Windows

1. If downloaded from GitHub, run `python build_template.py` to reconstruct `Noise_R_Us_Template.xlsx` (the file is stored as base64 for portability). Open `Noise_R_Us_Template.xlsx` and choose **Save As > Excel Macro-Enabled Workbook (.xlsm)**. The template itself contains no macros.
2. Press **Alt+F11**, select the new workbook in Project Explorer, and choose **File > Import File**. Select `Noise_R_Us_Completed.bas`.
3. In the VBA editor select **Debug > Compile VBAProject**. Save and return to Excel.
4. Press **Alt+F8**, run `GetInput` to enter a rental, then run `ReturnSoundSystem` when it is returned. Only enable macros from a source you trust.
5. Run the manual checks below before relying on it. If Excel flags an error, the project needs further debugging.

The template preserves the original sheet's layout. The macro uses the descriptions in G10:G12. A new rental replaces the previous displayed rental and invoice, after all inputs have been collected; keep separate copies for records. All amounts are entered in the same currency. No tax, deposit or discounts are applied.

## Assumptions

Condition codes are 1 (good), 2 (fair), 3 (bad). For a worsened condition the operator is asked to enter the agreed charge, including zero when appropriate; equal or improved condition yields no charge. The code assumes this workflow only: check the actual rental policy before billing. Defensive limits are 1–10,000 whole days and 0–1,000,000 per entered amount, to two decimal places. Use your Excel locale's decimal separator.

## Manual test plan (sample values, not actual rates)

- Group B, 3 days, delivery Y, condition 1; return condition 2, daily rate 120.50, delivery 50, charge 75: expect rental 361.50 and total 486.50.
- Group A, 2 days, delivery N, condition 2; return condition 1, daily rate 200: expect total 400 and zero delivery/condition charges.
- Invalid group, days, delivery, condition, negative amount and >2 decimal places should be rejected; cancel should leave the current values unchanged.
- Missing or invalid rental details should prevent an invoice. Unexpected sheet write errors can leave partial changes; check the sheet if an error occurs.

## Learning journey

The original code is retained in `learning-journey/` as an unfinished snapshot with personal student identifiers removed. It includes syntax errors and incomplete invoice logic and is **not intended to run**. The revised module addresses those issues but still needs Excel testing. The original workbook remains unchanged in Kagiso's private files.
