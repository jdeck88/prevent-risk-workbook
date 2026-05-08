# prevent-excel-calculator

Excel implementation of the AHA PREVENT cardiovascular risk calculator using the published simplified PREVENT equations.

## Get Started

1. Download [prevent_calculator.xlsx](https://raw.githubusercontent.com/jdeck88/prevent-risk-workbook/main/prevent_calculator.xlsx).
2. Open `prevent_calculator.xlsx` in Microsoft Excel.
3. Go to the `Calculator` sheet.
4. Enter values in column `B`.
5. Read the 10-year and 30-year risk estimates shown on the right.

If you are viewing the repository on GitHub, open the file `prevent_calculator.xlsx` and download it to your computer.

## What to Enter

Enter the following values as available. For `Sex`, use `Female` or `Male`.

- Sex
- Age
- Total cholesterol
- HDL cholesterol
- Systolic blood pressure
- Diabetes
- Current smoking
- BMI
- eGFR
- Anti-hypertension medication use
- Statin use
- UACR, optional
- HbA1c, optional
- SDI decile, optional

## Which Output Column to Use

- `Base`: use standard clinical variables only
- `UACR`: use when UACR is available
- `HbA1c`: use when HbA1c is available
- `SDI`: use when SDI is available
- `Full`: use when all optional variables are available

## If an Input Is Uncertain

If a variable does not fit cleanly into one of the workbook choices:

- Do not guess if you can avoid it.
- Use the best documented value from the chart, intake form, or patient interview.
- If you still are not sure, leave optional fields blank rather than forcing a value.
- For required yes/no fields, use the most defensible current clinical classification and document your assumption outside the workbook.

Examples:

- Smoking:
  If the patient is clearly a current smoker, use `Yes`.
  If the patient clearly does not currently smoke, use `No`.
  If the status is uncertain, clarify before using the result for decision-making.

- Diabetes, statin use, or anti-hypertension medication:
  Use the patient's current status at the time you are estimating risk.
  If the medication was prescribed but it is unclear whether the patient is actually taking it, confirm before relying on the estimate.

- Lab values such as cholesterol, eGFR, UACR, or HbA1c:
  Use the most recent reliable value available.
  If an optional lab is missing, leave it blank and use the corresponding simpler model column.

- SDI:
  If you do not have a valid SDI decile, leave it blank and do not use the SDI or Full column as your primary result.

Practical rule:

- If an input is uncertain, run the calculator using only values you trust.
- If needed, test more than one plausible value and treat the result as a sensitivity check rather than a final answer.

## Notes

- The workbook reproduces the published simplified PREVENT equations in Excel format.
- It is not the official AHA web calculator.

## For Maintainers

If you need to rebuild the workbook from source files:

```bash
npm install
npm run build
```

## Source

- PMC article: `https://pmc.ncbi.nlm.nih.gov/articles/PMC10910659/`
- PubMed: `https://pubmed.ncbi.nlm.nih.gov/37947085/`
