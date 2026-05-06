# prevent-excel-calculator

Excel implementation of the AHA PREVENT cardiovascular risk calculator using the published simplified PREVENT equations.

## Get Started

1. Download `prevent_calculator.xlsx` from this GitHub repository.
2. Open `prevent_calculator.xlsx` in Microsoft Excel.
3. Go to the `Calculator` sheet.
4. Enter values in column `B`.
5. Read the 10-year and 30-year risk estimates shown on the right.

If you are viewing the repository on GitHub, open the file `prevent_calculator.xlsx` and download it to your computer.

## What to Enter

Enter the following values as available:

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
