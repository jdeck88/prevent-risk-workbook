import ExcelJS from "exceljs";
import XLSX from "xlsx";

const SOURCE_WORKBOOK = "prevent_supplemental_tables_real.xlsx";
const OUTPUT_WORKBOOK = "prevent_calculator.xlsx";

const OUTCOMES = [
  "Total CVD",
  "ASCVD",
  "Heart Failure",
  "Coronary Heart Disease",
  "Stroke",
];

const SEXES = ["Women", "Men"];
const MAX_CALCULATOR_ROWS = 1000;
const CALCULATOR_HEADER_ROW = 5;
const CALCULATOR_DATA_START_ROW = 6;

const MODEL_SHEETS = [
  { model: "Base", horizon: "10-year", sheet: "Table S12A Base 10yr" },
  { model: "UACR", horizon: "10-year", sheet: "Table S12B ACR 10yr" },
  { model: "HbA1c", horizon: "10-year", sheet: "Table S12C A1c 10yr" },
  { model: "SDI", horizon: "10-year", sheet: "Table S12D SDI 10yr" },
  { model: "Full", horizon: "10-year", sheet: "Table S12E Full 10yr" },
  { model: "Base", horizon: "30-year", sheet: "Table S12F Base 30yr" },
  { model: "UACR", horizon: "30-year", sheet: "Table S12G ACR 30yr" },
  { model: "HbA1c", horizon: "30-year", sheet: "Table S12H A1c 30yr" },
  { model: "SDI", horizon: "30-year", sheet: "Table S12I SDI 30yr" },
  { model: "Full", horizon: "30-year", sheet: "Table S12J Full 30yr" },
];

const COLUMN_MAP = [
  { col: 2, outcome: "Total CVD", sex: "Women" },
  { col: 3, outcome: "Total CVD", sex: "Men" },
  { col: 4, outcome: "ASCVD", sex: "Women" },
  { col: 5, outcome: "ASCVD", sex: "Men" },
  { col: 6, outcome: "Heart Failure", sex: "Women" },
  { col: 7, outcome: "Heart Failure", sex: "Men" },
  { col: 8, outcome: "Coronary Heart Disease", sex: "Women" },
  { col: 9, outcome: "Coronary Heart Disease", sex: "Men" },
  { col: 10, outcome: "Stroke", sex: "Women" },
  { col: 11, outcome: "Stroke", sex: "Men" },
];

const TERM_ORDER = [
  "Age per 10 years",
  "Age per 10 years squared",
  "non-HDL-C per 1 mmol/L",
  "HDL-C per 0.3 mmol/L",
  "SBP <110 per 20 mmHg",
  "SBP ≥110 per 20 mmHg",
  "Diabetes",
  "Current smoking",
  "BMI <30, per 5 kg/m2",
  "BMI 30+, per 5 kg/m2",
  "eGFR <60, per -15 ml",
  "eGFR 60+, per -15 ml",
  "Anti-hypertensive use",
  "Statin use",
  "Treated SBP ≥110 mm Hg per 20 mm Hg",
  "Treated non-HDL-C",
  "Age per 10yr * non-HDL-C per 1 mmol/L",
  "Age per 10yr * HDL-C per 0.3 mml/L",
  "Age per 10yr * SBP ≥110 mm Hg per 20 mmHg",
  "Age per 10yr * diabetes",
  "Age per 10yr * current smoking",
  "Age per 10yr * BMI 30+ per 5 kg/m2",
  "Age per 10yr * eGFR <60, per -15 ml",
  "SDI decile categories 4-6 vs. 1-3",
  "SDI decile categories 7-10 vs. 1-3",
  "Missing SDI",
  "ln-UACR, mg/g, per 1 ln unit",
  "Missing UACR/PCR/Dipstick",
  "HbA1c in DM, per 1%",
  "HbA1c no DM, per 1%",
  "Missing HbA1c",
  "Constant",
];

function calcRefs(row) {
  return {
    sex: `Calculator!$B${row}`,
    age: `Calculator!$C${row}`,
    tc: `Calculator!$D${row}`,
    hdl: `Calculator!$E${row}`,
    sbp: `Calculator!$F${row}`,
    diabetes: `IF(Calculator!$G${row}="Yes",1,0)`,
    smoking: `IF(Calculator!$H${row}="Yes",1,0)`,
    bmi: `Calculator!$I${row}`,
    egfr: `Calculator!$J${row}`,
    antihtn: `IF(Calculator!$K${row}="Yes",1,0)`,
    statin: `IF(Calculator!$L${row}="Yes",1,0)`,
    uacr: `Calculator!$M${row}`,
    hba1c: `Calculator!$N${row}`,
    sdi: `Calculator!$O${row}`,
  };
}

function termFormula(term, row) {
  const refs = calcRefs(row);
  const formulas = {
    "Age per 10 years": `(${refs.age}-55)/10`,
    "Age per 10 years squared": `((${refs.age}-55)/10)^2`,
    "non-HDL-C per 1 mmol/L": `((${refs.tc}-${refs.hdl})*0.02586)-3.5`,
    "HDL-C per 0.3 mmol/L": `((${refs.hdl}*0.02586)-1.3)/0.3`,
    "SBP <110 per 20 mmHg": `(MIN(${refs.sbp},110)-110)/20`,
    "SBP ≥110 per 20 mmHg": `(MAX(${refs.sbp},110)-130)/20`,
    Diabetes: refs.diabetes,
    "Current smoking": refs.smoking,
    "BMI <30, per 5 kg/m2": `(MIN(${refs.bmi},30)-25)/5`,
    "BMI 30+, per 5 kg/m2": `(MAX(${refs.bmi},30)-30)/5`,
    "eGFR <60, per -15 ml": `(MIN(${refs.egfr},60)-60)/-15`,
    "eGFR 60+, per -15 ml": `(MAX(${refs.egfr},60)-90)/-15`,
    "Anti-hypertensive use": refs.antihtn,
    "Statin use": refs.statin,
    "Treated SBP ≥110 mm Hg per 20 mm Hg": `((MAX(${refs.sbp},110)-130)/20)*${refs.antihtn}`,
    "Treated non-HDL-C": `((((${refs.tc}-${refs.hdl})*0.02586)-3.5))*${refs.statin}`,
    "Age per 10yr * non-HDL-C per 1 mmol/L": `(((${refs.age}-55)/10)*(((((${refs.tc}-${refs.hdl})*0.02586)-3.5))))`,
    "Age per 10yr * HDL-C per 0.3 mml/L": `(((${refs.age}-55)/10)*(((${refs.hdl}*0.02586)-1.3)/0.3))`,
    "Age per 10yr * SBP ≥110 mm Hg per 20 mmHg": `(((${refs.age}-55)/10)*((MAX(${refs.sbp},110)-130)/20))`,
    "Age per 10yr * diabetes": `(((${refs.age}-55)/10)*${refs.diabetes})`,
    "Age per 10yr * current smoking": `(((${refs.age}-55)/10)*${refs.smoking})`,
    "Age per 10yr * BMI 30+ per 5 kg/m2": `(((${refs.age}-55)/10)*((MAX(${refs.bmi},30)-30)/5))`,
    "Age per 10yr * eGFR <60, per -15 ml": `(((${refs.age}-55)/10)*((MIN(${refs.egfr},60)-60)/-15))`,
    "SDI decile categories 4-6 vs. 1-3": `IF(AND(${refs.sdi}<>"",${refs.sdi}>=4,${refs.sdi}<=6),1,0)`,
    "SDI decile categories 7-10 vs. 1-3": `IF(AND(${refs.sdi}<>"",${refs.sdi}>=7,${refs.sdi}<=10),1,0)`,
    "Missing SDI": `IF(${refs.sdi}="",1,0)`,
    "ln-UACR, mg/g, per 1 ln unit": `IF(OR(${refs.uacr}="",${refs.uacr}=0),0,LN(${refs.uacr}))`,
    "Missing UACR/PCR/Dipstick": `IF(${refs.uacr}="",1,0)`,
    "HbA1c in DM, per 1%": `IF(${refs.hba1c}="",0,(${refs.hba1c}-5.3)*${refs.diabetes})`,
    "HbA1c no DM, per 1%": `IF(${refs.hba1c}="",0,(${refs.hba1c}-5.3)*(1-${refs.diabetes}))`,
    "Missing HbA1c": `IF(${refs.hba1c}="",1,0)`,
    Constant: "1",
  };

  return formulas[term];
}

const DEFAULT_INPUTS = {
  sex: "Women",
  age: 50,
  tc: 200,
  hdl: 45,
  sbp: 160,
  diabetes: 1,
  smoking: 0,
  bmi: 35,
  egfr: 90,
  antihtn: 1,
  statin: 0,
  uacr: 40,
  hba1c: 7.5,
  sdi: 8,
};

function normalizeTerm(term) {
  return String(term || "")
    .trim()
    .replace("ln-ACR, mg/g, per 1 ln unit", "ln-UACR, mg/g, per 1 ln unit")
    .replace("Missing ACR/PCR/Dipstick", "Missing UACR/PCR/Dipstick");
}

function colLetter(index) {
  let current = index;
  let result = "";
  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

function extractCoefficients(sourceWorkbook, sheetName) {
  const worksheet = sourceWorkbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: true,
    defval: "",
  });
  const riskFactorsRow = rows.findIndex((row) => row[0] === "Risk Factors");
  if (riskFactorsRow === -1) {
    throw new Error(`Could not find coefficient table in "${sheetName}".`);
  }

  const coefficients = {};
  for (let rowIndex = riskFactorsRow + 2; rowIndex < rows.length; rowIndex += 1) {
    const rawTerm = rows[rowIndex][0];
    const term = normalizeTerm(rawTerm);
    if (!term) {
      continue;
    }

    coefficients[term] = {};
    for (const mapping of COLUMN_MAP) {
      const rawValue = rows[rowIndex][mapping.col];
      const numericValue = rawValue === "" || rawValue == null ? 0 : Number(rawValue);
      coefficients[term][`${mapping.outcome}|${mapping.sex}`] = numericValue;
    }

    if (term === "Constant") {
      break;
    }
  }

  return coefficients;
}

function extractExampleRisks(sourceWorkbook, sheetName) {
  const worksheet = sourceWorkbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: true,
    defval: "",
  });
  const rowIndex = rows.findIndex((row, index) => index > 20 && row[0] === "Risk");
  if (rowIndex === -1) {
    throw new Error(`Could not find example risk row in "${sheetName}".`);
  }
  const risks = {};
  for (const mapping of COLUMN_MAP) {
    risks[`${mapping.outcome}|${mapping.sex}`] = Number(rows[rowIndex][mapping.col]);
  }
  return risks;
}

function computeDerived(inputs) {
  const age10 = (inputs.age - 55) / 10;
  const nonHdl = (inputs.tc - inputs.hdl) * 0.02586 - 3.5;
  const hdl03 = (inputs.hdl * 0.02586 - 1.3) / 0.3;
  const sbpLt110 = (Math.min(inputs.sbp, 110) - 110) / 20;
  const sbpGe110 = (Math.max(inputs.sbp, 110) - 130) / 20;
  const bmiLt30 = (Math.min(inputs.bmi, 30) - 25) / 5;
  const bmiGe30 = (Math.max(inputs.bmi, 30) - 30) / 5;
  const egfrLt60 = (Math.min(inputs.egfr, 60) - 60) / -15;
  const egfrGe60 = (Math.max(inputs.egfr, 60) - 90) / -15;
  const sdiCat46 = inputs.sdi >= 4 && inputs.sdi <= 6 ? 1 : 0;
  const sdiCat710 = inputs.sdi >= 7 && inputs.sdi <= 10 ? 1 : 0;

  return {
    "Age per 10 years": age10,
    "Age per 10 years squared": age10 ** 2,
    "non-HDL-C per 1 mmol/L": nonHdl,
    "HDL-C per 0.3 mmol/L": hdl03,
    "SBP <110 per 20 mmHg": sbpLt110,
    "SBP ≥110 per 20 mmHg": sbpGe110,
    Diabetes: inputs.diabetes,
    "Current smoking": inputs.smoking,
    "BMI <30, per 5 kg/m2": bmiLt30,
    "BMI 30+, per 5 kg/m2": bmiGe30,
    "eGFR <60, per -15 ml": egfrLt60,
    "eGFR 60+, per -15 ml": egfrGe60,
    "Anti-hypertensive use": inputs.antihtn,
    "Statin use": inputs.statin,
    "Treated SBP ≥110 mm Hg per 20 mm Hg": sbpGe110 * inputs.antihtn,
    "Treated non-HDL-C": nonHdl * inputs.statin,
    "Age per 10yr * non-HDL-C per 1 mmol/L": age10 * nonHdl,
    "Age per 10yr * HDL-C per 0.3 mml/L": age10 * hdl03,
    "Age per 10yr * SBP ≥110 mm Hg per 20 mmHg": age10 * sbpGe110,
    "Age per 10yr * diabetes": age10 * inputs.diabetes,
    "Age per 10yr * current smoking": age10 * inputs.smoking,
    "Age per 10yr * BMI 30+ per 5 kg/m2": age10 * bmiGe30,
    "Age per 10yr * eGFR <60, per -15 ml": age10 * egfrLt60,
    "SDI decile categories 4-6 vs. 1-3": sdiCat46,
    "SDI decile categories 7-10 vs. 1-3": sdiCat710,
    "Missing SDI": 0,
    "ln-UACR, mg/g, per 1 ln unit": Math.log(inputs.uacr),
    "Missing UACR/PCR/Dipstick": 0,
    "HbA1c in DM, per 1%": (inputs.hba1c - 5.3) * inputs.diabetes,
    "HbA1c no DM, per 1%": (inputs.hba1c - 5.3) * (1 - inputs.diabetes),
    "Missing HbA1c": 0,
    Constant: 1,
  };
}

function computeRisk(coefficients, derivedValues, outcome, sex) {
  const key = `${outcome}|${sex}`;
  let logOdds = 0;
  for (const term of TERM_ORDER) {
    const coefficient = coefficients[term]?.[key] || 0;
    logOdds += coefficient * (derivedValues[term] || 0);
  }
  return 1 / (1 + Math.exp(-logOdds));
}

function assertExampleMatches(modelSpecs, modelCoefficients, exampleRisks) {
  const derivedValues = computeDerived(DEFAULT_INPUTS);
  const tolerance = 1e-9;

  for (const spec of modelSpecs) {
    const coeffs = modelCoefficients[`${spec.model}|${spec.horizon}`];
    const expected = exampleRisks[`${spec.model}|${spec.horizon}`];
    for (const outcome of OUTCOMES) {
      for (const sex of SEXES) {
        const actualRisk = computeRisk(coeffs, derivedValues, outcome, sex);
        const expectedRisk = expected[`${outcome}|${sex}`];
        if (Math.abs(actualRisk - expectedRisk) > tolerance) {
          throw new Error(
            `Validation failed for ${spec.model} ${spec.horizon} ${outcome} ${sex}: ` +
              `${actualRisk} !== ${expectedRisk}`,
          );
        }
      }
    }
  }
}

function styleSectionTitle(cell) {
  cell.font = { bold: true, size: 12 };
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9EAF7" },
  };
}

async function main() {
  const sourceWorkbook = XLSX.readFile(SOURCE_WORKBOOK);

  const modelCoefficients = {};
  const exampleRisks = {};
  for (const spec of MODEL_SHEETS) {
    modelCoefficients[`${spec.model}|${spec.horizon}`] = extractCoefficients(
      sourceWorkbook,
      spec.sheet,
    );
    exampleRisks[`${spec.model}|${spec.horizon}`] = extractExampleRisks(
      sourceWorkbook,
      spec.sheet,
    );
  }

  assertExampleMatches(MODEL_SHEETS, modelCoefficients, exampleRisks);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "OpenAI Codex";
  workbook.company = "OpenAI";
  workbook.created = new Date("2026-05-06T00:00:00-07:00");
  workbook.modified = new Date("2026-05-06T00:00:00-07:00");
  workbook.calcProperties.fullCalcOnLoad = true;

  const calculator = workbook.addWorksheet("Calculator", {
    views: [{ state: "frozen", xSplit: 2, ySplit: CALCULATOR_HEADER_ROW }],
  });
  const sources = workbook.addWorksheet("Sources");
  const engine = workbook.addWorksheet("Engine");
  const coefficientSheet = workbook.addWorksheet("Coefficients");
  engine.state = "hidden";
  coefficientSheet.state = "hidden";

  const inputColumns = [
    { key: "record", header: "Scenario", width: 18 },
    { key: "sex", header: "Sex", width: 10 },
    { key: "age", header: "Age", width: 9 },
    { key: "tc", header: "Total chol", width: 12 },
    { key: "hdl", header: "HDL-C", width: 10 },
    { key: "sbp", header: "SBP", width: 10 },
    { key: "diabetes", header: "Diabetes", width: 11 },
    { key: "smoking", header: "Smoking", width: 11 },
    { key: "bmi", header: "BMI", width: 10 },
    { key: "egfr", header: "eGFR", width: 10 },
    { key: "antihtn", header: "Anti-HTN", width: 11 },
    { key: "statin", header: "Statin", width: 10 },
    { key: "uacr", header: "UACR", width: 10 },
    { key: "hba1c", header: "HbA1c", width: 10 },
    { key: "sdi", header: "SDI", width: 9 },
  ];
  const modelHeaders = ["Base", "UACR", "HbA1c", "SDI", "Full"];
  const outputStartColumn = inputColumns.length + 1;
  const outputColumnsPerHorizon = modelHeaders.length * OUTCOMES.length;
  const totalColumnCount = inputColumns.length + outputColumnsPerHorizon * 2;
  const lastColumnLetter = colLetter(totalColumnCount);

  calculator.columns = [
    ...inputColumns.map((column) => ({ width: column.width })),
    ...Array.from({ length: outputColumnsPerHorizon * 2 }, () => ({ width: 11 })),
  ];

  calculator.mergeCells(`A1:${lastColumnLetter}1`);
  calculator.getCell("A1").value = "AHA PREVENT Excel Calculator";
  calculator.getCell("A1").font = { bold: true, size: 16 };

  calculator.mergeCells(`A2:${lastColumnLetter}2`);
  calculator.getCell("A2").value =
    "Implements the published simplified PREVENT equations from Circulation 149(6):430-449 (February 6, 2024), Supplemental Tables S12A-J.";
  calculator.getCell("A2").font = { italic: true, color: { argb: "FF555555" } };

  calculator.mergeCells(`A3:O3`);
  calculator.getCell("A3").value =
    "Enter one patient or scenario per row starting on row 6. Optional variables are UACR, HbA1c, and SDI. Rows with missing required inputs stay blank.";
  calculator.getCell("A3").alignment = { wrapText: true };
  calculator.getCell("A3").font = { italic: true, color: { argb: "FF555555" } };
  styleSectionTitle(calculator.getCell("A4"));
  calculator.getCell("A4").value = "Inputs";

  const tenYearStartColumn = outputStartColumn;
  const thirtyYearStartColumn = outputStartColumn + outputColumnsPerHorizon;

  calculator.mergeCells(`P3:${colLetter(thirtyYearStartColumn - 1)}3`);
  calculator.getCell("P3").value = "10-year risk";
  styleSectionTitle(calculator.getCell("P3"));

  calculator.mergeCells(
    `${colLetter(thirtyYearStartColumn)}3:${lastColumnLetter}3`,
  );
  calculator.getCell(`${colLetter(thirtyYearStartColumn)}3`).value = "30-year risk";
  styleSectionTitle(calculator.getCell(`${colLetter(thirtyYearStartColumn)}3`));

  calculator.getRow(CALCULATOR_HEADER_ROW).height = 32;
  inputColumns.forEach((column, index) => {
    const cell = calculator.getCell(CALCULATOR_HEADER_ROW, index + 1);
    cell.value = column.header;
    cell.font = { bold: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFEAF4FB" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  function applyOutputHeader(startColumn) {
    modelHeaders.forEach((model, modelIndex) => {
      const blockStart = startColumn + modelIndex * OUTCOMES.length;
      const blockEnd = blockStart + OUTCOMES.length - 1;
      calculator.mergeCells(
        `${colLetter(blockStart)}4:${colLetter(blockEnd)}4`,
      );
      const titleCell = calculator.getCell(4, blockStart);
      titleCell.value = model;
      styleSectionTitle(titleCell);
      titleCell.alignment = { horizontal: "center" };
      OUTCOMES.forEach((outcome, outcomeIndex) => {
        const headerCell = calculator.getCell(
          CALCULATOR_HEADER_ROW,
          blockStart + outcomeIndex,
        );
        headerCell.value = outcome;
        headerCell.font = { bold: true };
        headerCell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFEAF4FB" },
        };
        headerCell.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
      });
    });
  }

  applyOutputHeader(tenYearStartColumn);
  applyOutputHeader(thirtyYearStartColumn);

  calculator.autoFilter = {
    from: { row: CALCULATOR_HEADER_ROW, column: 1 },
    to: { row: CALCULATOR_HEADER_ROW, column: totalColumnCount },
  };

  const defaultRow = CALCULATOR_DATA_START_ROW;
  calculator.getCell(`A${defaultRow}`).value = "Example";
  calculator.getCell(`B${defaultRow}`).value = DEFAULT_INPUTS.sex;
  calculator.getCell(`C${defaultRow}`).value = DEFAULT_INPUTS.age;
  calculator.getCell(`D${defaultRow}`).value = DEFAULT_INPUTS.tc;
  calculator.getCell(`E${defaultRow}`).value = DEFAULT_INPUTS.hdl;
  calculator.getCell(`F${defaultRow}`).value = DEFAULT_INPUTS.sbp;
  calculator.getCell(`G${defaultRow}`).value = DEFAULT_INPUTS.diabetes ? "Yes" : "No";
  calculator.getCell(`H${defaultRow}`).value = DEFAULT_INPUTS.smoking ? "Yes" : "No";
  calculator.getCell(`I${defaultRow}`).value = DEFAULT_INPUTS.bmi;
  calculator.getCell(`J${defaultRow}`).value = DEFAULT_INPUTS.egfr;
  calculator.getCell(`K${defaultRow}`).value = DEFAULT_INPUTS.antihtn ? "Yes" : "No";
  calculator.getCell(`L${defaultRow}`).value = DEFAULT_INPUTS.statin ? "Yes" : "No";
  calculator.getCell(`M${defaultRow}`).value = DEFAULT_INPUTS.uacr;
  calculator.getCell(`N${defaultRow}`).value = DEFAULT_INPUTS.hba1c;
  calculator.getCell(`O${defaultRow}`).value = DEFAULT_INPUTS.sdi;

  const dataEndRow = CALCULATOR_DATA_START_ROW + MAX_CALCULATOR_ROWS - 1;
  for (let row = CALCULATOR_DATA_START_ROW; row <= dataEndRow; row += 1) {
    for (const column of ["B", "G", "H", "K", "L"]) {
      calculator.getCell(`${column}${row}`).dataValidation = {
        type: "list",
        allowBlank: column !== "B",
        formulae: column === "B" ? ['"Women,Men"'] : ['"No,Yes"'],
      };
    }
    calculator.getCell(`C${row}`).dataValidation = {
      type: "whole",
      operator: "between",
      allowBlank: true,
      formulae: [30, 79],
    };
    calculator.getCell(`D${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [130, 320],
    };
    calculator.getCell(`E${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [20, 100],
    };
    calculator.getCell(`F${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [90, 200],
    };
    calculator.getCell(`I${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [18.5, 39.999],
    };
    calculator.getCell(`J${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [15, 150],
    };
    calculator.getCell(`M${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [0.1, 25000],
    };
    calculator.getCell(`N${row}`).dataValidation = {
      type: "decimal",
      operator: "between",
      allowBlank: true,
      formulae: [3, 15],
    };
    calculator.getCell(`O${row}`).dataValidation = {
      type: "whole",
      operator: "between",
      allowBlank: true,
      formulae: [1, 10],
    };
  }

  for (let row = CALCULATOR_DATA_START_ROW; row <= dataEndRow; row += 1) {
    for (let col = 1; col <= totalColumnCount; col += 1) {
      const cell = calculator.getCell(row, col);
      cell.border = {
        bottom: { style: "hair", color: { argb: "FFE3E3E3" } },
      };
      if (col <= inputColumns.length) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: row === defaultRow ? "FFFDF8E3" : "FFFFFFFF" },
        };
      }
    }
  }

  engine.columns = [{ width: 14 }, { width: 42 }, { width: 18 }];
  engine.getCell("A1").value = "Calc row";
  engine.getCell("B1").value = "Term";
  engine.getCell("C1").value = "Derived value";

  function engineBlockStart(calculatorRow) {
    return 2 + (calculatorRow - CALCULATOR_DATA_START_ROW) * TERM_ORDER.length;
  }

  for (let calculatorRow = CALCULATOR_DATA_START_ROW; calculatorRow <= dataEndRow; calculatorRow += 1) {
    const blockStart = engineBlockStart(calculatorRow);
    TERM_ORDER.forEach((term, index) => {
      const row = blockStart + index;
      engine.getCell(row, 1).value = calculatorRow;
      engine.getCell(row, 2).value = term;
      engine.getCell(row, 3).value = { formula: termFormula(term, calculatorRow) };
    });
  }

  coefficientSheet.getCell("A1").value = "Term";
  const coeffColumnByKey = {};
  let coeffColumnIndex = 2;

  for (const spec of MODEL_SHEETS) {
    const coefficients = modelCoefficients[`${spec.model}|${spec.horizon}`];
    for (const outcome of OUTCOMES) {
      for (const sex of SEXES) {
        const key = `${spec.model}|${spec.horizon}|${outcome}|${sex}`;
        coeffColumnByKey[key] = coeffColumnIndex;
        coefficientSheet.getCell(1, coeffColumnIndex).value = key;
        coeffColumnIndex += 1;
      }
    }

    for (let rowIndex = 0; rowIndex < TERM_ORDER.length; rowIndex += 1) {
      const row = rowIndex + 2;
      coefficientSheet.getCell(row, 1).value = TERM_ORDER[rowIndex];
      for (const outcome of OUTCOMES) {
        for (const sex of SEXES) {
          const key = `${spec.model}|${spec.horizon}|${outcome}|${sex}`;
          const value = coefficients[TERM_ORDER[rowIndex]]?.[`${outcome}|${sex}`] || 0;
          coefficientSheet.getCell(row, coeffColumnByKey[key]).value = value;
        }
      }
    }
  }

  const derivedValues = computeDerived(DEFAULT_INPUTS);

  function rowReadyFormula(row) {
    return `OR($B${row}="",$C${row}="",$D${row}="",$E${row}="",$F${row}="",$G${row}="",$H${row}="",$I${row}="",$J${row}="",$K${row}="",$L${row}="")`;
  }

  function outputFormula(model, horizon, outcome, row) {
    const womenKey = `${model}|${horizon}|${outcome}|Women`;
    const menKey = `${model}|${horizon}|${outcome}|Men`;
    const womenColumn = colLetter(coeffColumnByKey[womenKey]);
    const menColumn = colLetter(coeffColumnByKey[menKey]);
    const coeffRowStart = 2;
    const coeffRowEnd = TERM_ORDER.length + 1;
    const engineStart = engineBlockStart(row);
    const engineEnd = engineStart + TERM_ORDER.length - 1;
    const engineRange = `Engine!$C$${engineStart}:$C$${engineEnd}`;
    const womenRange = `"Coefficients!$${womenColumn}$${coeffRowStart}:$${womenColumn}$${coeffRowEnd}"`;
    const menRange = `"Coefficients!$${menColumn}$${coeffRowStart}:$${menColumn}$${coeffRowEnd}"`;

    return `IF(${rowReadyFormula(row)},"",IF($B${row}="Women",1/(1+EXP(-SUMPRODUCT(${engineRange},INDIRECT(${womenRange})))),1/(1+EXP(-SUMPRODUCT(${engineRange},INDIRECT(${menRange}))))))`;
  }

  function outputResult(model, horizon, outcome) {
    const coefficients = modelCoefficients[`${model}|${horizon}`];
    return computeRisk(coefficients, derivedValues, outcome, DEFAULT_INPUTS.sex);
  }

  for (let row = CALCULATOR_DATA_START_ROW; row <= dataEndRow; row += 1) {
    modelHeaders.forEach((model, modelIndex) => {
      OUTCOMES.forEach((outcome, outcomeIndex) => {
        const tenYearCell = calculator.getCell(
          row,
          tenYearStartColumn + modelIndex * OUTCOMES.length + outcomeIndex,
        );
        tenYearCell.value = {
          formula: outputFormula(model, "10-year", outcome, row),
          result: row === defaultRow ? outputResult(model, "10-year", outcome) : null,
        };
        tenYearCell.numFmt = "0.0%";
        tenYearCell.alignment = { horizontal: "center" };

        const thirtyYearCell = calculator.getCell(
          row,
          thirtyYearStartColumn + modelIndex * OUTCOMES.length + outcomeIndex,
        );
        thirtyYearCell.value = {
          formula: outputFormula(model, "30-year", outcome, row),
          result: row === defaultRow ? outputResult(model, "30-year", outcome) : null,
        };
        thirtyYearCell.numFmt = "0.0%";
        thirtyYearCell.alignment = { horizontal: "center" };
      });
    });
  }

  sources.columns = [{ width: 26 }, { width: 110 }];
  sources.getCell("A1").value = "Source";
  sources.getCell("B1").value = "Details";
  sources.getCell("A2").value = "Paper";
  sources.getCell("B2").value =
    "Khan SS, Matsushita K, Sang Y, et al. Development and Validation of the American Heart Association's PREVENT Equations. Circulation. 2024;149(6):430-449. PMID 37947085.";
  sources.getCell("A3").value = "PMC article";
  sources.getCell("B3").value = "https://pmc.ncbi.nlm.nih.gov/articles/PMC10910659/";
  sources.getCell("A4").value = "PubMed";
  sources.getCell("B4").value = "https://pubmed.ncbi.nlm.nih.gov/37947085/";
  sources.getCell("A5").value = "Supplemental tables";
  sources.getCell("B5").value =
    "https://pmc.ncbi.nlm.nih.gov/articles/instance/10910659/bin/NIHMS1953934-supplement-Supplemental_Tables.xlsx";
  sources.getCell("A6").value = "Workbook note";
  sources.getCell("B6").value =
    "This file uses the published simplified regression equations from Supplemental Tables S12A-J. Default inputs are the paper's worked example values.";

  for (const cell of ["B3", "B4", "B5"]) {
    const address = sources.getCell(cell).address;
    sources.getCell(address).font = { color: { argb: "FF0563C1" }, underline: true };
  }
  for (const row of [1]) {
    sources.getRow(row).font = { bold: true };
  }

  await workbook.xlsx.writeFile(OUTPUT_WORKBOOK);
  console.log(`Wrote ${OUTPUT_WORKBOOK}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
