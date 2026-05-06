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

const calcRefs = {
  sex: "Calculator!$B$5",
  age: "Calculator!$B$6",
  tc: "Calculator!$B$7",
  hdl: "Calculator!$B$8",
  sbp: "Calculator!$B$9",
  diabetes: 'IF(Calculator!$B$10="Yes",1,0)',
  smoking: 'IF(Calculator!$B$11="Yes",1,0)',
  bmi: "Calculator!$B$12",
  egfr: "Calculator!$B$13",
  antihtn: 'IF(Calculator!$B$14="Yes",1,0)',
  statin: 'IF(Calculator!$B$15="Yes",1,0)',
  uacr: "Calculator!$B$16",
  hba1c: "Calculator!$B$17",
  sdi: "Calculator!$B$18",
};

const FORMULAS = {
  "Age per 10 years": `(${calcRefs.age}-55)/10`,
  "Age per 10 years squared": `((${calcRefs.age}-55)/10)^2`,
  "non-HDL-C per 1 mmol/L": `((${calcRefs.tc}-${calcRefs.hdl})*0.02586)-3.5`,
  "HDL-C per 0.3 mmol/L": `((${calcRefs.hdl}*0.02586)-1.3)/0.3`,
  "SBP <110 per 20 mmHg": `(MIN(${calcRefs.sbp},110)-110)/20`,
  "SBP ≥110 per 20 mmHg": `(MAX(${calcRefs.sbp},110)-130)/20`,
  Diabetes: calcRefs.diabetes,
  "Current smoking": calcRefs.smoking,
  "BMI <30, per 5 kg/m2": `(MIN(${calcRefs.bmi},30)-25)/5`,
  "BMI 30+, per 5 kg/m2": `(MAX(${calcRefs.bmi},30)-30)/5`,
  "eGFR <60, per -15 ml": `(MIN(${calcRefs.egfr},60)-60)/-15`,
  "eGFR 60+, per -15 ml": `(MAX(${calcRefs.egfr},60)-90)/-15`,
  "Anti-hypertensive use": calcRefs.antihtn,
  "Statin use": calcRefs.statin,
  "Treated SBP ≥110 mm Hg per 20 mm Hg": `((MAX(${calcRefs.sbp},110)-130)/20)*${calcRefs.antihtn}`,
  "Treated non-HDL-C": `((((${calcRefs.tc}-${calcRefs.hdl})*0.02586)-3.5))*${calcRefs.statin}`,
  "Age per 10yr * non-HDL-C per 1 mmol/L": `(((${calcRefs.age}-55)/10)*(((((${calcRefs.tc}-${calcRefs.hdl})*0.02586)-3.5))))`,
  "Age per 10yr * HDL-C per 0.3 mml/L": `(((${calcRefs.age}-55)/10)*(((${calcRefs.hdl}*0.02586)-1.3)/0.3))`,
  "Age per 10yr * SBP ≥110 mm Hg per 20 mmHg": `(((${calcRefs.age}-55)/10)*((MAX(${calcRefs.sbp},110)-130)/20))`,
  "Age per 10yr * diabetes": `(((${calcRefs.age}-55)/10)*${calcRefs.diabetes})`,
  "Age per 10yr * current smoking": `(((${calcRefs.age}-55)/10)*${calcRefs.smoking})`,
  "Age per 10yr * BMI 30+ per 5 kg/m2": `(((${calcRefs.age}-55)/10)*((MAX(${calcRefs.bmi},30)-30)/5))`,
  "Age per 10yr * eGFR <60, per -15 ml": `(((${calcRefs.age}-55)/10)*((MIN(${calcRefs.egfr},60)-60)/-15))`,
  "SDI decile categories 4-6 vs. 1-3": `IF(AND(${calcRefs.sdi}<>"",${calcRefs.sdi}>=4,${calcRefs.sdi}<=6),1,0)`,
  "SDI decile categories 7-10 vs. 1-3": `IF(AND(${calcRefs.sdi}<>"",${calcRefs.sdi}>=7,${calcRefs.sdi}<=10),1,0)`,
  "Missing SDI": `IF(${calcRefs.sdi}="",1,0)`,
  "ln-UACR, mg/g, per 1 ln unit": `IF(OR(${calcRefs.uacr}="",${calcRefs.uacr}=0),0,LN(${calcRefs.uacr}))`,
  "Missing UACR/PCR/Dipstick": `IF(${calcRefs.uacr}="",1,0)`,
  "HbA1c in DM, per 1%": `IF(${calcRefs.hba1c}="",0,(${calcRefs.hba1c}-5.3)*${calcRefs.diabetes})`,
  "HbA1c no DM, per 1%": `IF(${calcRefs.hba1c}="",0,(${calcRefs.hba1c}-5.3)*(1-${calcRefs.diabetes}))`,
  "Missing HbA1c": `IF(${calcRefs.hba1c}="",1,0)`,
  Constant: "1",
};

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
    views: [{ state: "frozen", xSplit: 3, ySplit: 4 }],
  });
  const sources = workbook.addWorksheet("Sources");
  const engine = workbook.addWorksheet("Engine");
  const coefficientSheet = workbook.addWorksheet("Coefficients");
  engine.state = "hidden";
  coefficientSheet.state = "hidden";

  calculator.columns = [
    { width: 34 },
    { width: 16 },
    { width: 18 },
    { width: 24 },
    { width: 16 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
    { width: 12 },
  ];

  calculator.mergeCells("A1:I1");
  calculator.getCell("A1").value = "AHA PREVENT Excel Calculator";
  calculator.getCell("A1").font = { bold: true, size: 16 };

  calculator.mergeCells("A2:I2");
  calculator.getCell("A2").value =
    "Implements the published simplified PREVENT equations from Circulation 149(6):430-449 (February 6, 2024), Supplemental Tables S12A-J.";
  calculator.getCell("A2").font = { italic: true, color: { argb: "FF555555" } };

  styleSectionTitle(calculator.getCell("A4"));
  calculator.getCell("A4").value = "Inputs";
  calculator.getCell("A5").value = "Sex";
  calculator.getCell("A6").value = "Age, years";
  calculator.getCell("A7").value = "Total cholesterol, mg/dL";
  calculator.getCell("A8").value = "HDL-C, mg/dL";
  calculator.getCell("A9").value = "Systolic BP, mmHg";
  calculator.getCell("A10").value = "Diabetes";
  calculator.getCell("A11").value = "Current smoking";
  calculator.getCell("A12").value = "BMI, kg/m2";
  calculator.getCell("A13").value = "eGFR, mL/min/1.73m2";
  calculator.getCell("A14").value = "Anti-hypertension medication";
  calculator.getCell("A15").value = "Statin use";
  calculator.getCell("A16").value = "UACR, mg/g (optional)";
  calculator.getCell("A17").value = "HbA1c, % (optional)";
  calculator.getCell("A18").value = "SDI decile 1-10 (optional)";

  calculator.getCell("B5").value = DEFAULT_INPUTS.sex;
  calculator.getCell("B6").value = DEFAULT_INPUTS.age;
  calculator.getCell("B7").value = DEFAULT_INPUTS.tc;
  calculator.getCell("B8").value = DEFAULT_INPUTS.hdl;
  calculator.getCell("B9").value = DEFAULT_INPUTS.sbp;
  calculator.getCell("B10").value = DEFAULT_INPUTS.diabetes ? "Yes" : "No";
  calculator.getCell("B11").value = DEFAULT_INPUTS.smoking ? "Yes" : "No";
  calculator.getCell("B12").value = DEFAULT_INPUTS.bmi;
  calculator.getCell("B13").value = DEFAULT_INPUTS.egfr;
  calculator.getCell("B14").value = DEFAULT_INPUTS.antihtn ? "Yes" : "No";
  calculator.getCell("B15").value = DEFAULT_INPUTS.statin ? "Yes" : "No";
  calculator.getCell("B16").value = DEFAULT_INPUTS.uacr;
  calculator.getCell("B17").value = DEFAULT_INPUTS.hba1c;
  calculator.getCell("B18").value = DEFAULT_INPUTS.sdi;

  for (const address of ["B5", "B10", "B11", "B14", "B15"]) {
    calculator.getCell(address).dataValidation = {
      type: "list",
      allowBlank: false,
      formulae: address === "B5" ? ['"Women,Men"'] : ['"No,Yes"'],
    };
  }

  calculator.getCell("B6").dataValidation = {
    type: "whole",
    operator: "between",
    allowBlank: false,
    formulae: [30, 79],
  };
  calculator.getCell("B7").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: false,
    formulae: [130, 320],
  };
  calculator.getCell("B8").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: false,
    formulae: [20, 100],
  };
  calculator.getCell("B9").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: false,
    formulae: [90, 200],
  };
  calculator.getCell("B12").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: false,
    formulae: [18.5, 39.999],
  };
  calculator.getCell("B13").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: false,
    formulae: [15, 150],
  };
  calculator.getCell("B16").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: true,
    formulae: [0.1, 25000],
  };
  calculator.getCell("B17").dataValidation = {
    type: "decimal",
    operator: "between",
    allowBlank: true,
    formulae: [3, 15],
  };
  calculator.getCell("B18").dataValidation = {
    type: "whole",
    operator: "between",
    allowBlank: true,
    formulae: [1, 10],
  };

  styleSectionTitle(calculator.getCell("D4"));
  calculator.getCell("D4").value = "10-year risk";
  styleSectionTitle(calculator.getCell("D13"));
  calculator.getCell("D13").value = "30-year risk";

  const modelHeaders = ["Base", "UACR", "HbA1c", "SDI", "Full"];
  const firstOutputColumn = 5;

  calculator.getCell("D5").value = "Outcome";
  calculator.getCell("D14").value = "Outcome";
  for (let index = 0; index < modelHeaders.length; index += 1) {
    calculator.getCell(5, firstOutputColumn + index).value = modelHeaders[index];
    calculator.getCell(14, firstOutputColumn + index).value = modelHeaders[index];
  }

  OUTCOMES.forEach((outcome, index) => {
    calculator.getCell(6 + index, 4).value = outcome;
    calculator.getCell(15 + index, 4).value = outcome;
  });

  for (const row of [5, 14]) {
    for (let col = 4; col <= 9; col += 1) {
      const cell = calculator.getCell(row, col);
      cell.font = { bold: true };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFEAF4FB" },
      };
      cell.border = {
        bottom: { style: "thin", color: { argb: "FFB8CFE1" } },
      };
    }
  }

  calculator.getCell("A20").value =
    "Use Base when optional predictors are unavailable. Use the UACR, HbA1c, or SDI columns for the corresponding single enhanced model. The Full column includes all three optional predictors and published missing-indicator handling.";
  calculator.mergeCells("A20:I20");
  calculator.getCell("A20").alignment = { wrapText: true };
  calculator.getCell("A20").font = { italic: true, color: { argb: "FF555555" } };

  calculator.getCell("A21").value =
    "The PREVENT paper excluded extreme values outside the ranges shown above and reports that Total CVD and ASCVD risks are less than the sum of component outcomes because a person can experience more than one event.";
  calculator.mergeCells("A21:I21");
  calculator.getCell("A21").alignment = { wrapText: true };
  calculator.getCell("A21").font = { italic: true, color: { argb: "FF555555" } };

  calculator.eachRow((row, rowNumber) => {
    if (rowNumber >= 5 && rowNumber <= 18) {
      row.eachCell((cell) => {
        cell.border = {
          bottom: { style: "hair", color: { argb: "FFE3E3E3" } },
        };
      });
    }
  });

  engine.columns = [{ width: 42 }, { width: 18 }];
  engine.getCell("A1").value = "Term";
  engine.getCell("B1").value = "Derived value";
  TERM_ORDER.forEach((term, index) => {
    const row = index + 2;
    engine.getCell(row, 1).value = term;
    engine.getCell(row, 2).value = { formula: FORMULAS[term] };
  });

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

  const engineRange = `$B$2:$B$${TERM_ORDER.length + 1}`;
  const derivedValues = computeDerived(DEFAULT_INPUTS);

  function outputFormula(model, horizon, outcome) {
    const womenKey = `${model}|${horizon}|${outcome}|Women`;
    const menKey = `${model}|${horizon}|${outcome}|Men`;
    const womenColumn = colLetter(coeffColumnByKey[womenKey]);
    const menColumn = colLetter(coeffColumnByKey[menKey]);
    const coeffRowStart = 2;
    const coeffRowEnd = TERM_ORDER.length + 1;
    const womenRange = `"Coefficients!$${womenColumn}$${coeffRowStart}:$${womenColumn}$${coeffRowEnd}"`;
    const menRange = `"Coefficients!$${menColumn}$${coeffRowStart}:$${menColumn}$${coeffRowEnd}"`;

    return `IF($B$5="Women",1/(1+EXP(-SUMPRODUCT(Engine!${engineRange},INDIRECT(${womenRange})))),1/(1+EXP(-SUMPRODUCT(Engine!${engineRange},INDIRECT(${menRange})))))`;
  }

  function outputResult(model, horizon, outcome) {
    const coefficients = modelCoefficients[`${model}|${horizon}`];
    return computeRisk(coefficients, derivedValues, outcome, DEFAULT_INPUTS.sex);
  }

  modelHeaders.forEach((model, modelIndex) => {
    OUTCOMES.forEach((outcome, outcomeIndex) => {
      const tenYearCell = calculator.getCell(6 + outcomeIndex, firstOutputColumn + modelIndex);
      tenYearCell.value = {
        formula: outputFormula(model, "10-year", outcome),
        result: outputResult(model, "10-year", outcome),
      };
      tenYearCell.numFmt = "0.0%";
      tenYearCell.alignment = { horizontal: "center" };

      const thirtyYearCell = calculator.getCell(
        15 + outcomeIndex,
        firstOutputColumn + modelIndex,
      );
      thirtyYearCell.value = {
        formula: outputFormula(model, "30-year", outcome),
        result: outputResult(model, "30-year", outcome),
      };
      thirtyYearCell.numFmt = "0.0%";
      thirtyYearCell.alignment = { horizontal: "center" };
    });
  });

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
