import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import SheetJS from "../app/vendor/xlsx.full.min.js";

const XLSX = SheetJS.default || SheetJS;

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const outputRoot = path.join(projectRoot, "output", "test-fixtures");

const MAIN_NS =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";

const sheetNames = [
  "离职名单",
  "备忘",
  "工资表",
  "工资核对表",
  "代发薪",
  "BM津贴",
];

const mainHeaders = [
  "序号",
  "身份证",
  "岗位",
  "人员编号",
  "姓名",
  "部门",
  "基本工资",
  "岗位工资",
  "绩效工资标准",
  "工资合计",
  "月度绩效（季度发放）",
  "BM津贴",
  "销售提成",
  "其他工资",
  "考勤扣款",
  "其他扣款",
  "应计工资",
  "之前月份累计应计工资",
  "累计应计工资",
  "代扣养老保险",
  "代扣医疗保险",
  "代扣失业保险",
  "代扣房积金",
  "之前月份累计代扣养老保险",
  "之前月份累计代扣医疗保险",
  "之前月份累计代扣失业保险",
  "之前月份累计代扣房积金",
  "累计代扣养老保险",
  "累计代扣医疗保险",
  "累计代扣失业保险",
  "累计代扣房积金",
  "累计子女教育",
  "累计继续教育",
  "累计住房贷款利息",
  "累计住房租金",
  "累计赡养老人",
  "累计婴幼儿照护费用",
  "累计个人养老金扣除",
  "基本扣除费用",
  "之前月份累计基本扣除费用",
  "累计基本扣除费用",
  "纳税工资",
  "代扣税",
  "之前月份累计代扣税",
  "本月应缴代扣税",
  "累计实缴代扣税",
  "税后工资",
  "年度绩效",
  "年度绩效代扣税",
  "代扣借款",
  "实发合计",
  "个人承担社保部分扣款合计",
  "企业承担社保部分扣款合计",
  "毕业院校",
  "学历",
  "入司时间",
  "TRS司龄",
  "首次参加工作时间",
  "工作年限",
];

if (mainHeaders.length !== 59) {
  throw new Error(`合成工资表字段数应为 59，当前为 ${mainHeaders.length}`);
}

const headerColumn = new Map(
  mainHeaders.map((header, index) => [header, index + 1]),
);

function xmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function columnLetters(number) {
  let cursor = Number(number);
  let output = "";
  while (cursor > 0) {
    cursor -= 1;
    output = String.fromCharCode(65 + (cursor % 26)) + output;
    cursor = Math.floor(cursor / 26);
  }
  return output;
}

function cellReference(column, row) {
  return `${columnLetters(column)}${row}`;
}

function inlineCell(column, row, value, style = 0) {
  const reference = cellReference(column, row);
  if (value === null || value === undefined || value === "") {
    return `<c r="${reference}" s="${style}"/>`;
  }
  return `<c r="${reference}" s="${style}" t="inlineStr"><is><t>${xmlEscape(
    value,
  )}</t></is></c>`;
}

function numberCell(column, row, value, style = 0) {
  return `<c r="${cellReference(column, row)}" s="${style}"><v>${Number(
    value,
  )}</v></c>`;
}

function formulaCell(column, row, formula, cachedValue = 0, style = 0) {
  return `<c r="${cellReference(column, row)}" s="${style}"><f>${xmlEscape(
    formula,
  )}</f><v>${Number(cachedValue)}</v></c>`;
}

function rowXml(rowNumber, cells, options = {}) {
  const hidden = options.hidden ? ' hidden="1"' : "";
  return `<row r="${rowNumber}" spans="1:${options.lastColumn || 12}"${hidden}>${cells.join(
    "",
  )}</row>`;
}

function worksheetXml(rows, lastColumn, lastRow) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">
  <dimension ref="A1:${cellReference(lastColumn, lastRow)}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="${lastColumn}" width="13" customWidth="1"/></cols>
  <sheetData>${rows.join("")}</sheetData>
  <pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>
</worksheet>`;
}

function simpleSheet(title, headers, dataRows, options = {}) {
  const lastColumn = headers.length;
  const rows = [
    rowXml(
      1,
      [
        inlineCell(1, 1, title, 2),
        ...headers.slice(1).map((_, index) =>
          inlineCell(index + 2, 1, "", 2),
        ),
      ],
      { lastColumn },
    ),
    rowXml(
      2,
      headers.map((header, index) => inlineCell(index + 1, 2, header, 1)),
      { lastColumn },
    ),
  ];
  dataRows.forEach((values, index) => {
    const rowNumber = index + 3;
    const cells = headers.map((header, columnIndex) => {
      const value = values[header];
      if (value && typeof value === "object" && value.formula) {
        return formulaCell(
          columnIndex + 1,
          rowNumber,
          value.formula,
          value.cached,
        );
      }
      return typeof value === "number"
        ? numberCell(columnIndex + 1, rowNumber, value)
        : inlineCell(columnIndex + 1, rowNumber, value ?? "");
    });
    rows.push(
      rowXml(rowNumber, cells, {
        lastColumn,
        hidden: options.hiddenRows?.includes(rowNumber),
      }),
    );
  });
  return worksheetXml(rows, lastColumn, Math.max(2, dataRows.length + 2));
}

function formulaForField(
  field,
  row,
  indices,
  january,
  person,
  year,
  month,
) {
  const column = (name) => columnLetters(headerColumn.get(name));
  const current = (name) => `${column(name)}${row}`;
  const previousMappings = new Map([
    ["之前月份累计应计工资", "累计应计工资"],
    ["之前月份累计代扣养老保险", "累计代扣养老保险"],
    ["之前月份累计代扣医疗保险", "累计代扣医疗保险"],
    ["之前月份累计代扣失业保险", "累计代扣失业保险"],
    ["之前月份累计代扣房积金", "累计代扣房积金"],
    ["之前月份累计基本扣除费用", "累计基本扣除费用"],
    ["之前月份累计代扣税", "累计实缴代扣税"],
  ]);
  if (previousMappings.has(field)) {
    if (january || !person.身份证) {
      return null;
    }
    const target = previousMappings.get(field);
    const lookupIndex =
      headerColumn.get(target) - headerColumn.get("身份证") + 1;
    return `VLOOKUP(${current("身份证")},[${indices.history}]工资表!$${column(
      "身份证",
    )}:$${column(target)},${lookupIndex},0)`;
  }
  const taxMappings = new Map([
    ["累计子女教育", ["M", 10]],
    ["累计继续教育", ["N", 11]],
    ["累计住房贷款利息", ["O", 12]],
    ["累计住房租金", ["P", 13]],
    ["累计赡养老人", ["Q", 14]],
    ["累计婴幼儿照护费用", ["R", 15]],
    ["累计个人养老金扣除", ["S", 16]],
  ]);
  if (taxMappings.has(field)) {
    if (!person.身份证) {
      return null;
    }
    const [sourceEnd, returnIndex] = taxMappings.get(field);
    return `VLOOKUP(${current("身份证")},[${indices.tax}]Sheet1!$D$2:$${sourceEnd}$100,${returnIndex},0)`;
  }
  if (field === "企业承担社保部分扣款合计") {
    return person.身份证
      ? `VLOOKUP(${current("身份证")},'[${indices.insurance}]${year}.${month}'!$E:$AE,27,0)`
      : null;
  }
  const formulas = {
    工资合计: `SUM(${current("基本工资")}:${current("其他工资")})`,
    应计工资: `SUM(${current("工资合计")}:${current("其他扣款")})`,
    累计应计工资: `${current("之前月份累计应计工资")}+${current("应计工资")}`,
    累计代扣养老保险: `${current("之前月份累计代扣养老保险")}+${current(
      "代扣养老保险",
    )}`,
    累计代扣医疗保险: `${current("之前月份累计代扣医疗保险")}+${current(
      "代扣医疗保险",
    )}`,
    累计代扣失业保险: `${current("之前月份累计代扣失业保险")}+${current(
      "代扣失业保险",
    )}`,
    累计代扣房积金: `${current("之前月份累计代扣房积金")}+${current(
      "代扣房积金",
    )}`,
    累计基本扣除费用: `${current("之前月份累计基本扣除费用")}+${current(
      "基本扣除费用",
    )}`,
    纳税工资: `${current("累计应计工资")}-${current(
      "累计基本扣除费用",
    )}`,
    代扣税: `${current("纳税工资")}*0.03`,
    本月应缴代扣税: `${current("代扣税")}-${current("之前月份累计代扣税")}`,
    累计实缴代扣税: `${current("之前月份累计代扣税")}+${current("本月应缴代扣税")}`,
    税后工资: `${current("应计工资")}-${current("本月应缴代扣税")}`,
    实发合计: `${current("税后工资")}-${current("代扣借款")}`,
    个人承担社保部分扣款合计: `SUM(${current("代扣养老保险")}:${current("代扣房积金")})`,
  };
  return formulas[field] || null;
}

function mainSheetXml(year, month, indices) {
  const january = month === 1;
  const periodText = `${year}年${month}月工资表`;
  const rows = [
    rowXml(
      1,
      [
        inlineCell(1, 1, "", 2),
        inlineCell(2, 1, periodText, 2),
        ...Array.from({ length: mainHeaders.length - 2 }, (_, index) =>
          inlineCell(index + 3, 1, "", 2),
        ),
      ],
      { lastColumn: mainHeaders.length },
    ),
    rowXml(
      2,
      mainHeaders.map((header, index) =>
        inlineCell(index + 1, 2, header, 1),
      ),
      { lastColumn: mainHeaders.length },
    ),
  ];
  const people = [
    {
      序号: 1,
      人员编号: "TEST-001",
      身份证: "SYNTHETIC-ID-001",
      姓名: "测试甲",
      部门: "测试部门",
      岗位: "测试岗位",
      入司时间: "2020-01-01",
      基本工资: 1800,
      岗位工资: 3000,
      绩效工资标准: 1200,
      其他工资: 0,
      工资卡号: "SYNTHETIC-ACCOUNT-001",
      银行: "测试银行",
      实发合计: 5600,
    },
    {
      序号: 2,
      人员编号: "TEST-002",
      身份证: "SYNTHETIC-ID-002",
      姓名: "测试乙",
      部门: "测试部门",
      岗位: "测试岗位",
      入司时间: "2021-02-01",
      基本工资: 4500,
      岗位工资: 800,
      其他工资: 100,
      工资卡号: "SYNTHETIC-ACCOUNT-002",
      银行: "测试银行",
      实发合计: 5000,
    },
    {
      序号: 3,
      人员编号: "",
      身份证: "",
      姓名: "测试劳务",
      部门: "测试部门",
      岗位: "临时服务",
      基本工资: 800,
      实发合计: 760,
    },
  ];
  for (const person of people) {
    person.BM津贴 = person.身份证 ? 200 : 0;
  }
  if (month === 11) {
    Object.assign(people[0], {
      "月度绩效（季度发放）": 900,
      销售提成: 300,
      其他工资: 150,
      考勤扣款: -50,
      其他扣款: -20,
      年度绩效: 600,
      代扣借款: 100,
    });
  }
  for (let personIndex = 0; personIndex < people.length; personIndex += 1) {
    const rowNumber = personIndex + 3;
    const person = people[personIndex];
    const cells = mainHeaders.map((field, columnIndex) => {
      const column = columnIndex + 1;
      const formula = formulaForField(
        field,
        rowNumber,
        indices,
        january,
        person,
        year,
        month,
      );
      if (formula) {
        const cached =
          field === "实发合计"
            ? person.实发合计
            : field.includes("累计")
              ? month * 100
              : 100;
        return formulaCell(column, rowNumber, formula, cached);
      }
      let value = person[field];
      if (
        value === undefined &&
        (/^之前月份累计/.test(field) || field === "基本扣除费用")
      ) {
        value = january ? 0 : 100;
      } else if (value === undefined && field === "税率") {
        value = 0.03;
      } else if (value === undefined && field === "发放状态") {
        value = "待复核";
      } else if (value === undefined) {
        value = 0;
      }
      return typeof value === "number"
        ? numberCell(column, rowNumber, value)
        : inlineCell(column, rowNumber, value);
    });
    rows.push(
      rowXml(rowNumber, cells, { lastColumn: mainHeaders.length }),
    );
  }
  rows.push(
    rowXml(
      6,
      mainHeaders.map((_, index) => inlineCell(index + 1, 6, "")),
      { lastColumn: mainHeaders.length },
    ),
  );
  return worksheetXml(rows, mainHeaders.length, 6);
}

function workbookXml(externalCount) {
  const externalReferences = externalCount
    ? `<externalReferences>${Array.from(
        { length: externalCount },
        (_, index) => `<externalReference r:id="rId${8 + index}"/>`,
      ).join("")}</externalReferences>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="22000" windowHeight="12000"/></bookViews>
  <sheets>${sheetNames
    .map(
      (name, index) =>
        `<sheet name="${xmlEscape(name)}" sheetId="${index + 1}" r:id="rId${
          index + 1
        }"/>`,
    )
    .join("")}</sheets>
  ${externalReferences}
  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>
</workbook>`;
}

function workbookRelationshipsXml(externalCount) {
  const sheetRelationships = sheetNames
    .map(
      (_, index) =>
        `<Relationship Id="rId${index + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${
          index + 1
        }.xml"/>`,
    )
    .join("");
  const externalRelationships = Array.from(
    { length: externalCount },
    (_, index) =>
      `<Relationship Id="rId${8 + index}" Type="${REL_NS}/externalLink" Target="externalLinks/externalLink${
        index + 1
      }.xml"/>`,
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  ${sheetRelationships}
  <Relationship Id="rId7" Type="${REL_NS}/styles" Target="styles.xml"/>
  ${externalRelationships}
</Relationships>`;
}

function contentTypesXml(externalCount) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetNames
    .map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${
          index + 1
        }.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join("")}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${Array.from(
    { length: externalCount },
    (_, index) =>
      `<Override PartName="/xl/externalLinks/externalLink${
        index + 1
      }.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`,
  ).join("")}
</Types>`;
}

const rootRelationshipsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}">
  <fonts count="3">
    <font><sz val="11"/><name val="Arial"/></font>
    <font><b/><sz val="10"/><name val="Arial"/><color rgb="FFFFFFFF"/></font>
    <font><b/><sz val="14"/><name val="Arial"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

function externalLinkXml(sheetName) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">
  <externalBook r:id="rId1"><sheetNames><sheetName val="${xmlEscape(
    sheetName,
  )}"/></sheetNames></externalBook>
</externalLink>`;
}

function externalRelationshipXml(target) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/externalLinkPath" Target="${xmlEscape(
    target,
  )}" TargetMode="External"/>
</Relationships>`;
}

function externalDefinitions(year, month) {
  const padded = String(month).padStart(2, "0");
  const tax = {
    target: `file:///Synthetic/工资附件/${month}月/${year}${padded}_正常工资薪金所得.xls`,
    sheet: "Sheet1",
  };
  const insurance = {
    target: `file:///Synthetic/工资附件/${month}月/${year}.${month}示例公司社保公积金表.xlsx`,
    sheet: `${year}.${month}`,
  };
  const auxiliary = {
    target: "file:///Synthetic/2025年3月工资核对.xlsx",
    sheet: "核对",
  };
  if (month === 1) {
    return [tax, insurance, auxiliary];
  }
  const previousMonth = month - 1;
  return [
    {
      target: `file:///Synthetic/${year}.${previousMonth}示例公司工资.xlsx`,
      sheet: "工资表",
    },
    tax,
    insurance,
    auxiliary,
  ];
}

function sourceIndices(month) {
  return month === 1
    ? { history: 0, tax: 1, insurance: 2, auxiliary: 3 }
    : { history: 1, tax: 2, insurance: 3, auxiliary: 4 };
}

function auxiliarySheets(indices, options = {}) {
  const identities = [
    {
      人员编号: "TEST-001",
      身份证: "SYNTHETIC-ID-001",
      姓名: "测试甲",
      部门: "测试部门",
      岗位: "测试岗位",
    },
    {
      人员编号: "TEST-002",
      身份证: "SYNTHETIC-ID-002",
      姓名: "测试乙",
      部门: "测试部门",
      岗位: "测试岗位",
    },
  ];
  const archiveRows = [
    {
      人员编号: "ARCHIVED-TEST",
      身份证: "SYNTHETIC-ARCHIVE-ID",
      姓名: "历史测试",
      部门: "测试部门",
      岗位: "历史岗位",
      入职日期: "2020-01-01",
      离职日期: "2024-12-31",
      备注: "合成夹具历史行",
    },
  ];
  const memoRows = identities.map((identity) => ({ ...identity }));
  if (options.effectiveDeparture) {
    const departure = {
      人员编号: "TEST-002",
      身份证: "SYNTHETIC-ID-002",
      姓名: "测试乙",
      部门: "测试部门",
      岗位: "测试岗位",
      入职日期: "2021-02-01",
      离职日期: options.effectiveDeparture,
      备注: "次月起转回合成总部",
    };
    archiveRows.push(departure);
    Object.assign(
      memoRows.find((row) => row.人员编号 === "TEST-002"),
      {
        离职日期: options.effectiveDeparture,
        备注: "分表已有明确离职记录",
      },
    );
  }
  return {
    离职名单: simpleSheet(
      "离职名单",
      [
        "人员编号",
        "身份证",
        "姓名",
        "部门",
        "岗位",
        "入职日期",
        "离职日期",
        "备注",
      ],
      archiveRows,
    ),
    备忘: simpleSheet(
      "备忘",
      [
        "人员编号",
        "身份证",
        "姓名",
        "部门",
        "岗位",
        "入司时间",
        "离职日期",
        "备注",
      ],
      memoRows,
    ),
    工资核对表: simpleSheet(
      "工资核对表",
      ["人员编号", "身份证", "姓名", "其他工资", "实发合计"],
      identities.map((identity, index) => ({
        ...identity,
        其他工资: {
          formula: `VLOOKUP($A${index + 3},工资表!$B:$AQ,13,0)`,
          cached: index,
        },
        实发合计: {
          formula: `VLOOKUP($A${index + 3},工资表!$B:$AQ,42,0)`,
          cached: index === 0 ? 5600 : 5000,
        },
      })),
    ),
    代发薪: simpleSheet(
      "代发薪",
      ["人员编号", "身份证", "账户名称(*)", "账号(*)", "金额(*)"],
      identities.map((identity, index) => ({
        人员编号: identity.人员编号,
        身份证: identity.身份证,
        "账户名称(*)": identity.姓名,
        "账号(*)": `SYNTHETIC-ACCOUNT-00${index + 1}`,
        "金额(*)": index === 0 ? 5600 : 5000,
      })),
    ),
    BM津贴: simpleSheet(
      "BM津贴",
      ["姓名", "BM津贴"],
      [
        {
          姓名: "测试甲",
          BM津贴: {
            formula: `[${indices.auxiliary}]核对!B3`,
            cached: 0,
          },
        },
      ],
    ),
  };
}

async function buildWorkbook(year, month, destination, options = {}) {
  const zip = new JSZip();
  const external = externalDefinitions(year, month);
  const indices = sourceIndices(month);
  const auxiliary = auxiliarySheets(indices, options);

  zip.file("[Content_Types].xml", contentTypesXml(external.length));
  zip.file("_rels/.rels", rootRelationshipsXml);
  zip.file("xl/workbook.xml", workbookXml(external.length));
  zip.file(
    "xl/_rels/workbook.xml.rels",
    workbookRelationshipsXml(external.length),
  );
  zip.file("xl/styles.xml", stylesXml);
  zip.file("xl/worksheets/sheet1.xml", auxiliary.离职名单);
  zip.file("xl/worksheets/sheet2.xml", auxiliary.备忘);
  zip.file("xl/worksheets/sheet3.xml", mainSheetXml(year, month, indices));
  zip.file("xl/worksheets/sheet4.xml", auxiliary.工资核对表);
  zip.file("xl/worksheets/sheet5.xml", auxiliary.代发薪);
  zip.file("xl/worksheets/sheet6.xml", auxiliary.BM津贴);
  external.forEach((definition, index) => {
    zip.file(
      `xl/externalLinks/externalLink${index + 1}.xml`,
      externalLinkXml(definition.sheet),
    );
    zip.file(
      `xl/externalLinks/_rels/externalLink${index + 1}.xml.rels`,
      externalRelationshipXml(definition.target),
    );
  });
  const bytes = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  await fs.writeFile(destination, bytes);
}

const syntheticPeople = [
  {
    employeeId: "TEST-001",
    idCard: "SYNTHETIC-ID-001",
    name: "测试甲",
  },
  {
    employeeId: "TEST-002",
    idCard: "SYNTHETIC-ID-002",
    name: "测试乙",
  },
];

async function writeSourceWorkbook(
  destination,
  sheetName,
  matrix,
  bookType,
) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(matrix),
    sheetName,
  );
  const bytes = XLSX.write(workbook, {
    type: "buffer",
    bookType,
  });
  await fs.writeFile(destination, bytes);
}

async function buildAttachments(
  year,
  month,
  destination,
  options = {},
) {
  const padded = String(month).padStart(2, "0");
  const attachmentPeople = options.people || syntheticPeople;
  await fs.mkdir(destination, { recursive: true });
  const taxHeaders = [
    "工号",
    "姓名",
    "证件类型",
    "证件号码",
    "所得期间起",
    "所得期间止",
    "本期收入",
    "本期免税收入",
    "基本养老保险费",
    "基本医疗保险费",
    "失业保险费",
    "住房公积金",
    "累计子女教育",
    "累计继续教育",
    "累计住房贷款利息",
    "累计住房租金",
    "累计赡养老人",
    "累计3岁以下婴幼儿照护",
    "累计个人养老金",
    "企业(职业)年金",
    "商业健康保险",
    "税延养老保险",
    "其他",
    "准予扣除的捐赠额",
    "税前扣除项目合计",
    "减免税额",
    "减除费用标准",
    "已缴税额",
    "备注",
  ];
  const taxRows = attachmentPeople.map((person, index) => [
    person.employeeId,
    person.name,
    "居民身份证",
    person.idCard,
    `${year}-${padded}-01`,
    `${year}-${padded}-28`,
    6000 - index * 500,
    0,
    100 + index,
    20 + index,
    10 + index,
    200 + index,
    month * 10 + index,
    month * 20 + index,
    month * 30 + index,
    month * 40 + index,
    month * 50 + index,
    month * 60 + index,
    month * 70 + index,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    5000,
    0,
    "",
  ]);
  await writeSourceWorkbook(
    path.join(
      destination,
      `${year}${padded}_正常工资薪金所得.xls`,
    ),
    "Sheet1",
    [taxHeaders, ...taxRows],
    "biff8",
  );
  if (year === 2025 && month === 12) {
    const unmatched = [...taxRows[0]];
    unmatched[0] = "TEST-999";
    unmatched[1] = "未匹配测试";
    unmatched[3] = "SYNTHETIC-ID-999";
    await writeSourceWorkbook(
      path.join(
        destination,
        `${year}${padded}_正常工资薪金所得_含未匹配.xls`,
      ),
      "Sheet1",
      [taxHeaders, ...taxRows, unmatched],
      "biff8",
    );
  }

  const insuranceHeaders = [
    "序列",
    "年份",
    "月份",
    "姓名",
    "身份证号",
    "人员类别",
    "缴纳地",
    "所属部门",
    "年缴",
    "五险一金基数",
    "长春续保上家基数",
    "哈尔滨续保社平基数",
    "养老基数",
    "失业基数",
    "工伤基数",
    "生育基数",
    "医疗基数",
    "公积金基数",
    "养老单位",
    "养老个人",
    "失业单位",
    "失业个人",
    "工伤单位",
    "医疗单位",
    "医疗个人",
    "公积金公司",
    "公积金个人",
    "个人补缴",
    "公司补缴",
    "个人合计",
    "公司合计",
  ];
  const insuranceRows = attachmentPeople.map((person, index) => {
    const pension = 100 + index;
    const unemployment = 10 + index;
    const medical = 20 + index;
    const housingFund = 200 + index;
    return [
      index + 1,
      year,
      month,
      person.name,
      person.idCard,
      "在职",
      "测试地",
      "测试部门",
      "",
      5000,
      "",
      "",
      5000,
      5000,
      5000,
      5000,
      5000,
      5000,
      800,
      pension,
      50,
      unemployment,
      20,
      400,
      medical,
      600,
      housingFund,
      0,
      0,
      pension + unemployment + medical + housingFund,
      1870,
    ];
  });
  await writeSourceWorkbook(
    path.join(
      destination,
      `${year}.${month}示例公司社保公积金表.xlsx`,
    ),
    `${year}.${month}`,
    [insuranceHeaders, ...insuranceRows],
    "xlsx",
  );

  if (year >= 2026) {
    await writeSourceWorkbook(
      path.join(destination, `${year}年劳务费-示例.xlsx`),
      `${year}.${month}`,
      [
        ["主体", "区域", "姓名", "劳务费", "增值税税额"],
        ["合成主体", "测试区域", "测试劳务", 800, 40],
      ],
      "xlsx",
    );
  }

  if (year >= 2026) {
    await writeSourceWorkbook(
      path.join(
        destination,
        `${year}年${month}月员工动态表-示例.xlsx`,
      ),
      "示例",
      [
        ["转正"],
        [
          "序号",
          "身份证号",
          "姓名",
          "部门",
          "岗位",
          "转正日期",
          "备注",
        ],
        [
          1,
          "SYNTHETIC-ID-001",
          "测试甲",
          "测试部门",
          "测试岗位",
          `${year}-${String(month).padStart(2, "0")}-01`,
          "合成非敏感转正记录",
        ],
        [],
        ["转岗"],
        [
          "序号",
          "身份证号",
          "姓名",
          "现部门",
          "现岗位",
          "转岗后部门",
          "转岗后岗位",
          "备注",
        ],
        [
          1,
          "SYNTHETIC-ID-001",
          "测试甲",
          "测试部门",
          "测试岗位",
          "",
          "新测试岗位",
          "合成非敏感变动",
        ],
        [],
        ["调薪"],
        [
          "序号",
          "身份证号",
          "姓名",
          "调薪日期",
          "备注",
        ],
        [
          1,
          "SYNTHETIC-ID-001",
          "测试甲",
          `${year}-${String(month).padStart(2, "0")}-01`,
          "合成非敏感调薪记录",
        ],
      ],
      "xlsx",
    );
    await writeSourceWorkbook(
      path.join(
        destination,
        `${year}年${month}月入职转正薪资-示例.xlsx`,
      ),
      "示例",
      [
        ["转正"],
        [
          "序号",
          "身份证号",
          "姓名",
          "转正日期",
          "部门",
          "岗位",
          "基本工资",
          "岗位工资",
          "绩效工资",
          "备注",
        ],
        [
          1,
          "SYNTHETIC-ID-001",
          "测试甲",
          `${year}-${String(month).padStart(2, "0")}-01`,
          "测试部门",
          "测试岗位",
          1800,
          3000,
          1200,
          "合成非敏感核对",
        ],
        [],
        ["转岗"],
        [
          "序号",
          "身份证号",
          "姓名",
          "现部门",
          "现岗位",
          "转部门日期",
          "转岗后部门",
          "转岗后岗位",
          "备注",
        ],
        [
          1,
          "SYNTHETIC-ID-001",
          "测试甲",
          "测试部门",
          "测试岗位",
          `${year}-${String(month).padStart(2, "0")}-01`,
          "",
          "新测试岗位",
          "合成非敏感转岗复核",
        ],
        [],
        ["调薪"],
        [
          "序号",
          "身份证号",
          "姓名",
          "调薪日期",
          "调整后薪酬",
          "备注",
        ],
        [
          1,
          "SYNTHETIC-ID-001",
          "测试甲",
          `${year}-${String(month).padStart(2, "0")}-01`,
          6000,
          "合成非敏感调薪",
        ],
      ],
      "xlsx",
    );
    await writeSourceWorkbook(
      path.join(
        destination,
        `${year}年${month}月实习生津贴表-示例.xlsx`,
      ),
      "示例",
      [
        ["年", "月", "姓名", "身份证号", "津贴总额", "备注"],
        [
          year,
          month,
          "测试乙",
          "SYNTHETIC-ID-002",
          4500,
          "合成非敏感核对",
        ],
      ],
      "xlsx",
    );
    await writeSourceWorkbook(
      path.join(
        destination,
        `${year}年${month}月行政请假记录-示例.xlsx`,
      ),
      "示例",
      [
        [
          "序号",
          "工作部门",
          "身份证号",
          "姓名",
          "请假信息",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
        ],
        [
          "",
          "",
          "",
          "",
          "年假",
          "婚假",
          "丧假",
          "带薪病假",
          "产检",
          "陪产假",
          "扣款病假",
          "产假",
          "事假",
          "无薪休息",
          "备注",
        ],
        [
          1,
          "",
          "SYNTHETIC-ID-001",
          "测试甲",
          1,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          "",
        ],
      ],
      "xlsx",
    );
  }
}

await fs.mkdir(outputRoot, { recursive: true });
await buildWorkbook(
  2025,
  11,
  path.join(outputRoot, "2025.11合成工资表.xlsx"),
);
const annualRoot = path.join(outputRoot, "2025全年");
await fs.mkdir(annualRoot, { recursive: true });
for (let month = 1; month <= 12; month += 1) {
  await buildWorkbook(
    2025,
    month,
    path.join(annualRoot, `2025.${month}合成工资表.xlsx`),
  );
}
await buildAttachments(
  2025,
  12,
  path.join(outputRoot, "2025.12工资附件"),
);
await buildAttachments(
  2026,
  1,
  path.join(outputRoot, "2026.01工资附件"),
);
await buildWorkbook(
  2025,
  11,
  path.join(outputRoot, "2025.11合成工资表-含分表变动.xlsx"),
  { effectiveDeparture: "2025-11-30" },
);
await buildAttachments(
  2025,
  12,
  path.join(outputRoot, "2025.12分表变动附件"),
  { people: syntheticPeople.slice(0, 1) },
);

console.log(`已生成 ${outputRoot} 下的无个人信息合成工资表夹具`);
