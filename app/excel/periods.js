(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function periodParts(period) {
    const match = String(period || "").match(/^(\d{4})-(0?[1-9]|1[0-2])$/);
    return match
      ? { year: match[1], month: String(Number(match[2])) }
      : null;
  }

  function periodMatchers(period) {
    const parts = periodParts(period);
    if (!parts) {
      return [];
    }
    const paddedMonth = parts.month.padStart(2, "0");
    return [
      new RegExp(`${parts.year}\\s*年\\s*0?${parts.month}\\s*月`, "g"),
      new RegExp(`${parts.year}\\s*\\.\\s*${paddedMonth}(?!\\d)`, "g"),
      new RegExp(`${parts.year}\\s*\\.\\s*${parts.month}(?!\\d)`, "g"),
      new RegExp(`${parts.year}\\s*-\\s*${paddedMonth}(?!\\d)`, "g"),
      new RegExp(`${parts.year}\\s*\\/\\s*${paddedMonth}(?!\\d)`, "g"),
    ];
  }

  function replacePeriodText(text, fromPeriod, toPeriod) {
    const from = periodParts(fromPeriod);
    const to = periodParts(toPeriod);
    if (!from || !to) {
      return String(text ?? "");
    }
    let output = String(text ?? "");
    const fromPadded = from.month.padStart(2, "0");
    const toPadded = to.month.padStart(2, "0");
    const replacements = [
      [
        new RegExp(`${from.year}\\s*年\\s*0?${from.month}\\s*月`, "g"),
        `${to.year}年${to.month}月`,
      ],
      [
        new RegExp(`${from.year}\\s*\\.\\s*${fromPadded}(?!\\d)`, "g"),
        `${to.year}.${toPadded}`,
      ],
      [
        new RegExp(`${from.year}\\s*\\.\\s*${from.month}(?!\\d)`, "g"),
        `${to.year}.${to.month}`,
      ],
      [
        new RegExp(`${from.year}\\s*-\\s*${fromPadded}(?!\\d)`, "g"),
        `${to.year}-${toPadded}`,
      ],
      [
        new RegExp(`${from.year}\\s*\\/\\s*${fromPadded}(?!\\d)`, "g"),
        `${to.year}/${toPadded}`,
      ],
    ];
    for (const [matcher, replacement] of replacements) {
      output = output.replace(matcher, replacement);
    }
    return output;
  }

  function periodsFromText(value) {
    const text = String(value ?? "");
    const periods = new Set();
    const patterns = [
      /((?:19|20)\d{2})\s*年\s*(0?[1-9]|1[0-2])\s*月/g,
      /((?:19|20)\d{2})\s*[./-]\s*(0?[1-9]|1[0-2])(?!\d)/g,
    ];
    for (const pattern of patterns) {
      for (const match of text.matchAll(pattern)) {
        periods.add(
          `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`,
        );
      }
    }
    return [...periods];
  }

  async function findMonthMarkers(
    workbook,
    period,
    sheetNames = [],
  ) {
    const matchers = periodMatchers(period);
    if (!matchers.length) {
      return [];
    }
    const requestedNames = Array.isArray(sheetNames)
      ? sheetNames
      : [sheetNames];
    const selectedNames = requestedNames.filter(Boolean);
    const sheets = selectedNames.length
      ? selectedNames
          .map((name) => workbook.sheetByName.get(name))
          .filter(Boolean)
      : workbook.sheets;
    const markers = [];
    for (const sheet of sheets) {
      const sheetRecord = await workbook.loadSheetRecord(sheet.name);
      for (const cell of api.elementsByLocalName(sheetRecord.document, "c")) {
        const parsed = api.parseCellReference(cell.getAttribute("r"));
        if (!parsed || parsed.row > 3) {
          continue;
        }
        const descriptor = api.cellDescriptor(cell, workbook.sharedStrings);
        if (
          typeof descriptor.value !== "string" ||
          !/工资|薪酬|薪资/.test(descriptor.value)
        ) {
          continue;
        }
        if (
          !matchers.some((matcher) => {
            matcher.lastIndex = 0;
            return matcher.test(descriptor.value);
          })
        ) {
          continue;
        }
        markers.push({
          sheet: sheet.name,
          row: parsed.row,
          column: parsed.column,
          cell: descriptor.address,
          before: descriptor.value,
        });
      }
    }
    return markers;
  }

  async function findPeriodCandidates(workbook, sheetNames = []) {
    const requestedNames = Array.isArray(sheetNames)
      ? sheetNames
      : [sheetNames];
    const selectedNames = requestedNames.filter(Boolean);
    const sheets = selectedNames.length
      ? selectedNames
          .map((name) => workbook.sheetByName.get(name))
          .filter(Boolean)
      : workbook.sheets;
    const candidates = [];
    for (const sheet of sheets) {
      const sheetRecord = await workbook.loadSheetRecord(sheet.name);
      for (const cell of api.elementsByLocalName(sheetRecord.document, "c")) {
        const parsed = api.parseCellReference(cell.getAttribute("r"));
        if (!parsed || parsed.row > 10) {
          continue;
        }
        const descriptor = api.cellDescriptor(cell, workbook.sharedStrings);
        if (typeof descriptor.value !== "string") {
          continue;
        }
        for (const period of periodsFromText(descriptor.value)) {
          candidates.push({
            period,
            sheet: sheet.name,
            row: parsed.row,
            column: parsed.column,
            cell: descriptor.address,
            text: descriptor.value,
            titleLike:
              parsed.row <= 3 &&
              /工资|薪酬|薪资|社保|公积金/.test(descriptor.value),
          });
        }
      }
    }
    return candidates;
  }

  async function updateMonthMarkers(
    workbook,
    markers,
    fromPeriod,
    toPeriod,
  ) {
    const results = [];
    for (const marker of markers) {
      const after = replacePeriodText(marker.before, fromPeriod, toPeriod);
      if (after === marker.before) {
        continue;
      }
      const result = await api.updateCell(
        workbook,
        marker.sheet,
        marker.row,
        marker.column,
        after,
        { preserveFormula: false },
      );
      results.push({ ...marker, after, result });
    }
    return results;
  }

  Object.assign(api, {
    periodParts,
    periodMatchers,
    replacePeriodText,
    periodsFromText,
    findMonthMarkers,
    findPeriodCandidates,
    updateMonthMarkers,
  });
})();
