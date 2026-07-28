(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const MAIN_NS =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const REL_NS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PACKAGE_REL_NS =
    "http://schemas.openxmlformats.org/package/2006/relationships";
  const CONTENT_TYPES_NS =
    "http://schemas.openxmlformats.org/package/2006/content-types";

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFKC")
      .trim()
      .toLowerCase()
      .replace(/[\s_\-—–·:：()（）[\]【】/\\]+/g, "");
  }

  function parseXml(text, label) {
    const documentNode = new DOMParser().parseFromString(text, "application/xml");
    const parserError = documentNode.getElementsByTagName("parsererror")[0];
    if (parserError) {
      throw new Error(`${label} XML 无法解析`);
    }
    return documentNode;
  }

  function serializeXml(documentNode) {
    const serialized = new XMLSerializer().serializeToString(documentNode);
    return serialized.startsWith("<?xml")
      ? serialized
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${serialized}`;
  }

  function elementsByLocalName(root, localName) {
    return Array.from(root.getElementsByTagNameNS("*", localName));
  }

  function firstByLocalName(root, localName) {
    return root.getElementsByTagNameNS("*", localName)[0] || null;
  }

  function directChildrenByLocalName(root, localName) {
    return Array.from(root.childNodes).filter(
      (node) => node.nodeType === Node.ELEMENT_NODE && node.localName === localName,
    );
  }

  function childByLocalName(root, localName) {
    return directChildrenByLocalName(root, localName)[0] || null;
  }

  function resolveZipPath(basePath, target) {
    if (!target) {
      return "";
    }
    if (target.startsWith("/")) {
      return target.replace(/^\/+/, "");
    }
    const baseSegments = basePath.split("/");
    baseSegments.pop();
    for (const segment of target.split("/")) {
      if (!segment || segment === ".") {
        continue;
      }
      if (segment === "..") {
        baseSegments.pop();
      } else {
        baseSegments.push(segment);
      }
    }
    return baseSegments.join("/");
  }

  function relationshipMap(documentNode, basePath) {
    const output = new Map();
    for (const relationship of elementsByLocalName(documentNode, "Relationship")) {
      output.set(relationship.getAttribute("Id"), {
        id: relationship.getAttribute("Id"),
        type: relationship.getAttribute("Type") || "",
        target: relationship.getAttribute("Target") || "",
        targetMode: relationship.getAttribute("TargetMode") || "",
        resolvedTarget: resolveZipPath(
          basePath,
          relationship.getAttribute("Target") || "",
        ),
      });
    }
    return output;
  }

  function columnLettersToNumber(letters) {
    let result = 0;
    for (const character of String(letters).toUpperCase()) {
      result = result * 26 + character.charCodeAt(0) - 64;
    }
    return result;
  }

  function columnNumberToLetters(number) {
    let cursor = Number(number);
    let result = "";
    while (cursor > 0) {
      cursor -= 1;
      result = String.fromCharCode(65 + (cursor % 26)) + result;
      cursor = Math.floor(cursor / 26);
    }
    return result;
  }

  function parseCellReference(reference) {
    const match = String(reference || "").match(/^([A-Z]+)(\d+)$/i);
    if (!match) {
      return null;
    }
    return {
      column: columnLettersToNumber(match[1]),
      row: Number(match[2]),
    };
  }

  function translateFormulaA1(formula, rowDelta = 0, columnDelta = 0) {
    return String(formula || "").replace(
      /(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
      (match, absoluteColumn, columnLetters, absoluteRow, rowText) => {
        let column = columnLettersToNumber(columnLetters);
        let row = Number(rowText);
        if (!absoluteColumn) {
          column += Number(columnDelta) || 0;
        }
        if (!absoluteRow) {
          row += Number(rowDelta) || 0;
        }
        if (column <= 0 || row <= 0) {
          return match;
        }
        return `${absoluteColumn}${columnNumberToLetters(column)}${absoluteRow}${row}`;
      },
    );
  }

  function readTextRuns(root) {
    return elementsByLocalName(root, "t")
      .map((node) => node.textContent || "")
      .join("");
  }

  function readSharedStrings(documentNode) {
    if (!documentNode) {
      return [];
    }
    return elementsByLocalName(documentNode, "si").map((item) =>
      readTextRuns(item),
    );
  }

  function cellDescriptor(cell, sharedStrings) {
    if (!cell) {
      return {
        address: "",
        value: null,
        formula: "",
        hasFormula: false,
        hasCachedValue: false,
        type: "",
        styleId: "",
      };
    }
    const type = cell.getAttribute("t") || "";
    const formulaNode = childByLocalName(cell, "f");
    const valueNode = childByLocalName(cell, "v");
    const inlineString = childByLocalName(cell, "is");
    const formula = formulaNode?.textContent || "";
    const rawValue = valueNode?.textContent ?? "";
    let value = null;

    if (type === "inlineStr") {
      value = inlineString ? readTextRuns(inlineString) : "";
    } else if (type === "s") {
      const index = Number(rawValue);
      value = Number.isInteger(index) ? sharedStrings[index] ?? "" : "";
    } else if (type === "b") {
      value = rawValue === "1";
    } else if (type === "str" || type === "e" || type === "d") {
      value = rawValue;
    } else if (rawValue !== "") {
      const numeric = Number(rawValue);
      value = Number.isFinite(numeric) ? numeric : rawValue;
    } else if (inlineString) {
      value = readTextRuns(inlineString);
    }

    return {
      address: cell.getAttribute("r") || "",
      value,
      formula,
      hasFormula: Boolean(formulaNode),
      hasCachedValue: Boolean(valueNode || inlineString),
      type,
      styleId: cell.getAttribute("s") || "",
    };
  }

  function rowCellMap(documentNode) {
    const output = new Map();
    for (const cell of elementsByLocalName(documentNode, "c")) {
      const parsed = parseCellReference(cell.getAttribute("r"));
      if (!parsed) {
        continue;
      }
      if (!output.has(parsed.row)) {
        output.set(parsed.row, new Map());
      }
      output.get(parsed.row).set(parsed.column, cell);
    }
    return output;
  }

  Object.assign(api, {
    MAIN_NS,
    REL_NS,
    PACKAGE_REL_NS,
    CONTENT_TYPES_NS,
    normalizeText,
    parseXml,
    serializeXml,
    elementsByLocalName,
    firstByLocalName,
    directChildrenByLocalName,
    childByLocalName,
    resolveZipPath,
    relationshipMap,
    columnLettersToNumber,
    columnNumberToLetters,
    parseCellReference,
    translateFormulaA1,
    readTextRuns,
    readSharedStrings,
    cellDescriptor,
    rowCellMap,
  });
})();
