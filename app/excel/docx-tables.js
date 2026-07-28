(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const WORD_NS =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  function directChildren(node, localName) {
    return [...node.childNodes].filter(
      (child) =>
        child.nodeType === Node.ELEMENT_NODE &&
        child.localName === localName,
    );
  }

  function cellText(cell) {
    const parts = [];
    for (const node of cell.getElementsByTagNameNS(WORD_NS, "*")) {
      if (node.localName === "t" && node.textContent) {
        parts.push(node.textContent);
      } else if (node.localName === "tab") {
        parts.push("\t");
      } else if (node.localName === "br" || node.localName === "cr") {
        parts.push("\n");
      }
    }
    return parts.join("").trim();
  }

  function cellSpan(cell) {
    const span = cell.getElementsByTagNameNS(WORD_NS, "gridSpan")[0];
    const value =
      span?.getAttributeNS(WORD_NS, "val") ||
      span?.getAttribute("w:val") ||
      span?.getAttribute("val");
    const numeric = Number(value);
    return Number.isInteger(numeric) && numeric > 0 ? numeric : 1;
  }

  function tableMatrix(table) {
    return directChildren(table, "tr").map((row) => {
      const output = [];
      for (const cell of directChildren(row, "tc")) {
        output.push(cellText(cell));
        for (let offset = 1; offset < cellSpan(cell); offset += 1) {
          output.push("");
        }
      }
      return output;
    });
  }

  async function parseDocxTables(file) {
    if (!window.JSZip) {
      throw new Error("DOCX 读取组件未加载");
    }
    let archive;
    try {
      archive = await window.JSZip.loadAsync(await file.arrayBuffer());
    } catch (error) {
      throw new Error(`${file.name} 不是可读取的 DOCX：${error.message}`);
    }
    const documentPart = archive.file("word/document.xml");
    if (!documentPart) {
      throw new Error(`${file.name} 缺少 Word 正文`);
    }
    const xml = await documentPart.async("string");
    const documentXml = new DOMParser().parseFromString(
      xml,
      "application/xml",
    );
    if (documentXml.querySelector("parsererror")) {
      throw new Error(`${file.name} 的 Word 正文已损坏`);
    }
    const tables = [
      ...documentXml.getElementsByTagNameNS(WORD_NS, "tbl"),
    ]
      .map((table, index) => ({
        name: `表格${index + 1}`,
        matrix: tableMatrix(table),
      }))
      .filter((table) => table.matrix.some((row) => row.some(Boolean)));
    if (!tables.length) {
      throw new Error(`${file.name} 没有可读取的表格`);
    }
    return { name: file.name, tables };
  }

  Object.assign(api, {
    parseDocxTables,
  });
})();
