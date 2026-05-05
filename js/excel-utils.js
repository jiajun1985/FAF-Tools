(() => {
  async function parseWorkbook(file) {
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      throw new Error("仅支持 .xlsx 文件。");
    }
    if (file.name.startsWith("~$")) {
      throw new Error("检测到 Excel 临时锁文件，请上传正式工作簿。");
    }
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array", cellDates: true });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      throw new Error("工作簿中没有可读取的工作表。");
    }
    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false
    });
    if (!rows.length) {
      throw new Error("第一个工作表为空。");
    }
    const headers = rows[0].map((value) => normalizeText(value));
    return {
      fileName: file.name,
      sheetName: firstSheetName,
      headers,
      rows: rows.slice(1).filter((row) => !isBlankRow(row))
    };
  }

  function findHeaderIndex(headers, target) {
      return headers.findIndex((header) => normalizeText(header) === target);
    }

function findFirstHeaderIndex(headers, candidates) {
      for (const candidate of candidates) {
        const index = findHeaderIndex(headers, candidate);
        if (index !== -1) return index;
      }
      return -1;
    }

function normalizeText(value) {
      if (value === null || value === undefined) return "";
      return String(value).trim();
    }

function isBlankRow(row) {
      return row.every((value) => normalizeText(value) === "");
    }

  function sanitizeFileName(name, fallback = "未命名") {
    const safeFallback = normalizeText(fallback) || "未命名";
    const cleaned = normalizeText(name || safeFallback)
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, " ")
      .trim();
    return cleaned || safeFallback;
  }

  function getTodayStamp() {
      const date = new Date();
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, "0");
      const day = String(date.getDate()).padStart(2, "0");
      return `${year}${month}${day}`;
    }

function uniqueFileName(fileName, usedNames) {
      if (!usedNames.has(fileName)) {
        usedNames.add(fileName);
        return fileName;
      }
      const dotIndex = fileName.lastIndexOf(".");
      const base = dotIndex === -1 ? fileName : fileName.slice(0, dotIndex);
      const ext = dotIndex === -1 ? "" : fileName.slice(dotIndex);
      let counter = 2;
      let next = `${base}_${counter}${ext}`;
      while (usedNames.has(next)) {
        counter += 1;
        next = `${base}_${counter}${ext}`;
      }
      usedNames.add(next);
      return next;
    }

function downloadBlob(blob, fileName) {
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

function escapeHtml(value) {
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }

function parseAmount(value) {
      if (typeof value === "number") return Number.isFinite(value) ? value : 0;
      const text = normalizeText(value).replace(/,/g, "");
      if (!text) return 0;
      const amount = Number(text);
      return Number.isFinite(amount) ? amount : 0;
    }

function formatAmount(value) {
      return Number(value || 0).toLocaleString("zh-CN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    }

  window.FAFULIExcelUtils = {
    parseWorkbook,
    findHeaderIndex,
    findFirstHeaderIndex,
    normalizeText,
    isBlankRow,
    sanitizeFileName,
    getTodayStamp,
    uniqueFileName,
    downloadBlob,
    escapeHtml,
    parseAmount,
    formatAmount
  };
})();
