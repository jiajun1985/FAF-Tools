(() => {
  const {
    parseWorkbook,
    findHeaderIndex,
    normalizeText,
    getTodayStamp,
    downloadBlob,
    escapeHtml
  } = window.FAFULIExcelUtils;
  const { setupUpload } = window.FAFULIFileUpload;

  const OUTPUT_HEADERS = [
    "被保险人姓名",
    "被保险人性别",
    "被保险人出生日期",
    "被保险人证件类型",
    "被保险人证件号",
    "主保险人姓名",
    "主险人性别",
    "主险人出生日期",
    "主险人证件类型",
    "主险人证件号",
    "与主险人关系",
    "医保属地",
    "起保时间",
    "减保时间",
    "月薪",
    "方案编号",
    "产品编号",
    "产品名称",
    "方案特别约定",
    "产品系统投保保额",
    "客服",
    "公司ERP客编",
    "客户名称",
    "保险公司",
    "是否高风险人群",
    "重疾名称",
    "客服维护时间",
    "集团户名称",
    "保全提取时间",
    "理赔既往症",
    "是否理赔高风险"
  ];
  const FIELD_MAPPING = {
    "被保险人姓名": "亲属姓名",
    "被保险人性别": "亲属性别",
    "被保险人出生日期": "亲属生日",
    "被保险人证件类型": "亲属证件号类别",
    "被保险人证件号": "亲属证件号",
    "主保险人姓名": "姓名",
    "主险人性别": "员工性别",
    "主险人出生日期": "出生日期",
    "主险人证件类型": "证件类型",
    "主险人证件号": "身份证",
    "与主险人关系": "亲属关系",
    "医保属地": "委派单中的城市",
    "方案编号": "报价单编号",
    "公司ERP客编": "客户编号",
    "客户名称": "客户名称"
  };
  const FIXED_OUTPUT_VALUES = {
    "月薪": 0,
    "客服": "FA 客服"
  };
  const MAIN_PRODUCT_HEADER = "产品名称";
  const PRODUCT_MAPPING_HEADERS = ["产品名称", "产品编码", "产品名称", "保额"];
  const PRODUCT_OUTPUT_HEADERS = new Set(["产品编号", "产品名称", "产品系统投保保额"]);
  const START_DATE_HEADERS = ["入职日期", "产品收费开始日期"];
  const DATE_OUTPUT_HEADERS = new Set(["被保险人出生日期", "主险人出生日期", "起保时间"]);

  const state = {
    main: null,
    productMapping: null,
    result: null
  };
  let nodes = {};

  function render() {
    return `<header class="topbar">
          <div>
            <h1>ERP保全名单数据清洗</h1>
            <p class="tagline">上传 ERP 源数据，按保全名单模板清洗表头、顺序和起保时间。</p>
          </div>
          <div class="local-badge">本地处理 无需网络连接</div>
        </header>

        <section class="section">
          <div class="section-head">
            <h2>1. 上传文件</h2>
            <span class="hint">仅支持 .xlsx，默认读取第一个工作表</span>
          </div>
          <div class="upload-grid">
            <label class="dropzone" id="erpCleanDrop">
              <input id="erpCleanFile" type="file" accept=".xlsx">
              <div>
                <p class="drop-title">主表：ERP 源数据</p>
                <p class="drop-copy">拖拽或点击上传，将按固定保全名单字段顺序清洗导出。</p>
              </div>
              <div class="file-state">
                <div class="file-name" id="erpCleanFileName">未选择文件</div>
                <button class="clear-file" id="erpCleanClear" type="button">删除</button>
              </div>
            </label>
            <label class="dropzone" id="erpProductMapDrop">
              <input id="erpProductMapFile" type="file" accept=".xlsx">
              <div>
                <p class="drop-title">映射表：ERP 产品转换关系</p>
                <p class="drop-copy">第 1 列产品名称用于匹配，第 2 至 4 列导出产品编号、产品名称和投保保额。</p>
              </div>
              <div class="file-state">
                <div class="file-name" id="erpProductMapFileName">未选择文件</div>
                <button class="clear-file" id="erpProductMapClear" type="button">删除</button>
              </div>
            </label>
          </div>
        </section>

        <section class="section">
          <div class="section-head">
            <h2>2. 清洗预览</h2>
            <span class="hint">缺失源字段会填空值，不影响导出</span>
          </div>
          <div id="erpCleanStatusBox" class="status">请先上传 ERP 源数据主表。</div>
          <ul id="erpCleanErrorList" class="error-list hidden"></ul>
          <div id="erpCleanPreviewEmpty" class="empty" style="margin-top: 14px;">暂无预览结果。</div>
          <div id="erpCleanPreviewWrap" class="table-wrap hidden" style="margin-top: 14px;">
            <table>
              <thead>
                <tr>
                  <th>被保险人姓名</th>
                  <th>与主险人关系</th>
                  <th>医保属地</th>
                  <th>起保时间</th>
                  <th>产品名称</th>
                  <th>公司ERP客编</th>
                  <th>客户名称</th>
                </tr>
              </thead>
              <tbody id="erpCleanPreviewBody"></tbody>
            </table>
          </div>
        </section>

        <section class="section">
          <div class="section-head">
            <h2>3. 导出文件</h2>
            <span class="hint">按固定表头顺序导出单个 Excel</span>
          </div>
          <div class="actions">
            <span class="hint">「客服」统一写 FA 客服，「月薪」统一写 0，「减保时间」留空。</span>
            <button id="erpCleanExportButton" type="button" disabled>导出 Excel</button>
          </div>
        </section>`;
  }

  function init(root) {
    nodes = {
      erpCleanFile: root.querySelector("#erpCleanFile"),
      erpCleanFileName: root.querySelector("#erpCleanFileName"),
      erpCleanClear: root.querySelector("#erpCleanClear"),
      erpProductMapFile: root.querySelector("#erpProductMapFile"),
      erpProductMapFileName: root.querySelector("#erpProductMapFileName"),
      erpProductMapClear: root.querySelector("#erpProductMapClear"),
      erpCleanStatusBox: root.querySelector("#erpCleanStatusBox"),
      erpCleanErrorList: root.querySelector("#erpCleanErrorList"),
      erpCleanPreviewEmpty: root.querySelector("#erpCleanPreviewEmpty"),
      erpCleanPreviewWrap: root.querySelector("#erpCleanPreviewWrap"),
      erpCleanPreviewBody: root.querySelector("#erpCleanPreviewBody"),
      erpCleanExportButton: root.querySelector("#erpCleanExportButton")
    };

    setupUpload({
      input: nodes.erpCleanFile,
      dropzone: root.querySelector("#erpCleanDrop"),
      nameNode: nodes.erpCleanFileName,
      clearButton: nodes.erpCleanClear,
      parseWorkbook,
      onParsed: (parsed) => { state.main = parsed; },
      onCleared: () => { state.main = null; },
      onChange: processErpCleanFile
    });
    setupUpload({
      input: nodes.erpProductMapFile,
      dropzone: root.querySelector("#erpProductMapDrop"),
      nameNode: nodes.erpProductMapFileName,
      clearButton: nodes.erpProductMapClear,
      parseWorkbook,
      onParsed: (parsed) => { state.productMapping = parsed; },
      onCleared: () => { state.productMapping = null; },
      onChange: processErpCleanFile
    });
    nodes.erpCleanExportButton.addEventListener("click", exportErpCleanExcel);
  }

  function processErpCleanFile(errorMessage) {
    resetErpCleanViews();
    if (errorMessage) {
      setErpCleanStatus("error", errorMessage);
      return;
    }
    if (!state.main) {
      setErpCleanStatus("info", "请先上传 ERP 源数据主表。");
      return;
    }
    if (!state.productMapping) {
      setErpCleanStatus("warn", "请继续上传 ERP 产品转换关系映射表。");
      return;
    }
    const productMappingErrors = validateProductMapping(state.productMapping);
    if (productMappingErrors.length) {
      renderErpCleanMessages(productMappingErrors);
      setErpCleanStatus("error", "产品转换关系映射表校验未通过，请修正后重新上传。");
      return;
    }

    const result = buildErpCleanResult(state.main, state.productMapping);
    state.result = result;
    renderErpCleanMessages(buildWarningMessages(result));
    renderErpCleanPreview(result);
    setErpCleanStatus(buildStatusType(result), buildStatusText(result));
    nodes.erpCleanExportButton.disabled = result.rows.length === 0;
  }

  function buildErpCleanResult(workbookData, productMappingData) {
    const headerIndexes = buildHeaderIndexes(workbookData.headers);
    const missingHeaders = findMissingHeaders(headerIndexes);
    const outputIndexes = buildOutputIndexes();
    const productMapping = buildProductMapping(productMappingData);
    const productSourceIndex = headerIndexes[MAIN_PRODUCT_HEADER];
    let mappedProductCount = 0;
    const unmatchedProducts = new Set();
    const rows = workbookData.rows.map((row) => {
      const sourceProductName = productSourceIndex === -1 ? "" : normalizeText(row[productSourceIndex]);
      const productInfo = sourceProductName ? productMapping.get(sourceProductName) : null;
      if (productInfo) {
        mappedProductCount += 1;
      } else if (sourceProductName) {
        unmatchedProducts.add(sourceProductName);
      }
      return buildOutputRow(row, headerIndexes, outputIndexes, productInfo);
    });
    return {
      headers: OUTPUT_HEADERS,
      rows,
      missingHeaders,
      mappedProductCount,
      unmatchedProductCount: rows.length - mappedProductCount,
      unmatchedProducts: Array.from(unmatchedProducts),
      sourceRowCount: workbookData.rows.length
    };
  }

  function buildHeaderIndexes(headers) {
    const indexMap = {};
    Object.values(FIELD_MAPPING).forEach((sourceHeader) => {
      indexMap[sourceHeader] = findHeaderIndex(headers, sourceHeader);
    });
    START_DATE_HEADERS.forEach((sourceHeader) => {
      indexMap[sourceHeader] = findHeaderIndex(headers, sourceHeader);
    });
    indexMap[MAIN_PRODUCT_HEADER] = findHeaderIndex(headers, MAIN_PRODUCT_HEADER);
    return indexMap;
  }

  function findMissingHeaders(headerIndexes) {
    const requiredSourceHeaders = [...new Set([
      ...Object.values(FIELD_MAPPING),
      MAIN_PRODUCT_HEADER,
      ...START_DATE_HEADERS
    ])];
    return requiredSourceHeaders.filter((header) => headerIndexes[header] === -1);
  }

  function buildOutputIndexes() {
    return OUTPUT_HEADERS.reduce((indexes, header, index) => {
      indexes[header] = index;
      return indexes;
    }, {});
  }

  function buildOutputRow(sourceRow, headerIndexes, outputIndexes, productInfo) {
    const outputRow = OUTPUT_HEADERS.map((header) => {
      if (header === "起保时间") return buildStartDate(sourceRow, headerIndexes);
      if (Object.prototype.hasOwnProperty.call(FIXED_OUTPUT_VALUES, header)) return FIXED_OUTPUT_VALUES[header];
      if (PRODUCT_OUTPUT_HEADERS.has(header)) return productInfo ? productInfo[header] : "";
      const sourceHeader = FIELD_MAPPING[header];
      if (!sourceHeader || headerIndexes[sourceHeader] === -1) return "";
      const value = sourceRow[headerIndexes[sourceHeader]];
      if (DATE_OUTPUT_HEADERS.has(header)) return formatDateValue(parseDateValue(value));
      return normalizeText(value);
    });
    outputRow[outputIndexes["减保时间"]] = "";
    return outputRow;
  }

  function validateProductMapping(productMappingData) {
    const headers = productMappingData.headers || [];
    if (headers.length < PRODUCT_MAPPING_HEADERS.length) {
      return ["产品转换关系映射表至少需要前 4 列：产品名称、产品编码、产品名称、保额。"];
    }
    const errors = [];
    PRODUCT_MAPPING_HEADERS.forEach((expectedHeader, index) => {
      if (normalizeText(headers[index]) !== expectedHeader) {
        errors.push(`产品转换关系映射表第 ${index + 1} 列应为「${expectedHeader}」。`);
      }
    });
    return errors;
  }

  function buildProductMapping(productMappingData) {
    const productMapping = new Map();
    productMappingData.rows.forEach((row) => {
      const sourceProductName = normalizeText(row[0]);
      if (!sourceProductName || productMapping.has(sourceProductName)) return;
      productMapping.set(sourceProductName, {
        "产品编号": normalizeText(row[1]),
        "产品名称": normalizeText(row[2]),
        "产品系统投保保额": normalizeText(row[3])
      });
    });
    return productMapping;
  }

  function buildStartDate(sourceRow, headerIndexes) {
    const dates = START_DATE_HEADERS
      .map((header) => parseDateValue(sourceRow[headerIndexes[header]]))
      .filter(Boolean);
    if (!dates.length) return "";
    const latest = dates.reduce((max, date) => (date > max ? date : max), dates[0]);
    return formatDateValue(latest);
  }

  function parseDateValue(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return startOfDay(value);
    const text = normalizeText(value);
    if (!text) return null;
    const isoMatch = text.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?/);
    if (isoMatch) return buildDate(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));
    const compactMatch = text.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (compactMatch) return buildDate(Number(compactMatch[1]), Number(compactMatch[2]), Number(compactMatch[3]));
    const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (slashMatch) {
      return buildDate(normalizeYear(Number(slashMatch[3])), Number(slashMatch[1]), Number(slashMatch[2]));
    }
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : startOfDay(parsed);
  }

  function buildDate(year, month, day) {
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
    return date;
  }

  function normalizeYear(year) {
    return year < 100 ? 2000 + year : year;
  }

  function startOfDay(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function formatDateValue(date) {
    if (!date) return "";
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function renderErpCleanPreview(result) {
    if (!result.rows.length) {
      nodes.erpCleanPreviewEmpty.textContent = "主表没有可导出的数据行。";
      nodes.erpCleanPreviewEmpty.classList.remove("hidden");
      nodes.erpCleanPreviewWrap.classList.add("hidden");
      return;
    }
    const previewIndexes = ["被保险人姓名", "与主险人关系", "医保属地", "起保时间", "产品名称", "公司ERP客编", "客户名称"]
      .map((header) => result.headers.indexOf(header));
    nodes.erpCleanPreviewBody.innerHTML = result.rows.slice(0, 10).map((row) => `
      <tr>
        ${previewIndexes.map((index) => `<td>${escapeHtml(row[index])}</td>`).join("")}
      </tr>
    `).join("");
    nodes.erpCleanPreviewEmpty.classList.add("hidden");
    nodes.erpCleanPreviewWrap.classList.remove("hidden");
  }

  function renderErpCleanMessages(messages) {
    if (!messages.length) {
      nodes.erpCleanErrorList.classList.add("hidden");
      nodes.erpCleanErrorList.innerHTML = "";
      return;
    }
    nodes.erpCleanErrorList.innerHTML = messages.map((message) => `<li>${escapeHtml(message)}</li>`).join("");
    nodes.erpCleanErrorList.classList.remove("hidden");
  }

  function buildWarningMessages(result) {
    const messages = result.missingHeaders.map((header) => `源表缺少「${header}」，相关导出字段会填空值。`);
    if (result.unmatchedProducts.length) {
      const productNames = result.unmatchedProducts.slice(0, 8).join("、");
      const suffix = result.unmatchedProducts.length > 8 ? "等" : "";
      messages.push(`有 ${result.unmatchedProducts.length} 个产品名称未在映射表中找到：${productNames}${suffix}。`);
    }
    return messages;
  }

  function buildStatusText(result) {
    const warning = result.missingHeaders.length ? `；缺失 ${result.missingHeaders.length} 个源字段，已按空值处理` : "";
    const productText = `；产品映射命中 ${result.mappedProductCount} 行，未匹配 ${result.unmatchedProductCount} 行`;
    return `已清洗 ${result.rows.length} 行，导出字段 ${result.headers.length} 个${productText}${warning}。`;
  }

  function buildStatusType(result) {
    return result.missingHeaders.length || result.unmatchedProductCount ? "warn" : "ok";
  }

  async function exportErpCleanExcel() {
    if (!state.result || !state.result.rows.length) return;
    nodes.erpCleanExportButton.disabled = true;
    nodes.erpCleanExportButton.textContent = "正在导出...";
    try {
      const excelData = buildErpCleanWorkbook(state.result);
      const blob = new Blob([excelData], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      downloadBlob(blob, `ERP保全名单数据清洗_${getTodayStamp()}.xlsx`);
      nodes.erpCleanExportButton.textContent = "导出 Excel";
    } catch (error) {
      setErpCleanStatus("error", `导出失败：${error.message}`);
      nodes.erpCleanExportButton.textContent = "导出 Excel";
    } finally {
      nodes.erpCleanExportButton.disabled = false;
    }
  }

  function buildErpCleanWorkbook(result) {
    const worksheet = XLSX.utils.aoa_to_sheet([result.headers, ...result.rows]);
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    worksheet["!cols"] = result.headers.map((header, index) => ({
      wch: Math.min(Math.max(String(header).length + 4, index === 0 ? 16 : 10), 28)
    }));
    worksheet["!autofilter"] = { ref: XLSX.utils.encode_range(range) };

    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!worksheet[cellRef]) continue;
      worksheet[cellRef].s = {
        font: { bold: true, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "176B87" } },
        alignment: { horizontal: "center", vertical: "center" }
      };
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "清洗结果");
    return XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      cellStyles: true
    });
  }

  function resetErpCleanViews() {
    nodes.erpCleanErrorList.classList.add("hidden");
    nodes.erpCleanErrorList.innerHTML = "";
    nodes.erpCleanPreviewEmpty.textContent = "暂无预览结果。";
    nodes.erpCleanPreviewEmpty.classList.remove("hidden");
    nodes.erpCleanPreviewWrap.classList.add("hidden");
    nodes.erpCleanPreviewBody.innerHTML = "";
    nodes.erpCleanExportButton.disabled = true;
    state.result = null;
  }

  function setErpCleanStatus(type, message) {
    nodes.erpCleanStatusBox.className = "status";
    if (type !== "info") {
      nodes.erpCleanStatusBox.classList.add(type);
    }
    nodes.erpCleanStatusBox.textContent = message;
  }

  window.FAFULITools = window.FAFULITools || {};
  window.FAFULITools.erpInsuranceClean = {
    id: "erp-insurance-clean",
    group: "保险运营",
    name: "ERP保全名单数据清洗",
    render,
    init
  };
})();
