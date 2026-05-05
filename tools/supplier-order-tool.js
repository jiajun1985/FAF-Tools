(() => {
  const {
    parseWorkbook,
    findHeaderIndex,
    normalizeText,
    sanitizeFileName,
    getTodayStamp,
    uniqueFileName,
    downloadBlob,
    escapeHtml,
    parseAmount,
    formatAmount
  } = window.FAFULIExcelUtils;
  const { setupUpload } = window.FAFULIFileUpload;

  const SUPPLIER_HEADER = "供应商";
  const SETTLEMENT_TOTAL_HEADER = "结算总价";
  const UNFILLED_SUPPLIER = "未填写供应商";
  const ORDER_EXPORT_EXCLUDED_FIELDS = [
    "所属企业",
    "企业编号",
    "集团户名称",
    "会员名称",
    "会员手机号",
    "会员证件号",
    "账单编号",
    "erp编号",
    "客服",
    "部门",
    "积分售价",
    "优惠类型",
    "优惠总金额",
    "订单金额",
    "订单金额（收方含税价已还原售价）",
    "收方税率",
    "收方不含税价",
    "超级积分金额",
    "商城积分金额",
    "现金支付金额",
    "第三方支付方式",
    "支付类型",
    "支付订单号",
    "第三方支付流水",
    "erp商品名称",
    "方案名称",
    "证件号",
    "姓名"
  ];

  const state = {
    orderMain: null,
    orderResult: null
  };
  let nodes = {};

  function render() {
    return `<header class="topbar">
            <div>
              <h1>商城订单按供应商拆分</h1>
              <p class="tagline">上传商城订单主表，按供应商拆分成多个 Excel 文件并打包导出。</p>
            </div>
            <div class="local-badge">本地处理 无需网络连接</div>
          </header>

          <section class="section">
            <div class="section-head">
              <h2>1. 上传文件</h2>
              <span class="hint">仅支持 .xlsx，默认读取第一个工作表</span>
            </div>
            <div class="upload-grid single-upload-grid">
              <label class="dropzone" id="orderDrop">
                <input id="orderFile" type="file" accept=".xlsx">
                <div>
                  <p class="drop-title">主表：商城订单</p>
                  <p class="drop-copy">拖拽或点击上传，需包含「供应商」。</p>
                </div>
                <div class="file-state">
                  <div class="file-name" id="orderFileName">未选择文件</div>
                  <button class="clear-file" id="orderClear" type="button">删除</button>
                </div>
              </label>
            </div>
          </section>

          <section class="section">
            <div class="section-head">
              <h2>2. 供应商预览</h2>
              <span class="hint">按供应商汇总订单行数和导出文件名</span>
            </div>
            <div id="orderStatusBox" class="status">请先上传商城订单主表。</div>
            <ul id="orderErrorList" class="error-list hidden"></ul>
            <div id="orderPreviewEmpty" class="empty" style="margin-top: 14px;">暂无预览结果。</div>
            <div id="orderPreviewWrap" class="table-wrap hidden" style="margin-top: 14px;">
              <table>
                <thead>
                  <tr>
                    <th>供应商</th>
                    <th class="number">订单行数</th>
                    <th class="number">结算总价</th>
                  </tr>
                </thead>
                <tbody id="orderPreviewBody"></tbody>
              </table>
            </div>
          </section>

          <section class="section">
            <div class="section-head">
              <h2>3. 导出文件</h2>
              <span class="hint">每个供应商一个 Excel，全部打包成 ZIP</span>
            </div>
            <div class="actions">
              <span class="hint">导出文件名会自动替换系统不允许的字符。</span>
              <button id="orderExportButton" type="button" disabled>导出 ZIP</button>
            </div>
          </section>
        `;
  }

  function init(root) {
    nodes = {
      orderFile: root.querySelector("#orderFile"),
      orderFileName: root.querySelector("#orderFileName"),
      orderClear: root.querySelector("#orderClear"),
      orderStatusBox: root.querySelector("#orderStatusBox"),
      orderErrorList: root.querySelector("#orderErrorList"),
      orderPreviewEmpty: root.querySelector("#orderPreviewEmpty"),
      orderPreviewWrap: root.querySelector("#orderPreviewWrap"),
      orderPreviewBody: root.querySelector("#orderPreviewBody"),
      orderExportButton: root.querySelector("#orderExportButton")
    };

    setupUpload({
      input: nodes.orderFile,
      dropzone: root.querySelector("#orderDrop"),
      nameNode: nodes.orderFileName,
      clearButton: nodes.orderClear,
      parseWorkbook,
      onParsed: (parsed) => { state.orderMain = parsed; },
      onCleared: () => { state.orderMain = null; },
      onChange: processOrderFile
    });
    nodes.orderExportButton.addEventListener("click", exportOrderZip);
  }

  function processOrderFile(errorMessage) {
      resetOrderViews();
      if (errorMessage) {
        setOrderStatus("error", errorMessage);
        return;
      }
      if (!state.orderMain) {
        setOrderStatus("info", "请先上传商城订单主表。");
        return;
      }

      const supplierIndex = findHeaderIndex(state.orderMain.headers, SUPPLIER_HEADER);
      const settlementTotalIndex = findHeaderIndex(state.orderMain.headers, SETTLEMENT_TOTAL_HEADER);
      if (supplierIndex === -1) {
        renderOrderErrors(["商城订单主表缺少表头：「供应商」。"]);
        setOrderStatus("error", "表头校验未通过，请修正后重新上传。");
        return;
      }
      if (settlementTotalIndex === -1) {
        renderOrderErrors(["商城订单主表缺少表头：「结算总价」。"]);
        setOrderStatus("error", "表头校验未通过，请修正后重新上传。");
        return;
      }

      const result = buildOrderSplitResult(state.orderMain, supplierIndex, settlementTotalIndex);
      state.orderResult = result;
      renderOrderPreview(result.groups);
      setOrderStatus("ok", `已识别 ${result.groups.length} 个供应商，共 ${result.totalRows} 条订单，可导出。`);
      nodes.orderExportButton.disabled = result.groups.length === 0;
    }

function buildOrderSplitResult(workbookData, supplierIndex, settlementTotalIndex) {
      const groups = new Map();
      workbookData.rows.forEach((row) => {
        const supplier = normalizeText(row[supplierIndex]) || UNFILLED_SUPPLIER;
        const cells = workbookData.headers.map((_, index) => normalizeText(row[index]));
        const settlementTotal = parseAmount(row[settlementTotalIndex]);
        if (!groups.has(supplier)) {
          groups.set(supplier, {
            supplier,
            rows: [],
            total: 0,
            settlementTotal: 0
          });
        }
        const group = groups.get(supplier);
        group.rows.push(cells);
        group.total += 1;
        group.settlementTotal += settlementTotal;
      });

      const exportColumnOrder = buildOrderExportColumnOrder(workbookData.headers);
      return {
        headers: workbookData.headers,
        supplierIndex,
        settlementTotalIndex,
        exportColumnOrder,
        groups: Array.from(groups.values()).sort((a, b) => b.settlementTotal - a.settlementTotal),
        totalRows: workbookData.rows.length
      };
    }

function renderOrderPreview(groups) {
      nodes.orderPreviewBody.innerHTML = groups.map((group) => `
        <tr>
          <td>${escapeHtml(group.supplier)}</td>
          <td class="number">${group.total.toLocaleString("zh-CN")}</td>
          <td class="number">${formatAmount(group.settlementTotal)}</td>
        </tr>
      `).join("");
      nodes.orderPreviewEmpty.classList.add("hidden");
      nodes.orderPreviewWrap.classList.remove("hidden");
    }

async function exportOrderZip() {
      if (!state.orderResult || !state.orderResult.groups.length) return;
      nodes.orderExportButton.disabled = true;
      nodes.orderExportButton.textContent = "正在导出...";
      try {
        const zip = new JSZip();
        const usedNames = new Set();
        state.orderResult.groups.forEach((group) => {
          const excelData = buildOrderWorkbookForSupplier(group, state.orderResult.headers, state.orderResult.exportColumnOrder);
          const fileName = uniqueFileName(buildSupplierFileName(group.supplier), usedNames);
          zip.file(fileName, excelData);
        });
        const blob = await zip.generateAsync({ type: "blob" });
        downloadBlob(blob, `商城订单_按供应商拆分_${getTodayStamp()}.zip`);
        nodes.orderExportButton.textContent = "导出 ZIP";
      } catch (error) {
        setOrderStatus("error", `导出失败：${error.message}`);
        nodes.orderExportButton.textContent = "导出 ZIP";
      } finally {
        nodes.orderExportButton.disabled = false;
      }
    }

function buildOrderWorkbookForSupplier(group, headers, exportColumnOrder) {
      const outputHeaders = exportColumnOrder.map((index) => headers[index]);
      const outputRows = group.rows.map((row) => exportColumnOrder.map((index) => row[index]));
      const aoa = [outputHeaders, ...outputRows];
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      worksheet["!cols"] = outputHeaders.map((header, index) => ({
        wch: Math.min(Math.max(String(header).length + 4, index === 0 ? 20 : 10), 28)
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
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      return XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
        cellStyles: true
      });
    }

function buildOrderExportColumnOrder(headers) {
      const excluded = new Set(ORDER_EXPORT_EXCLUDED_FIELDS);
      return headers
        .map((header, index) => ({ header, index }))
        .filter((item) => !excluded.has(item.header))
        .map((item) => item.index);
    }

function resetOrderViews() {
      nodes.orderErrorList.classList.add("hidden");
      nodes.orderErrorList.innerHTML = "";
      nodes.orderPreviewEmpty.classList.remove("hidden");
      nodes.orderPreviewWrap.classList.add("hidden");
      nodes.orderPreviewBody.innerHTML = "";
      nodes.orderExportButton.disabled = true;
      state.orderResult = null;
    }

function setOrderStatus(type, message) {
      nodes.orderStatusBox.className = "status";
      if (type !== "info") {
        nodes.orderStatusBox.classList.add(type);
      }
      nodes.orderStatusBox.textContent = message;
    }

function renderOrderErrors(errors) {
      if (!errors.length) {
        nodes.orderErrorList.classList.add("hidden");
        nodes.orderErrorList.innerHTML = "";
        return;
      }
      nodes.orderErrorList.innerHTML = errors.map((error) => `<li>${escapeHtml(error)}</li>`).join("");
      nodes.orderErrorList.classList.remove("hidden");
    }

  function buildSupplierFileName(supplier) {
    return `${sanitizeFileName(supplier || UNFILLED_SUPPLIER, UNFILLED_SUPPLIER)}_${getTodayStamp()}.xlsx`;
  }

  window.FAFULITools = window.FAFULITools || {};
  window.FAFULITools.supplierOrder = {
    id: "mall-order",
    group: "平台运营",
    name: "商城订单按供应商拆分",
    render,
    init
  };
})();
