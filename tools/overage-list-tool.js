(() => {
  const {
    parseWorkbook,
    findHeaderIndex,
    findFirstHeaderIndex,
    normalizeText,
    getTodayStamp,
    downloadBlob,
    escapeHtml
  } = window.FAFULIExcelUtils;
  const { setupUpload } = window.FAFULIFileUpload;

  const RELATION_HEADER = "与主险人关系";
  const BIRTHDAY_HEADERS = ["被保人生日", "生日", "出生日期", "被保险人出生日期"];
  const NAME_HEADERS = ["被保险人姓名", "姓名"];
  const ID_HEADERS = ["被保险人证件号", "证件号", "身份证号"];
  const POLICYHOLDER_HEADERS = ["主险人姓名", "员工姓名"];
  const EMPLOYEE_RULE = "员工年龄大于等于65周岁";
  const CHILD_RULE = "子女年龄大于等于18周岁";

  const state = {
    main: null,
    result: null
  };
  let nodes = {};

  function render() {
    return `<header class="topbar">
          <div>
            <h1>超龄名单匹配工具</h1>
            <p class="tagline">上传人员清单，根据被保人生日筛选达到年龄条件的员工和子女并导出新表。</p>
          </div>
          <div class="local-badge">本地处理 无需网络连接</div>
        </header>

        <section class="section">
          <div class="section-head">
            <h2>1. 上传文件</h2>
            <span class="hint">仅支持 .xlsx，默认读取第一个工作表</span>
          </div>
          <div class="upload-grid single-upload-grid">
            <label class="dropzone" id="overageDrop">
              <input id="overageFile" type="file" accept=".xlsx">
              <div>
                <p class="drop-title">主表：人员清单</p>
                <p class="drop-copy">拖拽或点击上传，需包含「与主险人关系」「被保人生日」。</p>
              </div>
              <div class="file-state">
                <div class="file-name" id="overageFileName">未选择文件</div>
                <button class="clear-file" id="overageClear" type="button">删除</button>
              </div>
            </label>
          </div>
        </section>

        <section class="section">
          <div class="section-head">
            <h2>2. 匹配预览</h2>
            <span class="hint">员工导出 65 周岁及以上；子女导出 18 周岁及以上</span>
          </div>
          <div id="overageStatusBox" class="status">请先上传人员清单主表。</div>
          <ul id="overageErrorList" class="error-list hidden"></ul>
          <div id="overagePreviewEmpty" class="empty" style="margin-top: 14px;">暂无预览结果。</div>
          <div id="overagePreviewWrap" class="table-wrap hidden" style="margin-top: 14px;">
            <table>
              <thead>
                <tr>
                  <th>与主险人关系</th>
                  <th>被保险人姓名</th>
                  <th>被保人生日</th>
                  <th class="number">年龄</th>
                  <th>主险人姓名</th>
                  <th>匹配规则</th>
                </tr>
              </thead>
              <tbody id="overagePreviewBody"></tbody>
            </table>
          </div>
        </section>

        <section class="section">
          <div class="section-head">
            <h2>3. 导出文件</h2>
            <span class="hint">导出筛选后的人员新表</span>
          </div>
          <div class="actions">
            <span class="hint">导出表会保留原字段，并追加「年龄」「匹配规则」。</span>
            <button id="overageExportButton" type="button" disabled>导出 Excel</button>
          </div>
        </section>`;
  }

  function init(root) {
    nodes = {
      overageFile: root.querySelector("#overageFile"),
      overageFileName: root.querySelector("#overageFileName"),
      overageClear: root.querySelector("#overageClear"),
      overageStatusBox: root.querySelector("#overageStatusBox"),
      overageErrorList: root.querySelector("#overageErrorList"),
      overagePreviewEmpty: root.querySelector("#overagePreviewEmpty"),
      overagePreviewWrap: root.querySelector("#overagePreviewWrap"),
      overagePreviewBody: root.querySelector("#overagePreviewBody"),
      overageExportButton: root.querySelector("#overageExportButton")
    };

    setupUpload({
      input: nodes.overageFile,
      dropzone: root.querySelector("#overageDrop"),
      nameNode: nodes.overageFileName,
      clearButton: nodes.overageClear,
      parseWorkbook,
      onParsed: (parsed) => { state.main = parsed; },
      onCleared: () => { state.main = null; },
      onChange: processOverageFile
    });
    nodes.overageExportButton.addEventListener("click", exportOverageExcel);
  }

  function processOverageFile(errorMessage) {
    resetOverageViews();
    if (errorMessage) {
      setOverageStatus("error", errorMessage);
      return;
    }
    if (!state.main) {
      setOverageStatus("info", "请先上传人员清单主表。");
      return;
    }

    const relationIndex = findHeaderIndex(state.main.headers, RELATION_HEADER);
    const birthdayIndex = findFirstHeaderIndex(state.main.headers, BIRTHDAY_HEADERS);
    const nameIndex = findFirstHeaderIndex(state.main.headers, NAME_HEADERS);
    const idIndex = findFirstHeaderIndex(state.main.headers, ID_HEADERS);
    const policyholderIndex = findFirstHeaderIndex(state.main.headers, POLICYHOLDER_HEADERS);
    const errors = [];
    if (relationIndex === -1) errors.push("人员清单主表缺少表头：「与主险人关系」。");
    if (birthdayIndex === -1) errors.push("人员清单主表缺少表头：「被保人生日」。");
    if (errors.length) {
      renderOverageErrors(errors);
      setOverageStatus("error", "表头校验未通过，请修正后重新上传。");
      return;
    }

    const result = buildOverageResult(state.main, {
      relationIndex,
      birthdayIndex,
      nameIndex,
      idIndex,
      policyholderIndex
    });
    state.result = result;
    renderOveragePreview(result);
    setOverageStatus("ok", buildStatusText(result));
    nodes.overageExportButton.disabled = result.rows.length === 0;
  }

  function buildOverageResult(workbookData, indexes) {
    const asOfDate = startOfToday();
    const rows = [];
    const skipped = {
      invalidBirthday: 0,
      unsupportedRelation: 0,
      outOfRange: 0
    };
    let employeeCount = 0;
    let childCount = 0;

    workbookData.rows.forEach((row) => {
      const relation = normalizeRelation(row[indexes.relationIndex]);
      const birthday = parseBirthDate(row[indexes.birthdayIndex]) || parseBirthdayFromId(row[indexes.idIndex]);
      if (!birthday) {
        skipped.invalidBirthday += 1;
        return;
      }

      const age = calculateAge(birthday, asOfDate);
      let rule = "";
      if (isEmployeeRelation(relation)) {
        if (age < 65) {
          skipped.outOfRange += 1;
          return;
        }
        employeeCount += 1;
        rule = EMPLOYEE_RULE;
      } else if (isChildRelation(relation)) {
        if (age < 18) {
          skipped.outOfRange += 1;
          return;
        }
        childCount += 1;
        rule = CHILD_RULE;
      } else {
        skipped.unsupportedRelation += 1;
        return;
      }

      rows.push({
        relation,
        name: getCellValue(row, indexes.nameIndex),
        birthdayText: getCellValue(row, indexes.birthdayIndex),
        age,
        policyholder: getCellValue(row, indexes.policyholderIndex),
        rule,
        cells: workbookData.headers.map((_, index) => normalizeText(row[index]))
      });
    });

    return {
      headers: workbookData.headers,
      rows,
      employeeCount,
      childCount,
      skipped,
      totalRows: workbookData.rows.length,
      asOfDate
    };
  }

  function renderOveragePreview(result) {
    if (!result.rows.length) {
      nodes.overagePreviewEmpty.textContent = "没有符合条件的人员。";
      nodes.overagePreviewEmpty.classList.remove("hidden");
      return;
    }

    nodes.overagePreviewBody.innerHTML = result.rows.map((row) => `
      <tr>
        <td>${escapeHtml(row.relation)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${escapeHtml(row.birthdayText)}</td>
        <td class="number">${row.age}</td>
        <td>${escapeHtml(row.policyholder)}</td>
        <td>${escapeHtml(row.rule)}</td>
      </tr>
    `).join("");
    nodes.overagePreviewEmpty.classList.add("hidden");
    nodes.overagePreviewWrap.classList.remove("hidden");
  }

  function buildStatusText(result) {
    const skippedTotal = result.skipped.invalidBirthday + result.skipped.unsupportedRelation + result.skipped.outOfRange;
    return `已匹配 ${result.rows.length} 人：员工 ${result.employeeCount} 人，子女 ${result.childCount} 人；未保留 ${skippedTotal} 人。`;
  }

  async function exportOverageExcel() {
    if (!state.result || !state.result.rows.length) return;
    nodes.overageExportButton.disabled = true;
    nodes.overageExportButton.textContent = "正在导出...";
    try {
      const excelData = buildOverageWorkbook(state.result);
      const blob = new Blob([excelData], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });
      downloadBlob(blob, `超龄名单匹配结果_${getTodayStamp()}.xlsx`);
      nodes.overageExportButton.textContent = "导出 Excel";
    } catch (error) {
      setOverageStatus("error", `导出失败：${error.message}`);
      nodes.overageExportButton.textContent = "导出 Excel";
    } finally {
      nodes.overageExportButton.disabled = false;
    }
  }

  function buildOverageWorkbook(result) {
    const outputHeaders = [...result.headers, "年龄", "匹配规则"];
    const outputRows = result.rows.map((row) => [...row.cells, row.age, row.rule]);
    const worksheet = XLSX.utils.aoa_to_sheet([outputHeaders, ...outputRows]);
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    worksheet["!cols"] = outputHeaders.map((header, index) => ({
      wch: Math.min(Math.max(String(header).length + 4, index === 0 ? 16 : 10), 32)
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
    XLSX.utils.book_append_sheet(workbook, worksheet, "匹配结果");
    return XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      cellStyles: true
    });
  }

  function parseBirthDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    const text = normalizeText(value);
    if (!text) return null;

    const isoMatch = text.match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$/);
    if (isoMatch) return buildDate(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));

    const compactMatch = text.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (compactMatch) return buildDate(Number(compactMatch[1]), Number(compactMatch[2]), Number(compactMatch[3]));

    const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (slashMatch) {
      const year = normalizeYear(Number(slashMatch[3]));
      return buildDate(year, Number(slashMatch[1]), Number(slashMatch[2]));
    }

    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }
    return null;
  }

  function parseBirthdayFromId(value) {
    const text = normalizeText(value);
    const match = text.match(/^\d{6}(\d{4})(\d{2})(\d{2})[\dXx]$/);
    if (!match) return null;
    return buildDate(Number(match[1]), Number(match[2]), Number(match[3]));
  }

  function normalizeYear(year) {
    if (year >= 100) return year;
    const currentYear = new Date().getFullYear();
    const currentTwoDigitYear = currentYear % 100;
    return year > currentTwoDigitYear ? 1900 + year : 2000 + year;
  }

  function buildDate(year, month, day) {
    const date = new Date(year, month - 1, day);
    if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
      return null;
    }
    return date;
  }

  function calculateAge(birthday, asOfDate) {
    let age = asOfDate.getFullYear() - birthday.getFullYear();
    const monthDiff = asOfDate.getMonth() - birthday.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && asOfDate.getDate() < birthday.getDate())) {
      age -= 1;
    }
    return age;
  }

  function startOfToday() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate());
  }

  function normalizeRelation(value) {
    return normalizeText(value).replace(/\s+/g, "");
  }

  function isEmployeeRelation(relation) {
    return relation === "员工" || relation === "本人" || relation === "主险人";
  }

  function isChildRelation(relation) {
    return relation === "子女" || relation.includes("子女");
  }

  function getCellValue(row, index) {
    return index === -1 ? "" : normalizeText(row[index]);
  }

  function resetOverageViews() {
    nodes.overageErrorList.classList.add("hidden");
    nodes.overageErrorList.innerHTML = "";
    nodes.overagePreviewEmpty.textContent = "暂无预览结果。";
    nodes.overagePreviewEmpty.classList.remove("hidden");
    nodes.overagePreviewWrap.classList.add("hidden");
    nodes.overagePreviewBody.innerHTML = "";
    nodes.overageExportButton.disabled = true;
    state.result = null;
  }

  function setOverageStatus(type, message) {
    nodes.overageStatusBox.className = "status";
    if (type !== "info") {
      nodes.overageStatusBox.classList.add(type);
    }
    nodes.overageStatusBox.textContent = message;
  }

  function renderOverageErrors(errors) {
    if (!errors.length) {
      nodes.overageErrorList.classList.add("hidden");
      nodes.overageErrorList.innerHTML = "";
      return;
    }
    nodes.overageErrorList.innerHTML = errors.map((error) => `<li>${escapeHtml(error)}</li>`).join("");
    nodes.overageErrorList.classList.remove("hidden");
  }

  window.FAFULITools = window.FAFULITools || {};
  window.FAFULITools.overageList = {
    id: "overage-list",
    group: "客户服务",
    name: "超龄名单匹配工具",
    render,
    init
  };
})();
