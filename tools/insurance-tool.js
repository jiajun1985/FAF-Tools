(() => {
  const {
    parseWorkbook,
    findHeaderIndex,
    findFirstHeaderIndex,
    normalizeText,
    sanitizeFileName,
    getTodayStamp,
    uniqueFileName,
    downloadBlob,
    escapeHtml
  } = window.FAFULIExcelUtils;
  const { setupUpload } = window.FAFULIFileUpload;

  const PLAN_HEADER = "方案编号";
  const ERP_CUSTOMER_HEADER = "公司ERP客编";
  const PRODUCT_CODE_HEADER = "产品编号";
  const COMPANY_HEADER = "保险公司";
  const COMPANY_PLAN_HEADER = "保险公司方案编号";
  const INSURED_BIRTHDAY_HEADER = "被保险人出生日期";
  const MEDICAL_LOCATION_HEADER = "医保属地";
  const EMPTY_MEDICAL_LOCATION_VALUE = "无";
  const RELATION_HEADER = "与主险人关系";
  const MEDICAL_ATTRIBUTE_HEADER = "医保属性";
  const PICC_HEALTH_ATTRIBUTE_HEADER = "人保健康属性";
  const CHILD_RELATION_VALUE = "子女";
  const PICC_HEALTH_COMPANY = "中国人民健康保险股份有限公司上海分公司";
  const SPECIAL_COMPANY_HEADERS = ["特殊保险公司名称", "保险公司名称", "保险公司"];
  const INSURANCE_MAPPING_MATCH_HEADERS = [
    ERP_CUSTOMER_HEADER,
    PLAN_HEADER,
    PRODUCT_CODE_HEADER,
    RELATION_HEADER,
    MEDICAL_ATTRIBUTE_HEADER,
    COMPANY_HEADER
  ];
  const UNFILLED_COMPANY = "未填写保险公司";
  const AGE_EXTRA_FIELD_COMPANIES = new Set([
    "阳光财产保险股份有限公司苏州中心支公司",
    "中国人民健康保险股份有限公司上海分公司"
  ]);
  const TRAILING_SOURCE_FIELDS = [
    "方案特别约定",
    "客服",
    "客户名称",
    "是否高风险人群",
    "重疾名称",
    "客服维护时间",
    "集团户名称",
    "保全提取时间",
    "是否理赔高风险"
  ];
  const EXTRA_FIELDS = {
    "百年人寿保险股份有限公司浙江分公司": ["被保人人序号", "保险计划", "份数", "职业代码", "职业类别", "国籍"],
    "亚太财产保险有限公司深圳分公司": ["保险责任", "赔付比例", "免赔额", "是否含一级医院", "一年期标准保费", "险种", "保单号", "批次", "批改信息", "批改保费", "生效止期", "投保项目", "既往症名称"],
    "阳光财产保险股份有限公司苏州中心支公司": ["年龄", "职业代码", "职业类别", "批改标志", "保司方案代码"],
    "中国人民健康保险股份有限公司上海分公司": ["状态", "年龄", "国籍", "生僻字标识", "职业代码", "职业类别", "岗位", "保险计划"]
  };

  const state = {
    main: null,
    mapping: null,
    medicalMapping: null,
    specialCompanyMapping: null,
    result: null
  };
  let nodes = {};

  function render() {
    return `<header class="topbar">
          <div>
            <h1>保全方案编号匹配工具</h1>
            <p class="tagline">上传主表和映射表，按保险公司拆分导出匹配后的 Excel 文件。</p>
          </div>
          <div class="local-badge">本地处理 无需网络连接</div>
        </header>

    <section class="section">
      <div class="section-head">
        <h2>1. 上传文件</h2>
        <span class="hint">仅支持 .xlsx，默认读取第一个工作表</span>
      </div>
      <div class="upload-grid">
        <label class="dropzone" id="mainDrop">
          <input id="mainFile" type="file" accept=".xlsx">
          <div>
            <p class="drop-title">主表：保全名单</p>
            <p class="drop-copy">拖拽或点击上传，需包含「公司ERP客编」「方案编号」「产品编号」「保险公司」。</p>
          </div>
          <div class="file-state">
            <div class="file-name" id="mainFileName">未选择文件</div>
            <button class="clear-file" id="mainClear" type="button">删除</button>
          </div>
        </label>
        <label class="dropzone" id="mapDrop">
          <input id="mapFile" type="file" accept=".xlsx">
          <div>
            <p class="drop-title">映射表：保险方案映射关系</p>
            <p class="drop-copy">拖拽或点击上传，需包含「公司ERP客编」「方案编号」「产品编号」「与主险人关系」「医保属性」「保险公司」「保险公司方案编号」。</p>
          </div>
          <div class="file-state">
            <div class="file-name" id="mapFileName">未选择文件</div>
            <button class="clear-file" id="mapClear" type="button">删除</button>
          </div>
        </label>
        <label class="dropzone" id="medicalMapDrop">
          <input id="medicalMapFile" type="file" accept=".xlsx">
          <div>
            <p class="drop-title">映射表：医保属性映射关系</p>
            <p class="drop-copy">拖拽或点击上传，需包含「医保属地」「医保属性」「人保健康属性」。</p>
          </div>
          <div class="file-state">
            <div class="file-name" id="medicalMapFileName">未选择文件</div>
            <button class="clear-file" id="medicalMapClear" type="button">删除</button>
          </div>
        </label>
        <label class="dropzone" id="specialCompanyMapDrop">
          <input id="specialCompanyMapFile" type="file" accept=".xlsx">
          <div>
            <p class="drop-title">映射表：特殊保险公司映射关系</p>
            <p class="drop-copy">拖拽或点击上传，需包含「特殊保险公司名称」或「保险公司名称」。</p>
          </div>
          <div class="file-state">
            <div class="file-name" id="specialCompanyMapFileName">未选择文件</div>
            <button class="clear-file" id="specialCompanyMapClear" type="button">删除</button>
          </div>
        </label>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <h2>特殊字段提示</h2>
        <span class="hint">以下字段会按保险公司追加在导出表最后，字段值默认留空</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>保险公司名称</th>
              <th>导出时追加的特殊字段</th>
            </tr>
          </thead>
          <tbody id="specialFieldsBody"></tbody>
        </table>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <h2>2. 按保险公司校验结果</h2>
        <span class="hint">分别展示每家保险公司的匹配情况</span>
      </div>
      <div id="statusBox" class="status">请先上传主表、保险方案映射表、医保属性映射表和特殊保险公司映射表。</div>
      <ul id="errorList" class="error-list hidden"></ul>
      <div id="summaryEmpty" class="empty" style="margin-top: 14px;">暂无校验结果。</div>
      <div id="summaryWrap" class="table-wrap hidden" style="margin-top: 14px;">
        <table>
          <thead>
            <tr>
              <th>保险公司名称</th>
              <th class="number">数据行数</th>
              <th class="number">匹配成功</th>
              <th class="number">未匹配</th>
            </tr>
          </thead>
          <tbody id="summaryBody"></tbody>
        </table>
      </div>
    </section>

    <section id="unmatchedSection" class="section hidden">
      <div class="section-head">
        <h2>3. 未匹配明细</h2>
        <span class="hint">这些记录仍会导出，新增列写「未匹配」并标红</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>保险公司名称</th>
              <th>方案编号</th>
              <th class="number">出现行数</th>
            </tr>
          </thead>
          <tbody id="unmatchedBody"></tbody>
        </table>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <h2>3. 导出文件</h2>
        <span class="hint">每个保险公司一个 Excel，全部打包成 ZIP</span>
      </div>
      <div id="exportEmpty" class="empty">完成校验后会显示导出清单。</div>
      <div id="exportWrap" class="table-wrap hidden">
        <table>
          <thead>
            <tr>
              <th>将导出文件</th>
              <th class="number">数据行数</th>
              <th>文件名</th>
            </tr>
          </thead>
          <tbody id="exportBody"></tbody>
        </table>
      </div>
      <div class="actions" style="margin-top: 16px;">
        <span class="hint">导出文件名会自动替换系统不允许的字符。</span>
        <button id="exportButton" type="button" disabled>导出 ZIP</button>
      </div>
        </section>
        `;
  }

  function init(root) {
    nodes = {
      mainFile: root.querySelector("#mainFile"),
      mapFile: root.querySelector("#mapFile"),
      medicalMapFile: root.querySelector("#medicalMapFile"),
      specialCompanyMapFile: root.querySelector("#specialCompanyMapFile"),
      mainFileName: root.querySelector("#mainFileName"),
      mapFileName: root.querySelector("#mapFileName"),
      medicalMapFileName: root.querySelector("#medicalMapFileName"),
      specialCompanyMapFileName: root.querySelector("#specialCompanyMapFileName"),
      mainClear: root.querySelector("#mainClear"),
      mapClear: root.querySelector("#mapClear"),
      medicalMapClear: root.querySelector("#medicalMapClear"),
      specialCompanyMapClear: root.querySelector("#specialCompanyMapClear"),
      statusBox: root.querySelector("#statusBox"),
      errorList: root.querySelector("#errorList"),
      summaryEmpty: root.querySelector("#summaryEmpty"),
      summaryWrap: root.querySelector("#summaryWrap"),
      summaryBody: root.querySelector("#summaryBody"),
      unmatchedSection: root.querySelector("#unmatchedSection"),
      unmatchedBody: root.querySelector("#unmatchedBody"),
      specialFieldsBody: root.querySelector("#specialFieldsBody"),
      exportEmpty: root.querySelector("#exportEmpty"),
      exportWrap: root.querySelector("#exportWrap"),
      exportBody: root.querySelector("#exportBody"),
      exportButton: root.querySelector("#exportButton")
    };

    setupUpload({
      input: nodes.mainFile,
      dropzone: root.querySelector("#mainDrop"),
      nameNode: nodes.mainFileName,
      clearButton: nodes.mainClear,
      parseWorkbook,
      onParsed: (parsed) => { state.main = parsed; },
      onCleared: () => { state.main = null; },
      onChange: processFiles
    });
    setupUpload({
      input: nodes.mapFile,
      dropzone: root.querySelector("#mapDrop"),
      nameNode: nodes.mapFileName,
      clearButton: nodes.mapClear,
      parseWorkbook,
      onParsed: (parsed) => { state.mapping = parsed; },
      onCleared: () => { state.mapping = null; },
      onChange: processFiles
    });
    setupUpload({
      input: nodes.medicalMapFile,
      dropzone: root.querySelector("#medicalMapDrop"),
      nameNode: nodes.medicalMapFileName,
      clearButton: nodes.medicalMapClear,
      parseWorkbook,
      onParsed: (parsed) => { state.medicalMapping = parsed; },
      onCleared: () => { state.medicalMapping = null; },
      onChange: processFiles
    });
    setupUpload({
      input: nodes.specialCompanyMapFile,
      dropzone: root.querySelector("#specialCompanyMapDrop"),
      nameNode: nodes.specialCompanyMapFileName,
      clearButton: nodes.specialCompanyMapClear,
      parseWorkbook,
      onParsed: (parsed) => { state.specialCompanyMapping = parsed; },
      onCleared: () => { state.specialCompanyMapping = null; },
      onChange: processFiles
    });

    nodes.exportButton.addEventListener("click", exportZip);
    renderSpecialFields();
  }

  function processFiles() {
      resetViews();
      if (!state.main || !state.mapping || !state.medicalMapping || !state.specialCompanyMapping) {
        setStatus("info", "请先上传主表、保险方案映射表、医保属性映射表和特殊保险公司映射表。");
        return;
      }

      const errors = [];
      const mainErpCustomerIndex = findHeaderIndex(state.main.headers, ERP_CUSTOMER_HEADER);
      const mainPlanIndex = findHeaderIndex(state.main.headers, PLAN_HEADER);
      const mainProductCodeIndex = findHeaderIndex(state.main.headers, PRODUCT_CODE_HEADER);
      const mainCompanyIndex = findHeaderIndex(state.main.headers, COMPANY_HEADER);
      const mainMedicalLocationIndex = findHeaderIndex(state.main.headers, MEDICAL_LOCATION_HEADER);
      const mainRelationIndex = findHeaderIndex(state.main.headers, RELATION_HEADER);
      const mapErpCustomerIndex = findHeaderIndex(state.mapping.headers, ERP_CUSTOMER_HEADER);
      const mapPlanIndex = findHeaderIndex(state.mapping.headers, PLAN_HEADER);
      const mapProductCodeIndex = findHeaderIndex(state.mapping.headers, PRODUCT_CODE_HEADER);
      const mapRelationIndex = findHeaderIndex(state.mapping.headers, RELATION_HEADER);
      const mapMedicalAttributeIndex = findHeaderIndex(state.mapping.headers, MEDICAL_ATTRIBUTE_HEADER);
      const mapCompanyIndex = findHeaderIndex(state.mapping.headers, COMPANY_HEADER);
      const mapCompanyPlanIndex = findHeaderIndex(state.mapping.headers, COMPANY_PLAN_HEADER);
      const medicalMapLocationIndex = findHeaderIndex(state.medicalMapping.headers, MEDICAL_LOCATION_HEADER);
      const medicalMapAttributeIndex = findHeaderIndex(state.medicalMapping.headers, MEDICAL_ATTRIBUTE_HEADER);
      const medicalMapPiccAttributeIndex = findHeaderIndex(state.medicalMapping.headers, PICC_HEALTH_ATTRIBUTE_HEADER);
      const specialCompanyIndex = findFirstHeaderIndex(state.specialCompanyMapping.headers, SPECIAL_COMPANY_HEADERS);

      if (mainErpCustomerIndex === -1) errors.push("主表缺少表头：「公司ERP客编」。");
      if (mainPlanIndex === -1) errors.push("主表缺少表头：「方案编号」。");
      if (mainProductCodeIndex === -1) errors.push("主表缺少表头：「产品编号」。");
      if (mainCompanyIndex === -1) errors.push("主表缺少表头：「保险公司」。");
      if (mainMedicalLocationIndex === -1) errors.push("主表缺少表头：「医保属地」。");
      if (mainRelationIndex === -1) errors.push("主表缺少表头：「与主险人关系」。");
      if (mapErpCustomerIndex === -1) errors.push("映射表缺少表头：「公司ERP客编」。");
      if (mapPlanIndex === -1) errors.push("映射表缺少表头：「方案编号」。");
      if (mapProductCodeIndex === -1) errors.push("映射表缺少表头：「产品编号」。");
      if (mapRelationIndex === -1) errors.push("映射表缺少表头：「与主险人关系」。");
      if (mapMedicalAttributeIndex === -1) errors.push("映射表缺少表头：「医保属性」。");
      if (mapCompanyIndex === -1) errors.push("映射表缺少表头：「保险公司」。");
      if (mapCompanyPlanIndex === -1) errors.push("映射表缺少表头：「保险公司方案编号」。");
      if (medicalMapLocationIndex === -1) errors.push("医保属性映射表缺少表头：「医保属地」。");
      if (medicalMapAttributeIndex === -1) errors.push("医保属性映射表缺少表头：「医保属性」。");
      if (medicalMapPiccAttributeIndex === -1) errors.push("医保属性映射表缺少表头：「人保健康属性」。");
      if (specialCompanyIndex === -1) errors.push("特殊保险公司映射表缺少表头：「特殊保险公司名称」「保险公司名称」或「保险公司」。");

      if (errors.length) {
        state.result = null;
        renderErrors(errors);
        setStatus("error", "表头校验未通过，请修正后重新上传。");
        return;
      }

      const mapping = buildMapping(state.mapping.rows, {
        erpCustomerIndex: mapErpCustomerIndex,
        planIndex: mapPlanIndex,
        productCodeIndex: mapProductCodeIndex,
        relationIndex: mapRelationIndex,
        medicalAttributeIndex: mapMedicalAttributeIndex,
        companyIndex: mapCompanyIndex,
        companyPlanIndex: mapCompanyPlanIndex
      });
      if (mapping.errors.length || mapping.conflicts.length) {
        state.result = null;
        renderErrors([...mapping.errors, ...mapping.conflicts]);
        setStatus("error", "保险方案映射关系校验未通过，请修正后重新上传。");
        return;
      }
      const medicalMapping = buildMedicalMapping(
        state.medicalMapping.rows,
        medicalMapLocationIndex,
        medicalMapAttributeIndex,
        medicalMapPiccAttributeIndex
      );
      const specialCompanies = buildSpecialCompanySet(state.specialCompanyMapping.rows, specialCompanyIndex);
      const output = buildOutput(
        state.main,
        {
          erpCustomerIndex: mainErpCustomerIndex,
          planIndex: mainPlanIndex,
          productCodeIndex: mainProductCodeIndex,
          companyIndex: mainCompanyIndex,
          medicalLocationIndex: mainMedicalLocationIndex,
          relationIndex: mainRelationIndex
        },
        mapping.map,
        medicalMapping,
        specialCompanies
      );
      state.result = {
        ...output,
        sheetName: state.main.sheetName,
        conflicts: []
      };

      renderResult(state.result);
    }

function buildMapping(rows, indexes) {
      const map = new Map();
      const conflicts = [];
      const errors = [];
      rows.forEach((row, index) => {
        const rowNumber = index + 2;
        const matchValues = getMappingMatchValues(row, indexes);
        const missingHeaders = INSURANCE_MAPPING_MATCH_HEADERS.filter((_, valueIndex) => !matchValues[valueIndex]);
        const companyPlanCode = normalizeText(row[indexes.companyPlanIndex]);
        if (missingHeaders.length) {
          errors.push(`保险方案映射表第 ${rowNumber} 行匹配字段为空：${missingHeaders.map((header) => `「${header}」`).join("、")}。`);
          return;
        }
        const key = buildMappingKey(matchValues);
        const existing = map.get(key);
        if (existing && existing.companyPlanCode !== companyPlanCode) {
          conflicts.push(
            `保险方案映射表第 ${rowNumber} 行与第 ${existing.rowNumber} 行匹配字段完全相同，`
            + `但「保险公司方案编号」不同（「${existing.companyPlanCode}」 / 「${companyPlanCode}」），请修正后重新上传。`
          );
          return;
        }
        if (!existing) {
          map.set(key, {
            rowNumber,
            companyPlanCode
          });
        }
      });
      return { map, errors, conflicts };
    }

function getMappingMatchValues(row, indexes) {
      return [
        normalizeText(row[indexes.erpCustomerIndex]),
        normalizeText(row[indexes.planIndex]),
        normalizeText(row[indexes.productCodeIndex]),
        normalizeText(row[indexes.relationIndex]),
        normalizeText(row[indexes.medicalAttributeIndex]),
        normalizeText(row[indexes.companyIndex])
      ];
    }

function buildMappingKey(values) {
      return values.map((value) => normalizeText(value)).join("\u0000");
    }

function buildMedicalMapping(rows, locationIndex, attributeIndex, piccAttributeIndex) {
      const map = new Map();
      rows.forEach((row) => {
        const location = normalizeText(row[locationIndex]);
        if (!location || map.has(location)) return;
        map.set(location, {
          normal: normalizeText(row[attributeIndex]),
          piccHealth: normalizeText(row[piccAttributeIndex])
        });
      });
      return map;
    }

function buildSpecialCompanySet(rows, companyIndex) {
      const set = new Set();
      rows.forEach((row) => {
        const company = normalizeText(row[companyIndex]);
        if (company) set.add(company);
      });
      return set;
    }

function buildOutput(main, indexes, mapping, medicalMapping, specialCompanies) {
      const {
        erpCustomerIndex,
        planIndex,
        productCodeIndex,
        companyIndex,
        medicalLocationIndex,
        relationIndex
      } = indexes;
      const outputHeaders = buildOutputHeaders(main.headers, planIndex, medicalLocationIndex);
      const companyPlanIndex = outputHeaders.indexOf(COMPANY_PLAN_HEADER);
      const groups = new Map();
      const unmatchedMap = new Map();

      main.rows.forEach((row) => {
        const planCode = normalizeText(row[planIndex]);
        const company = normalizeText(row[companyIndex]) || UNFILLED_COMPANY;
        const baseRow = main.headers.map((_, index) => normalizeText(row[index]));
        if (!baseRow[medicalLocationIndex]) {
          baseRow[medicalLocationIndex] = EMPTY_MEDICAL_LOCATION_VALUE;
        }
        const medicalAttributeValue = resolveMedicalAttribute({
          company,
          relation: baseRow[relationIndex],
          location: baseRow[medicalLocationIndex],
          medicalMapping,
          specialCompanies
        });
        const mappingKey = buildMappingKey([
          baseRow[erpCustomerIndex],
          planCode,
          baseRow[productCodeIndex],
          baseRow[relationIndex],
          medicalAttributeValue,
          company
        ]);
        const mappedItem = mapping.get(mappingKey);
        const matched = Boolean(mappedItem) && mappedItem.companyPlanCode !== "";
        const companyPlanValue = matched ? mappedItem.companyPlanCode : "未匹配";
        const outputRow = buildOutputRow(baseRow, planIndex, medicalLocationIndex, companyPlanValue, medicalAttributeValue);

        if (!groups.has(company)) {
          groups.set(company, {
            company,
            rows: [],
            total: 0,
            matched: 0,
            unmatched: 0
          });
        }

        const group = groups.get(company);
        group.rows.push({ cells: outputRow, matched, planCode });
        group.total += 1;
        if (matched) {
          group.matched += 1;
        } else {
          group.unmatched += 1;
          const detailKey = `${company}\u0000${planCode || "空方案编号"}`;
          if (!unmatchedMap.has(detailKey)) {
            unmatchedMap.set(detailKey, {
              company,
              planCode: planCode || "空方案编号",
              count: 0
            });
          }
          unmatchedMap.get(detailKey).count += 1;
        }
      });

      return {
        headers: outputHeaders,
        companyPlanIndex,
        groups: Array.from(groups.values()),
        unmatchedDetails: Array.from(unmatchedMap.values())
      };
    }

function buildOutputHeaders(headers, planIndex, medicalLocationIndex) {
      const outputHeaders = [];
      headers.forEach((header, index) => {
        outputHeaders.push(header);
        if (index === medicalLocationIndex) {
          outputHeaders.push(MEDICAL_ATTRIBUTE_HEADER);
        }
        if (index === planIndex) {
          outputHeaders.push(COMPANY_PLAN_HEADER);
        }
      });
      return outputHeaders;
    }

function buildOutputRow(row, planIndex, medicalLocationIndex, companyPlanValue, medicalAttributeValue) {
      const outputRow = [];
      row.forEach((value, index) => {
        outputRow.push(value);
        if (index === medicalLocationIndex) {
          outputRow.push(medicalAttributeValue);
        }
        if (index === planIndex) {
          outputRow.push(companyPlanValue);
        }
      });
      return outputRow;
    }

function resolveMedicalAttribute({ company, relation, location, medicalMapping, specialCompanies }) {
      if (normalizeText(relation) === CHILD_RELATION_VALUE) {
        return "有医保";
      }
      const mapping = medicalMapping.get(normalizeText(location));
      if (!mapping) return "未匹配";
      const shouldUsePiccHealthAttribute = company === PICC_HEALTH_COMPANY || specialCompanies.has(company);
      const value = shouldUsePiccHealthAttribute ? mapping.piccHealth : mapping.normal;
      return value || "未匹配";
    }


function renderResult(result) {
      renderSummary(result.groups);
      renderUnmatched(result.unmatchedDetails);
      renderExportList(result.groups);
      renderErrors([]);

      const unmatchedCount = result.groups.reduce((sum, group) => sum + group.unmatched, 0);
      if (unmatchedCount > 0) {
        setStatus("warn", `存在 ${unmatchedCount} 条未匹配，仍可导出；导出文件中会写「未匹配」并标红。`);
      } else {
        setStatus("ok", "已通过校验，可导出。");
      }
      nodes.exportButton.disabled = result.groups.length === 0;
    }

function renderSummary(groups) {
      nodes.summaryBody.innerHTML = groups.map((group) => `
        <tr>
          <td>${escapeHtml(group.company)}</td>
          <td class="number">${group.total}</td>
          <td class="number">${group.matched}</td>
          <td class="number">${group.unmatched}</td>
        </tr>
      `).join("");
      nodes.summaryEmpty.classList.add("hidden");
      nodes.summaryWrap.classList.remove("hidden");
    }

function renderUnmatched(details) {
      if (!details.length) {
        nodes.unmatchedSection.classList.add("hidden");
        nodes.unmatchedBody.innerHTML = "";
        return;
      }
      nodes.unmatchedBody.innerHTML = details.map((item) => `
        <tr>
          <td>${escapeHtml(item.company)}</td>
          <td>${escapeHtml(item.planCode)}</td>
          <td class="number">${item.count}</td>
        </tr>
      `).join("");
      nodes.unmatchedSection.classList.remove("hidden");
    }

function renderExportList(groups) {
      nodes.exportBody.innerHTML = groups.map((group) => `
        <tr>
          <td>${escapeHtml(group.company)}</td>
          <td class="number">${group.total}</td>
          <td>${escapeHtml(buildCompanyFileName(group.company))}</td>
        </tr>
      `).join("");
      nodes.exportEmpty.classList.add("hidden");
      nodes.exportWrap.classList.remove("hidden");
    }

function renderSpecialFields() {
      nodes.specialFieldsBody.innerHTML = Object.entries(EXTRA_FIELDS).map(([company, fields]) => `
        <tr>
          <td>${escapeHtml(company)}</td>
          <td>
            <div class="field-tags">
              ${fields.map((field) => `<span class="field-tag">${escapeHtml(field)}</span>`).join("")}
            </div>
          </td>
        </tr>
      `).join("");
    }

async function exportZip() {
      if (!state.result || !state.result.groups.length) return;
      nodes.exportButton.disabled = true;
      nodes.exportButton.textContent = "正在导出...";
      try {
        const zip = new JSZip();
        const usedNames = new Set();
        state.result.groups.forEach((group) => {
          const excelData = buildWorkbookForGroup(
            group,
            state.result.headers,
            state.result.companyPlanIndex,
            state.result.sheetName
          );
          const fileName = uniqueFileName(buildCompanyFileName(group.company), usedNames);
          zip.file(fileName, excelData);
        });
        const blob = await zip.generateAsync({ type: "blob" });
        downloadBlob(blob, `保全名单_按保险公司拆分_${getTodayStamp()}.zip`);
        nodes.exportButton.textContent = "导出 ZIP";
      } catch (error) {
        setStatus("error", `导出失败：${error.message}`);
        nodes.exportButton.textContent = "导出 ZIP";
      } finally {
        nodes.exportButton.disabled = false;
      }
    }

function buildWorkbookForGroup(group, baseHeaders, companyPlanIndex, sheetName) {
      const extraFields = EXTRA_FIELDS[group.company] || [];
      const birthdayIndex = baseHeaders.indexOf(INSURED_BIRTHDAY_HEADER);
      const sourceColumnOrder = buildSourceColumnOrder(baseHeaders);
      const headers = [
        ...sourceColumnOrder.normal.map((index) => baseHeaders[index]),
        ...extraFields,
        ...sourceColumnOrder.trailing.map((index) => baseHeaders[index])
      ];
      const rows = group.rows.map((item) => [
        ...sourceColumnOrder.normal.map((index) => item.cells[index]),
        ...buildExtraFieldValues(group.company, extraFields, item.cells, birthdayIndex),
        ...sourceColumnOrder.trailing.map((index) => item.cells[index])
      ]);
      const aoa = [headers, ...rows];
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      worksheet["!cols"] = headers.map((header, index) => ({
        wch: Math.min(Math.max(String(header).length + 4, index === 0 ? 14 : 10), 28)
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

      group.rows.forEach((item, index) => {
        if (item.matched) return;
        const reorderedCompanyPlanIndex = headers.indexOf(COMPANY_PLAN_HEADER);
        const cellRef = XLSX.utils.encode_cell({ r: index + 1, c: reorderedCompanyPlanIndex });
        if (!worksheet[cellRef]) return;
        worksheet[cellRef].s = {
          font: { bold: true, color: { rgb: "9C2F2A" } },
          fill: { fgColor: { rgb: "FCE4E2" } },
          alignment: { horizontal: "center", vertical: "center" }
        };
      });

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName || "Sheet1");
      return XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
        cellStyles: true
      });
    }

function buildSourceColumnOrder(headers) {
      const trailingSet = new Set(TRAILING_SOURCE_FIELDS);
      const normalColumns = [];
      const trailingColumns = [];

      headers.forEach((header, index) => {
        if (trailingSet.has(header)) {
          trailingColumns.push(index);
        } else {
          normalColumns.push(index);
        }
      });

      const orderedTrailingColumns = TRAILING_SOURCE_FIELDS
        .map((field) => headers.indexOf(field))
        .filter((index) => index !== -1);

      return {
        normal: normalColumns,
        trailing: orderedTrailingColumns
      };
    }

function buildExtraFieldValues(company, extraFields, row, birthdayIndex) {
      return extraFields.map((field) => {
        if (field === "年龄" && AGE_EXTRA_FIELD_COMPANIES.has(company)) {
          return calculateAge(row[birthdayIndex]);
        }
        return "";
      });
    }

function calculateAge(birthdayValue) {
      if (birthdayValue === null || birthdayValue === undefined || birthdayValue === "") return "";
      const birthday = parseDateValue(birthdayValue);
      if (!birthday) return "";

      const today = new Date();
      let age = today.getFullYear() - birthday.getFullYear();
      const monthDiff = today.getMonth() - birthday.getMonth();
      if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthday.getDate())) {
        age -= 1;
      }
      return age >= 0 ? age : "";
    }

function parseDateValue(value) {
      if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return value;
      }
      const text = normalizeText(value);
      if (!text) return null;

      const normalized = text
        .replace(/[./]/g, "-")
        .replace(/年|月/g, "-")
        .replace(/日/g, "")
        .split(" ")[0];
      const match = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (!match) return null;

      const year = Number(match[1]);
      const month = Number(match[2]);
      const day = Number(match[3]);
      const date = new Date(year, month - 1, day);
      if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
        return null;
      }
      return date;
    }

function resetViews() {
      nodes.errorList.classList.add("hidden");
      nodes.errorList.innerHTML = "";
      nodes.summaryEmpty.classList.remove("hidden");
      nodes.summaryWrap.classList.add("hidden");
      nodes.summaryBody.innerHTML = "";
      nodes.unmatchedSection.classList.add("hidden");
      nodes.unmatchedBody.innerHTML = "";
      nodes.exportEmpty.classList.remove("hidden");
      nodes.exportWrap.classList.add("hidden");
      nodes.exportBody.innerHTML = "";
      nodes.exportButton.disabled = true;
      state.result = null;
    }

function setStatus(type, message) {
      nodes.statusBox.className = "status";
      if (type !== "info") {
        nodes.statusBox.classList.add(type);
      }
      nodes.statusBox.textContent = message;
    }

function renderErrors(errors) {
      if (!errors.length) {
        nodes.errorList.classList.add("hidden");
        nodes.errorList.innerHTML = "";
        return;
      }
      nodes.errorList.innerHTML = errors.map((error) => `<li>${escapeHtml(error)}</li>`).join("");
      nodes.errorList.classList.remove("hidden");
    }

  function buildCompanyFileName(company) {
    return `${sanitizeFileName(company, UNFILLED_COMPANY)}_${getTodayStamp()}.xlsx`;
  }

  window.FAFULITools = window.FAFULITools || {};
  window.FAFULITools.insurance = {
    id: "insurance",
    group: "保险运营",
    name: "保全方案编号匹配工具",
    render,
    init
  };
})();
