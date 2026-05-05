# FAFULI 工具箱项目说明

本文件用于帮助后续维护者或 AI agent 快速理解当前项目。编码任务默认先理解现状，再做最小必要修改；不要顺手重构无关内容。

## 项目定位

- 项目名称：`FAFULI 工具箱`
- 使用方式：业务用户双击 `index.html` 即可打开。
- 运行方式：纯本地浏览器运行，无需网络连接、无需本地服务器、无需构建流程。
- 数据处理：Excel 读取、匹配、拆分和 ZIP 导出都在浏览器本地完成，不上传服务器。
- 当前只支持 `.xlsx`，默认读取第一个工作表。

## 当前目录结构

```text
FAFULI 工具箱/
├── index.html                 # 工具箱入口，只负责加载样式、依赖和模块
├── assets/
│   └── styles.css             # 公共 UI 样式和响应式布局
├── js/
│   ├── app.js                 # 工具注册、左侧目录渲染、工具切换
│   ├── excel-utils.js         # Excel/ZIP/文件名/日期/文本处理等公共方法
│   └── file-upload.js         # 上传、拖拽、删除、文件状态显示
├── tools/
│   ├── insurance-tool.js      # 保全方案编号匹配工具
│   └── supplier-order-tool.js # 商城订单按供应商拆分工具
├── vendor/
│   ├── xlsx.full.min.js       # 本地 Excel 读取和生成依赖
│   └── jszip.min.js           # 本地 ZIP 打包依赖
├── docs-assets/               # 用户操作手册截图素材
├── 用户操作手册.md             # 面向业务用户的操作文档
└── *.xlsx                     # 示例或参考表格
```

## 前端架构

- `index.html` 只保留页面骨架和脚本引用，脚本使用 `defer` 顺序加载。
- 加载顺序很重要：
  1. `vendor/xlsx.full.min.js`
  2. `vendor/jszip.min.js`
  3. `js/excel-utils.js`
  4. `js/file-upload.js`
  5. `tools/*.js`
  6. `js/app.js`
- 工具模块通过 `window.FAFULITools` 注册。
- 每个工具模块需要提供：
  - `id`：工具唯一标识，对应 `data-tool-panel`
  - `group`：左侧目录分组名称
  - `name`：左侧目录名称，也是业务工具名称
  - `render()`：返回工具 HTML
  - `init(root)`：绑定当前工具内部事件和状态
- `js/app.js` 会读取 `window.FAFULITools`，自动生成左侧目录和右侧工具面板。

## 公共模块说明

### `js/excel-utils.js`

提供公共工具方法：

- `parseWorkbook(file)`：读取 `.xlsx` 第一个工作表，返回表头和数据行。
- `findHeaderIndex()` / `findFirstHeaderIndex()`：按表头查找字段。
- `normalizeText()`：统一文本清理。
- `sanitizeFileName()`：清理文件名非法字符。
- `getTodayStamp()`：生成 `YYYYMMDD` 日期。
- `uniqueFileName()`：避免 ZIP 内文件重名。
- `downloadBlob()`：触发浏览器下载。
- `parseAmount()` / `formatAmount()`：金额解析和展示。
- `escapeHtml()`：渲染 HTML 前转义文本。

### `js/file-upload.js`

封装上传交互：

- 支持点击上传和拖拽上传。
- 上传成功后文件名显示蓝色。
- 支持删除已选文件。
- 解析失败时把错误交给对应工具模块处理。

## 已有工具

### 保全方案编号匹配工具

注册信息：

- `id`: `insurance`
- `group`: `保险运营`
- `name`: `保全方案编号匹配工具`

输入文件：

- 主表：保全名单，必须包含 `方案编号`、`保险公司`、`医保属地`、`与主险人关系`。
- 映射表：保险方案映射关系，必须包含 `方案编号`、`保险公司方案编号`。
- 映射表：医保属性映射关系，必须包含 `医保属地`、`医保属性`、`人保健康属性`。
- 映射表：特殊保险公司映射关系，必须包含 `特殊保险公司名称`、`保险公司名称` 或 `保险公司` 任一字段。

核心规则：

- 通过主表 `方案编号` 匹配映射表中的 `保险公司方案编号`。
- 在主表 `方案编号` 后新增 `保险公司方案编号`。
- 在主表 `医保属地` 后新增 `医保属性`。
- 未匹配的保险公司方案编号写 `未匹配`，导出 Excel 中高亮。
- 按主表 `保险公司` 拆分导出，一个保险公司一个 `.xlsx`。
- ZIP 文件名：`保全名单_按保险公司拆分_YYYYMMDD.zip`。
- 单个 Excel 文件名：`保险公司名称_YYYYMMDD.xlsx`。

医保属性规则：

1. `与主险人关系 = 子女` 时，医保属性直接写 `有医保`。
2. `保险公司 = 中国人民健康保险股份有限公司上海分公司` 时，使用医保映射表的 `人保健康属性`。
3. 保险公司命中特殊保险公司映射表时，也使用 `人保健康属性`。
4. 其他保险公司使用医保映射表的 `医保属性`。
5. 医保属地未匹配时写 `未匹配`。

特殊字段：

- 4 家指定保险公司会在导出表最后追加各自特殊字段。
- 阳光财产保险股份有限公司苏州中心支公司、中国人民健康保险股份有限公司上海分公司的 `年龄` 字段由主表 `被保险人出生日期` 计算。
- 以下主表字段需要排在特殊字段之后、表格最后：
  `方案特别约定`、`客服`、`客户名称`、`是否高风险人群`、`重疾名称`、`客服维护时间`、`集团户名称`、`保全提取时间`、`是否理赔高风险`。

### 商城订单按供应商拆分

注册信息：

- `id`: `mall-order`
- `group`: `平台运营`
- `name`: `商城订单按供应商拆分`

输入文件：

- 商城订单主表，必须包含 `供应商`、`结算总价`。

核心规则：

- 按 `供应商` 拆分订单。
- 预览展示 `供应商`、`订单行数`、`结算总价`、`文件名`。
- 预览按 `结算总价` 从高到低排序。
- 每个供应商导出一个 `.xlsx`，全部打包 ZIP。
- ZIP 文件名：`商城订单_按供应商拆分_YYYYMMDD.zip`。
- 单个 Excel 文件名：`供应商_YYYYMMDD.xlsx`。
- 导出表会删除指定敏感或无关字段，例如企业信息、会员信息、支付流水等，具体字段维护在 `ORDER_EXPORT_EXCLUDED_FIELDS`。

## UI 和交互约定

- 页面风格集中在 `assets/styles.css`，新增工具应复用现有 `.section`、`.dropzone`、`.status`、`.table-wrap`、`.actions` 等样式。
- 左侧目录宽度为 300px，右侧内容最大宽度为 1000px。
- 1024px 小屏电脑要避免整体横向滚动；表格允许在自身容器内横向滚动。
- `< 900px` 时左侧目录切到顶部。
- 不要新增深色模式；当前默认浅色。
- 文案面向业务用户，避免代码实现术语。

## 新增工具流程

1. 在 `tools/` 下新增一个工具文件，例如 `tools/new-tool.js`。
2. 工具内部用立即执行函数包裹，避免污染全局变量。
3. 从 `window.FAFULIExcelUtils` 和 `window.FAFULIFileUpload` 取公共方法。
4. 实现 `render()` 和 `init(root)`。
5. 注册到 `window.FAFULITools`。
6. 在 `index.html` 中按顺序增加 `<script defer src="./tools/new-tool.js"></script>`，位置应在 `js/app.js` 之前。
7. 复用现有 UI 样式，不要为单个工具重复写大量新样式。
8. 更新 `用户操作手册.md` 和 `docs-assets/` 截图。

工具注册示例：

```js
window.FAFULITools = window.FAFULITools || {};
window.FAFULITools.example = {
  id: "example",
  group: "分组名称",
  name: "工具名称",
  render,
  init
};
```

## 验证建议

修改后至少执行：

```bash
node --check js/app.js
node --check js/excel-utils.js
node --check js/file-upload.js
node --check tools/insurance-tool.js
node --check tools/supplier-order-tool.js
git diff --check
```

功能回归建议：

- 浏览器打开 `index.html`，确认左侧目录和默认工具正常显示。
- 切换 `保全方案编号匹配工具` 和 `商城订单按供应商拆分`。
- 使用示例 Excel 上传验证：
  - 保全工具应能通过校验并启用导出。
  - 商城订单工具应识别 38 个供应商、34033 条订单。
- 点击导出 ZIP，确认压缩包可生成。

注意：某些 Excel 在浏览器控制台可能出现 XLSX 依赖输出的压缩包提示信息；只要页面校验、预览和导出正常，不视为功能失败。

## 维护注意事项

- 不要改动 `vendor/` 依赖文件，除非明确升级本地依赖。
- 不要引入必须联网的 CDN、框架或构建流程。
- 不要把业务数据上传到外部服务。
- 不要删除示例文件，除非用户明确要求。
- `.DS_Store` 是 macOS 自动生成文件，通常不需要纳入版本管理。
- 如需改导出字段、表头规则、特殊保险公司规则，应优先在对应工具模块顶部常量中维护。

