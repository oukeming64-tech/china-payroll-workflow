# 工资表工具代码结构

这个目录是工资表月度生成应用。

## 依赖方向

```text
vendor/jszip + vendor/xlsx 0.20.3 + vendor/officecrypto-tool 0.0.19
    ↓
core/namespace
    ↓
excel/* → xlsx-engine.js
    ↓
rules/* → payroll-engine.js
    ↓
ui/*
    ↓
app.js
```

依赖只能向下，不允许规则层依赖界面，也不允许 Excel 层引用具体月份业务规则。
`index.html` 中的 `<script>` 顺序是显式依赖声明，由
`tools/check_app_architecture.mjs` 自动核对。

## 目录职责

- `core/namespace.js`：建立唯一全局命名空间。
- `excel/xml.js`：OOXML/XML 基础工具。
- `excel/table.js`：OOXML 或矩阵来源的字段头、行、单元格和公式节点识别。
- `excel/source-workbook.js`：统一读取`.xls`、`.xlsx`、`.xlsm`，识别并在内存中
  解密受密码保护的工作簿。
- `excel/docx-tables.js`：读取 Word 正文表格并转换为统一矩阵。
- `excel/mutations.js`：单元格、保留空白行和隐藏行操作。
- `excel/periods.js`：表内月份证据、月份标题和自然月滚动。
- `excel/external-links.js`：外链诊断与外部查找公式解析。
- `excel/external-detach.js`：验证外部公式已解析，固定安全缓存并移除外链包。
- `excel/formulas.js`：逐格读取公式并按字段归纳内部/外部来源。
- `excel/schema.js`：经工作簿证明的列插入、公式引用迁移和跨年字段升级。
- `excel/cumulative.js`：普通月份把上月累计缓存接入“之前月份累计”字段。
- `excel/personnel.js`：辅助表人员映射、同步、归档和安全停用。
- `excel/table-regions.js`：识别同一工作表中的转正、转岗、调薪等多个独立表格区段。
- `excel/workbook.js`：工作簿门面和导出。
- `rules/common.js`：人员、字段、月份、数值和冲突的公共规则。
- `rules/monthly-routes.js`：普通月份、一月跨年路由与全年历史验证。
- `rules/monthly-business.js`：一次性字段重置、季度绩效与保密补贴复核门禁。
- `rules/january-rollover.js`：一月 schema、累计重置和年度来源版本。
- `rules/attachment-resolution.js`：个税、社保/公积金、劳务附件字段与人员规则。
- `rules/attachment-batch.js`：不限制数量地合并同类附件覆盖、去重并识别冲突。
- `rules/monthly-change-sources.js`：声明员工动态、入职转正薪资、调薪和考勤分区。
- `rules/salary-event-proration.js`、`rules/salary-events.js`：转正、调薪工资构成与
  按天折算门禁。
- `rules/employment-events.js`：入职、离职工资预览和完整资料门禁。
- `rules/attendance-deductions.js`：病假、事假考勤扣款公式及需人工确认的假别。
- `rules/labor-fee-review.js`：劳务费金额差异提示与高税档停止规则。
- `rules/workbook-personnel.js`：完整工作簿人员/工资变动预检与同步契约。
- `rules/natural-language.js`：中文文字变动。
- `rules/tabular-changes.js`：CSV、长表和宽表变动。
- `rules/external-source.js`：来源字段到目标字段的证据匹配。
- `rules/business-sources.js`：实习生津贴、保密津贴和保密补贴审批的历史证据适配。
- `rules/attachment-periods.js`：附件月份证据与“文件名目标月、工作表名仍为上月”的
  受控兼容。
- `rules/social-base-july.js`：每年 7 月社保基数候选。
- `ui/state.js`：页面状态、脱敏和通用界面工具。
- `ui/password-flow.js`：加密工作簿密码弹窗、错误重试和取消；密码不进入页面状态
  或浏览器存储。
- `ui/render.js`：只负责渲染。
- `ui/personnel-sync-flow.js`：人员/工资变动在 6 表工作簿中的事务式应用。
- `ui/attachment-flow.js`：附件预览、事务式写入、回滚和零外链复检。
- `ui/source-regions-flow.js`：编排同一来源文件中的多个表格区段。
- `ui/requirements-render.js`：展示《工资表需求》11 项覆盖与人工边界。
- `ui/*-flow.js`：按业务流程编排，不复制底层规则。
- `xlsx-engine.js`、`payroll-engine.js`：兼容薄入口。
- `app.js`：启动和事件绑定。

## 新增或修改规则

1. 在 `rules/` 新建或窄改一个独立模块；
2. 通过 `window.PayrollLocal.rules` 导出纯函数和规则元数据；
3. 不在规则里读取 DOM，不在 UI 里重写规则；
4. 在 `index.html` 明确登记依赖顺序；
5. 增加可复现测试，覆盖正常结果、缺失信息和冲突；
6. 同步 `requirements/PRODUCT_SPEC.md`、`requirements/ACCEPTANCE.md`、README 和
   CHANGELOG；
7. 运行架构、浏览器、导出与发布包隐私检查。

## 月度入口契约

- 用户只选择上月完整工资表；
- 当前月份优先从主工资表标题读取，文件名只作为辅助证据；
- 目标月份固定为下一个自然月；
- 普通月份接续上月累计缓存，并要求目标月个税、社保/公积金附件；
- 目标为一月时必须读取上一年度 12 个月并通过历史、schema 和来源校验；
- 含`扣减税额`的一月另要求劳务费附件；
- 必要附件全部通过并确认写入后才允许导出，导出文件必须为零外链；
- 上月工资表本身是公式和样式骨架，不另设公式模板入口；
- 初次载入不得把上月工资表转成变动提案；
- 人员和工资变动由文字或表格入口接收，表格格式自动识别。
- 任一用户选择的加密工作簿都会弹出密码框；不同文件可使用不同密码，取消后不导入。
- 同批附件中只有一个文件加密时，输入正确密码后保留整批已读取结果；“已读取”和
  “通过核对”分开显示，人员差异继续要求先在变动入口确认。
- 目标月附件字段即使与主表当前值相同也会实际写入；是否沿用按来源判断，不按数值
  是否变化判断。
- 人员 / 工资变动入口支持已证明格式的业务 Excel 和 Word；只有人员名单、没有金额
  时明确停止，不推算工资或补贴。
- 目标月份附件一次选择多少份就读取多少份；有身份证的变动来源严格按身份证匹配。
- 劳务费与当前草案金额不同只显示中性复核提醒，不阻断写入；应纳税所得额超过
  20000 元且税档公式未证明时仍停止。

## 通用发布

运行`node tools/build_release.mjs`生成跨平台 ZIP。发布包只包含`app/`、根目录
`打开工资表工具.html`和使用说明；Windows、macOS 与桌面 Linux 均通过同一个本地
HTML 入口启动，不依赖 Node、Python 或平台专用应用。构建脚本会拒绝把 Excel、CSV
等数据文件装入发布包，产物只写入 Git 忽略的`output/releases/`。

运行`node tools/build_maintenance_release.mjs`生成 Codex 维护包。它只复制白名单内
的程序、无 PII 合成测试、维护说明和跨平台构建脚本，不包含内部文档索引、真实回归、
需求原件或 Git 历史；文本隐私扫描有任何命中都会停止。

## 文件大小门禁

- `rules/*.js`：最多 400 行；
- `excel/*.js`、`ui/*.js`：最多 500 行；
- 三个入口文件：最多 100 行。

如果一个模块接近门禁，应按“输入解析 / 规则计算 / 写入 / 展示”等单一职责拆分，
不能通过压缩代码绕过限制。

## 公式与隐私边界

- 不根据字段名猜工资公式；
- 保留模板内部公式；外部公式只用于识别来源，必须由目标月工资附件解析；
- 主表外部公式未解析或缓存不完整时停止，不提供强制断链开关；
- 导出前移除外部关系、外链部件和外部公式并复检为 0；
- 无法唯一匹配的人员或字段必须报错；
- 工资变化、人员新增或停用必须同步经验证的相关工作表；任一同步不安全时整批回滚；
- 默认界面脱敏；
- 工作簿密码只在当前一次读取调用中使用，不写入状态、日志、浏览器存储或导出文件；
- 不加入上传、遥测、登录或云端存储。
