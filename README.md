# China Payroll Workflow

一个本地优先、无云上传的中国月度工资表生成工具。它在电脑浏览器中读取上月已确认
工资表、目标月份的部门变动和个税、社保、公积金等附件，生成下一月份的待复核工资表。

项目适合财务人员日常使用，也适合由 Codex 等代码 Agent 持续适配未来新增或变化的部门
表单格式。

> 本项目提供工作流和技术实现，不构成税务、社保、劳动或会计专业意见。生成结果必须在
> 桌面版 Excel 中完整重算并由财务人员复核。

## 下载

请从 [GitHub Releases](https://github.com/oukeming64-tech/china-payroll-workflow/releases/latest)
下载最新版。每个版本提供两个压缩包：

| 压缩包 | 给谁使用 | 作用 |
|---|---|---|
| `china-payroll-workflow-universal-v1.0.0.zip`（通用版） | 领导、财务人员 | 解压后双击 `打开工资表工具.html`，直接生成待复核工资表 |
| `china-payroll-workflow-maintenance-v1.0.0.zip`（维护版） | 领导电脑上的 Codex 或开发人员 | 包含源码、合成测试、规则说明和重新打包工具，用于维护或增加表单适配 |

两个包都不包含真实工资表、员工姓名、身份证号、银行卡号或工资数值。

## 如果文件被误删

不需要找原发送人重新索取：

1. 打开项目主页：<https://github.com/oukeming64-tech/china-payroll-workflow>
2. 点击右侧 **Releases**，进入最新版本；
3. 日常使用重新下载“通用版”；
4. 需要 Codex 修改程序时重新下载“维护版”，也可以直接克隆本仓库。

```bash
git clone https://github.com/oukeming64-tech/china-payroll-workflow.git
cd china-payroll-workflow
```

## 业务模型

```text
上月已确认工资表
+ 目标月份的人员、工资、考勤、绩效等部门数据
+ 目标月份的个税、社保、公积金及适用附件
+ 已确认且可追溯的规则
→ 目标月份待复核工资表
```

- 目标月份固定为上月的下一个自然月；
- 目标为一月时，工具会另外要求上一年度 12 个月工资表；
- 普通月份接续已确认的累计数据，一月重新识别跨年结构与累计重置；
- 来源月份、字段或人员无法唯一匹配时停止，不静默填 0；
- 导出前解析已确认附件并移除外链；
- 人员和工资变动同步全部经验证的相关工作表；
- 输出始终是“待复核”，不能直接替代财务审批。

## 日常使用

1. 完整解压“通用版”，不要在压缩包预览窗口中运行；
2. 双击 `打开工资表工具.html`；
3. 选择上月已经确认的完整工资表；
4. 输入人员、工资、考勤、绩效或提成等目标月变化；
5. 选择目标月份的个税、社保、公积金及其他必要附件；
6. 逐项确认人员、字段和错误预览；
7. 生成后用桌面版 Excel 完整重算并复核。

支持 Windows、macOS 和常见桌面 Linux。建议使用最新版 Edge、Chrome 或 Safari。
手机微信和微信内置预览不属于支持的运行环境。

## 让 Codex 维护

让 Codex 先完整阅读：

1. [`AGENTS.md`](AGENTS.md)
2. [`CODEX.md`](CODEX.md)
3. [`requirements/PRODUCT_SPEC.md`](requirements/PRODUCT_SPEC.md)
4. [`requirements/ACCEPTANCE.md`](requirements/ACCEPTANCE.md)

然后提供新表单的空白模板或已脱敏样例，并说明希望识别的字段。不要把真实工资数据提交到
GitHub、Issue、测试夹具或公开聊天记录。

完整维护流程见 [`CODEX.md`](CODEX.md)。

## 开发与验证

需要 Node.js 20 或更高版本：

```bash
npm install
npx playwright install chromium
npm run check:architecture
npm test
npm run build
npm run package:maintenance
```

构建产物位于 Git 忽略的 `output/releases/`。测试只使用运行时生成的无个人信息合成
工作簿。

## 代码结构

```text
app/excel/  → Excel 与 OOXML 读取、写入和外链处理
app/rules/  → 月份、人员、附件和工资业务规则
app/ui/     → 非技术界面和流程编排
tests/      → 无个人信息合成端到端测试
tools/      → 架构检查和跨平台发布构建
```

依赖方向固定为 `excel → rules → ui`。未经工作簿、正式来源或用户确认的工资公式、社保
上下限、税率、取整或人员规则不得写入程序。

## 数据与隐私

- 所有工资文件只在当前电脑浏览器中读取；
- 程序没有登录、遥测、云存储或上传接口；
- 原文件只读，工具只下载新文件；
- 发布和测试禁止包含真实人员信息、工资数值、真实回放和内部工作簿指纹；
- 发现人员、字段、月份或公式缓存不安全时必须停止。

## 贡献

提交修改前请阅读 [`CONTRIBUTING.md`](CONTRIBUTING.md)。任何 Issue、PR、截图和测试
文件都不得包含真实人员或工资数据。

## 许可证

项目代码采用 [Apache License 2.0](LICENSE)。SheetJS 与 JSZip 按各自许可证分发，
详见 [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md)。版本变化见
[`CHANGELOG.md`](CHANGELOG.md)。
