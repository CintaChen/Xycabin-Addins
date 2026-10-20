**三项目档案与版本**

最后更新：2026-10-20 ｜ 当前版本：**0.9.9.1** ｜ 交付物：`Excel_Iterate_Setup.exe`、`WordUpdater_Setup.exe`、`XyCabin.html`

# 一. Agent及开发者接手概览

本仓库维护三个输出：WPS 文字加载项 WordUpdater、WPS 表格加载项 Excel-Iterate，以及两者共用的单文件离线网站 XyCabin。文档分为接手入口、使用说明、项目档案、开发经验、门禁五类。加载项主选项卡的“扩展→Agent扩展”显示 Agent-Evo，提供概览/全文、网站和自定配置。版本 0.9.9.1，构建不安装或操作用户 WPS。

## 1. 目录与阅读顺序

| 文档                                                 | 内容                               |
| -------------------------------------------------- | -------------------------------- |
| [B0-Entry.md](B0-Entry.md)                         | 入口与阅读顺序、项目关系、交付边界                |
| [A-UserGuide.md](A-UserGuide.md)                   | 三产品使用说明                          |
| B1-Project-Files.md                                | 本章与三个项目档案：定位、版本摘要、命令按钮、源码地图、当前状态 |
| [B2-DevExperience.md](B2-DevExperience.md)         | 共享开发经验与三个专属经验                    |
| [B3-Gate Control.md](B3-Gate Control.md) | 共享门禁与网站门禁                    |

## 2. 项目与交付物

| 项目            | 程序集           | 交付物                     |
| ------------- | ------------- | ----------------------- |
| Excel-Iterate | Excel_Iterate | Excel_Iterate_Setup.exe |
| WordUpdater   | WordUpdater   | WordUpdater_Setup.exe   |
| XyCabin       | —             | XyCabin.html            |

## 3. 源码与资源目录

三个项目统一为 `AA-Program`（配置器、启动器、正式产物）与 `AB-Graphic`（界面截图、功能区图标、示意 SVG）；源码在 `src/`，构建中间物与验证证据在 `build/`。文档只在仓库根 `C-04-Devlog` 维护，项目内不再保留副本。

# 二. Excel-Iterate 项目档案

## 1. 定位与输出

WPS表格VSTO加载项：引用/从属追踪、公式、筛选、链接、目录、文件整理/命名、文件名称/OCR、表头匹配合并及合并抵消。命名空间/程序集 `Excel_Iterate`，唯一安装器[AA-Program/Excel\_Iterate\_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02各保留完整公共源码，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见[门禁](B3-Gate Control.md#1-文档编号与经验更新)，本次标记为0.9.9.1并沿用原固定发布者。

| 版本      | 重要变更                                                                                                                                                                                                 |
| ------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0.9.5.0 | 1. 首次提供WPS表格加载项，支持引用与从属追踪；2. 采用VSTO构建，提供单文件安装包。                                                                                                                                                      |
| 0.9.6.0 | 1. 完善引用与从属追踪；2. 统一窗口和图标；3. 增加运行库检查与固定证书签名，提升安装可靠性。                                                                                                                                                   |
| 0.9.7.0 | 1. 新增公式封装与计算、筛选增强、链接助手和工作表目录；2. 改善选区兼容与目录显示；3. 新增Agent-Evo选项卡、离线文档预览及网站入口。                                                                                                                           |
| 0.9.8.0 | 1. 新增文件整理与清单改名、链接清单保存替换，支持取消还原与确认清理；2. 改善筛选与反馈；3. Agent-Evo提供离线文档及树形尺寸设置；4. XyCabin提供离线下载、音乐和独立音量。                                                                                                   |
| 0.9.9.0 | 1. 文件整理、清单命名及多前缀OCR，按行断链、整簿恢复与编辑保护；2. 六列预览、布局回缩，即时配置及控件缩放，表头单行居中；3. 追踪引用新增显示错误来源：沿引用链定位错误源头格，可跳转标黄；4. 审计工具箱新增合并抵消：往来按单位对冲、坏账按比例冲回，收入成本逐行算未实现利润，列映射+分录预览+备份还原；5. 修复追踪弹窗锁死宿主：去除工具窗样式并对齐链接助手，根治抢焦点/失激活。 |
| 0.9.9.1 | 1. 合并抵消改净额法取数：资产端=债权方资产−负债、负债端=债务方负债−资产，债权方按净额较大方动态判定，净额保留真实符号；2. 公司对按实际往来自动配对，对方单位精确匹配，取消前缀匹配与别名归一；3. 差异表扩至15列，展示四类原始余额与净额过程，结果表跨表引用抵消金额；4. 修复未勾写数值时结果表金额引用错行。                                      |


## 3. 当前命令按钮概览

本表以当前生成的功能区 XML 现状为准，逐行列出全部顶层命令按钮及菜单子项，唯一数据源为 `AA-Program/C1-Exceliterate_Configure.py` 的 `BUTTONS` 与 `EXTRA_TABS`；按钮、菜单或网址调整后须同步本表，Agent-Evo 为默认隐藏页，命令名称与顺序以用户自定配置为准。

| 命令按钮    | 选项卡与分组           | 子选项                                                                                                                                             | 功能摘要                                                                                    | 实现来源                                                                                                                                                               |
| ------- | ---------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 追踪引用    | Excel-Iterate·数据 | —                                                                                                                                               | 解析当前单元格公式，逐项展开被引用单元格的详情；错误行红字，选中后可「显示错误来源」回溯定位源头格并跳转                                    | OnTracePrecedents（Connect.cs:392）→ TrackerForm.cs；ErrorSourceEngine.cs/ErrorSourceForm.cs                                                                          |
| 追踪从属    | Excel-Iterate·数据 | —                                                                                                                                               | 遍历所有已打开工作簿，找出哪些单元格的公式引用了当前单元格                                                           | OnTraceDependents（Connect.cs:420）→ TrackerForm.cs                                                                                                                  |
| 筛选增强    | Excel-Iterate·数据 | 按列筛选、快捷高筛                                                                                                                                       | 按活动单元格所在表格的表头选择显示列                                                                      | OnFilterColumns（Connect.cs:277）→ ColumnFilterForm.cs；OnQuickAdvancedFilter（Connect.cs:256）→ WorkbookFeatures.cs                                                    |
| 审计工具箱   | Excel-Iterate·数据 | 表头匹配增加、合并抵消（往来抵消、收入成本抵消）                                                                                                                        | 按表头列名批量匹配多个工作簿，把列数据合并到目标工作簿；往来抵消按列映射生成抵消分录，收入成本抵消按三维聚合生成矩阵式底稿（借贷恒等式保证平衡），两者均写入前备份、可一键还原 | OnMergeHeaders（Connect.MergeHeaders.cs:24）→ MergeHeadersForm.cs；OnEliminationReceivable/OnEliminationRevenueCost（Connect.Elimination.cs:46/49）→ EliminationForm.cs |
| 一键公式    | Excel-Iterate·操作 | 进行公式封装（IFERROR封装〔默认返回空、自定义封装〕、ROUND封装〔默认保留2位、自定义封装〕、ABS封装、自定义函数）、移除公式封装、数值计算（加〔默认加100、自定义计算〕、减〔默认减100、自定义计算〕、乘〔默认乘100、自定义计算〕、除〔默认除10000、自定义计算〕） | 封装或移除当前选区公式的最外层函数                                                                       | 15 个回调均在 Connect.cs:197-228，委托 FormulaCommandEngine.cs                                                                                                             |
| 文件操作    | Excel-Iterate·操作 | 文件批量整理、文件批量命名、文件提取名称                                                                                                                            | 整理文件或通过工作表清单批量重命名                                                                       | OnOrganizeFiles/OnRenameFiles（Connect.cs:242/249）→ FileToolForms.cs；OnFileNames（Connect.FileNames.cs:422）                                                          |
| 链接助手    | Excel-Iterate·操作 | —                                                                                                                                               | 扫描 xlsx 外部引用，C列填新路径或「请断开该链接」后执行                                                         | OnLinkAssistant（Connect.cs:270）→ LinkAssistant.cs                                                                                                                  |
| 生成目录    | Excel-Iterate·操作 | —                                                                                                                                               | 链接全部工作表，隐藏表点击后显示并跳转                                                                     | OnGenerateDirectory（Connect.cs:263）→ WorkbookFeatures.cs                                                                                                           |
| 问题反馈    | Excel-Iterate·反馈 | —                                                                                                                                               | 有BUG或新增功能请发送邮箱（打开 QQ 邮箱网页版）                                                             | OnOpenUrl_btnFeedback（Connect.cs:453）→ RibbonConfig.GetUrl 字典                                                                                                      |
| Agent扩展 | Excel-Iterate·扩展 | —                                                                                                                                               | 显示或隐藏 Agent-Evo 选项卡（开关，按用户记住选择）                                                         | OnToggleAgentExtension（Connect.Agent.cs:131）                                                                                                                       |
| 调用API   | Agent-Evo·详情     | 智能云-API（百度智能云、阿里智能云）、Agent-API（国内Agent：WorkBuddy、TRAE、Qoder、DeepSeek、豆包、Kimi；国外Agent：ChatGPT、Claude、Gemini）                                     | 四级嵌套菜单，打开对应云控制台或 AI 网站（无 screentip，url 按钮）                                              | OnOpenUrl_btnAgentXxx×9 与 OnOpenUrl_btnAgentBaidu/Aliyun（Connect.Agent.cs）→ RibbonConfig.GetUrl 字典                                                                 |
| 开发流程    | Agent-Evo·详情     | 开发预览、用户预览、XyCabin                                                                                                                               | 打开内嵌开发/用户文档与网站首页                                                                        | OnOpenAgentReadme/OnOpenAgentUserGuide（Connect.Agent.cs:154/155）→ AgentDocumentForm.cs；OnAgentHome（Connect.Agent.cs:161）→ AgentLocalFiles.cs                       |
| 自定配置    | Agent-Evo·操作     | 自定排序、自定名称、自定格式、交互弹窗                                                                                                                             | Agent 弹窗的自定义项                                                                           | OnAgentSort/Names/Formats（Connect.Customization.cs:16-18）→ AgentCustomization.cs；OnOpenAgentDialogSizes（Connect.Agent.cs:147）→ DialogSizeEditor.cs                 |
| 实验计划    | Agent-Evo·操作     | OCR识别（置灰占位，未开放）、需求记录                                                                                                                            | 实验性功能入口                                                                                 | OnAgentOcrPlaceholder（Connect.Customization.cs:20，空实现）；OnAgentRequirements（Connect.Customization.cs:19）→ AgentCustomization.cs                                     |

## 4. 源码地图与运行入口

| 子系统        | 当前入口（未加目录者均在src/addin）                                                                                                                                |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| 身份/生成      | project.json、project_support.py、AA-Program/C1-Exceliterate_Configure.py、src/configuration/configure_core.py                                           |
| 加载/界面      | ThisAddIn、Connect及分部、Ribbon.g、InteropDefs、UiKit、HostDialogService                                                                                     |
| Agent/自定配置 | Connect.Agent/Customization、AgentDocumentForm/LocalFiles、AgentCustomization、DialogFormatting、FormatEditorThread/Session、CustomizationSaving           |
| 尺寸/预览      | DialogSizes、DialogSizeEditor、DialogPreviewFactory；src/tools/dialog_catalog_probe与DialogSizesProbe                                                     |
| 单元格/公式/筛选  | FormulaRefParser、Tracker引擎、ErrorSourceEngine/Form、FormulaCommandEngine/Interaction/WindowThread、WorksheetIdentity、ColumnVisibilityEngine              |
| 目录/合并      | WorkbookFeatures、FormulaRefRewriter、MergeHeadersEngine/Host/Form、EliminationEngine/Matrix/Diff/Host/Form、Connect.MergeHeaders、Connect.Elimination     |
| 链接/文件/OCR  | XlsxLinkPackage、LinkAssistant/BreakActions/Batch/OperationLog、FileOperations/OperationBatch/RecycleFiles、FileOcr/Selection/CloudOcr、Connect.FileNames |
| 构建/安装      | build/rebuild/resign/gen_manifest/build_installer/verify_package/validate.py；src/installer及src/tools                                                  |
| 验证/网站      | src/testharness、src/tools；同级Website完整源码与父verify_shared.py                                                                                             |

单项目：`python rebuild.py --mode full` 后 `python validate.py --background`。联合文档/网页交付：Website `python rebuild.py --background --mode full`，随后 `python validate.py --background`。用已实际--version验证的解释器；新设备从[从零流程](B2-DevExperience.md#1-从零构建加载项)开始。

## 5. 当前状态与交付

### (1). 当前实现与待验收

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。新增表头匹配合并的接线、公式相对行规则与独立验收边界。

本批（0.9.9.0）修复追踪弹窗锁死宿主并经真实WPS实机验证：弹窗打开后可正常键入单元格、点开有子选项的功能区菜单；合并抵消由置灰占位转为可用：引擎/宿主/窗体三层拆分，写入前备份、可一键还原。收入成本抵消已改为矩阵式审计底稿（三维聚合+活公式+动态年度列，详见第五章(3)），离线探针144项锁定会计模型。底稿写入与还原待联机验收。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

### (2). 本轮交付证据

10月20日完成门禁清单化：B3 更名为 B3-Gate Control.md，条款按 R-xx（门禁自动核对）与 H-xx（人工核对）分类，门禁总览上移至文档首节，删除与门禁无关的实现叙述、格式规格并入 R-27/R-28，失败语义改为受控词表并新增执行点静态校验；rebuild.py 构建前/发布前文档指纹双检与逐门禁运行摘要，三个入口 bat 补收尾 validate.py --background，workflow_probe 锁步三维失败矩阵。本轮功能代码与版本号未变，随联合交付重新打包。

| 检查    | 结果                                                                                                                                                             |
| ----- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 安装器   | [Excel\_Iterate\_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，3,928,576 字节，版本 0.9.9.1                                                   |
| 构建与验证 | [Excel\_Iterate-validation.log](../C-03-Website·Xycabin/build/Excel_Iterate-validation.log)、[payload](../C-03-Website·Xycabin/build/Excel_Iterate-payload.log) |
| 联合交付  | [current-delivery.json](../C-03-Website·Xycabin/build/current-delivery.json)                                                                                   |
| 文档门禁  | verify_devlog 五份 PASS 且 23 门禁摘要齐全；verify_shared 61 文件 PASS；workflow_probe 77 PASS；layout 14 项 PASS；document_assets 50 项 PASS                                                                      |

# 三. WordUpdater 项目档案

## 1. 定位与输出

WPS文字VSTO加载项：逐格链接Excel、更新/查看链接及选择正文顶层表格。曾用名WpsLinkUpdater，仅迁移自身旧身份。命名空间/程序集 `WordUpdater`，唯一安装器[AA-Program/WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02各保留完整公共源码，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见[门禁](B3-Gate Control.md#1-文档编号与经验更新)，本次标记为0.9.9.1并沿用原固定发布者。

| 版本      | 重要变更                                                                                                                           |
| ------- | ------------------------------------------------------------------------------------------------------------------------------ |
| 0.9.5.0 | 1. 首次提供WPS文字加载项，支持链接Excel数据；2. 采用VSTO构建，提供单文件安装包。                                                                              |
| 0.9.6.0 | 1. 完善Excel链接、数据更新与查看功能；2. 统一窗口和图标；3. 增加运行库检查与固定证书签名，提升安装可靠性。                                                                   |
| 0.9.7.0 | 1. 按操作流程整理链接菜单，区分操作与反馈入口；2. 新增Agent-Evo选项卡；3. 支持离线阅读开发、使用文档及访问网站。                                                              |
| 0.9.8.0 | 1. 完善链接建立、更新和查看；2. Agent-Evo提供离线开发/用户预览、全文应用选择和树形尺寸设置；3. XyCabin统一指南、离线下载、音乐和独立音量；4. 规范交付、文档与清理流程。                             |
| 0.9.9.0 | 1. Agent-Evo同步排序、改名及控件缩放；2. 名称排序即时、格式尺寸下次开窗生效；3. 独立STA预览、窗口快照和需求记录；4. 通用确认隔离并完善Owner校验；5. 开发全文口令验证及离线配图缓存；6. 网格表头完整单行、居中且默认加粗。 |
| 0.9.9.1 | 1. 随联合交付重新打包，同步版本号与内嵌文档；2. WordUpdater 自身功能无变更。                                                                                |

## 3. 当前命令按钮概览

本表以当前生成的功能区 XML 现状为准，逐行列出全部顶层命令按钮及菜单子项，唯一数据源为 `AA-Program/C2-Wordupdater_Configure.py` 的 `BUTTONS` 与 `EXTRA_TABS`；按钮、菜单或网址调整后须同步本表，Agent-Evo 为默认隐藏页，命令名称与顺序以用户自定配置为准。

| 命令按钮    | 选项卡与分组         | 子选项                                                                                                         | 功能摘要                                       | 实现来源                                                                                                                                               |
| ------- | -------------- | ----------------------------------------------------------------------------------------------------------- | ------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| 更新链接    | WordUpdater·链接 | —                                                                                                           | 一键批量更新所选范围内的全部链接                           | OnUpdateLinks（Connect.cs:236）→ UpdateEngine.cs                                                                                                     |
| 链接Excel | WordUpdater·链接 | —                                                                                                           | 把复制的单元格区域逐格链接到你选中的表格区域                     | OnLinkToExcel（Connect.cs:498）→ LinkPasteEngine.cs                                                                                                  |
| 查看链接    | WordUpdater·链接 | —                                                                                                           | 查看本文档引用了哪些Excel文件及其工作表/区域                  | OnViewLinks（Connect.cs:319）→ LinkViewForm.cs + LinkViewEngine.cs                                                                                   |
| 全选表格    | WordUpdater·操作 | —                                                                                                           | 选中本文档中的全部表格（正文不受影响）                        | OnSelectAllTables（Connect.cs:396）→ TableSelectEngine.cs                                                                                            |
| 问题反馈    | WordUpdater·反馈 | —                                                                                                           | 有BUG或新增功能请发送邮箱（打开 QQ 邮箱网页版）                | OnOpenUrl_btnFeedback（Connect.cs:170）                                                                                                              |
| Agent扩展 | WordUpdater·扩展 | —                                                                                                           | 显示或隐藏 Agent-Evo 选项卡（开关）                    | OnToggleAgentExtension（Connect.Agent.cs:131）                                                                                                       |
| 调用API   | Agent-Evo·详情   | 智能云-API（百度智能云、阿里智能云）、Agent-API（国内Agent：WorkBuddy、TRAE、Qoder、DeepSeek、豆包、Kimi；国外Agent：ChatGPT、Claude、Gemini） | 四级嵌套菜单，打开对应云控制台或 AI 网站（无 screentip，url 按钮） | OnOpenUrl_btnAgentXxx×9 与 OnOpenUrl_btnAgentBaidu/Aliyun（Connect.Agent.cs）→ RibbonConfig.GetUrl 字典                                                 |
| 开发流程    | Agent-Evo·详情   | 开发预览、用户预览、XyCabin                                                                                           | 打开内嵌开发/用户文档与网站首页                           | OnOpenAgentReadme/OnOpenAgentUserGuide（Connect.Agent.cs:154/155）→ AgentDocumentForm.cs；OnAgentHome（Connect.Agent.cs:161）→ AgentLocalFiles.cs       |
| 自定配置    | Agent-Evo·操作   | 自定排序、自定名称、自定格式、交互弹窗                                                                                         | Agent 弹窗的自定义项                              | OnAgentSort/Names/Formats（Connect.Customization.cs:16-18）→ AgentCustomization.cs；OnOpenAgentDialogSizes（Connect.Agent.cs:147）→ DialogSizeEditor.cs |
| 实验计划    | Agent-Evo·操作   | OCR识别（置灰占位，未开放）、需求记录                                                                                        | 实验性功能入口                                    | OnAgentOcrPlaceholder（Connect.Customization.cs:20，空实现）；OnAgentRequirements（Connect.Customization.cs:19）→ AgentCustomization.cs                     |

## 4. 源码地图与运行入口

| 子系统        | 当前入口（未加目录者均在src/addin）                                                                                                                      |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| 身份/生成      | project.json、project_support.py、AA-Program/C2-Wordupdater_Configure.py、src/configuration/configure_core.py                                  |
| 加载/界面      | ThisAddIn、Connect及分部、Ribbon.g、InteropDefs、UiKit、HostDialogService                                                                           |
| Agent/自定配置 | Connect.Agent/Customization、AgentDocumentForm/LocalFiles、AgentCustomization、DialogFormatting、FormatEditorThread/Session、CustomizationSaving |
| 尺寸/预览      | DialogSizes、DialogSizeEditor、DialogPreviewFactory；src/tools/dialog_catalog_probe与DialogSizesProbe                                           |
| LINK/选区    | LinkPasteGeometry/Engine、EtSelectionReader、TableSelectEngine                                                                                |
| 更新/查看/清理   | UpdateEngine、LinkViewEngine/Form、FlashHighlighter、EtCleanup                                                                                 |
| 构建/安装      | build/rebuild/resign/gen_manifest/build_installer/verify_package/validate.py；src/installer及src/tools                                        |
| 验证/网站      | src/testharness、src/tools；同级Website完整源码与父verify_shared.py                                                                                   |

单项目：`python rebuild.py --mode full` 后 `python validate.py --background`。联合文档/网页交付：Website `python rebuild.py --background --mode full`，随后 `python validate.py --background`。用已实际--version验证的解释器；新设备从[从零流程](B2-DevExperience.md#1-从零构建加载项)开始。

## 5. 当前状态与交付

### (1). 当前实现与待验收

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。已审阅Excel专属合并、OCR与文件批次，只同步界面/维护方法，不移植工作表业务。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

### (2). 本轮交付证据

10月20日完成门禁清单化：B3 更名为 B3-Gate Control.md，条款按 R-xx（门禁自动核对）与 H-xx（人工核对）分类，门禁总览上移至文档首节，删除与门禁无关的实现叙述、格式规格并入 R-27/R-28，失败语义改为受控词表并新增执行点静态校验；rebuild.py 构建前/发布前文档指纹双检与逐门禁运行摘要，三个入口 bat 补收尾 validate.py --background，workflow_probe 锁步三维失败矩阵。本轮功能代码与版本号未变，随联合交付重新打包。

| 检查    | 结果                                                                                                                                                      |
| ----- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 安装器   | [WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，3,279,872 字节，版本 0.9.9.1                                                  |
| 构建与验证 | [WordUpdater-validation.log](../C-03-Website·Xycabin/build/WordUpdater-validation.log)、[payload](../C-03-Website·Xycabin/build/WordUpdater-payload.log) |
| 联合交付  | [current-delivery.json](../C-03-Website·Xycabin/build/current-delivery.json)                                                                            |
| 文档门禁  | AgentDocumentCoreProbe 123 PASS；规则注册表双向校验 PASS；23 门禁运行摘要齐全                                                                             |

# 四. XyCabin 项目档案

## 1. 定位、关系与源码地图

网站介绍两个加载项，提供指南/版本/下载、开发流程、反馈与音乐；邮件/网盘需要网络。HTML内嵌CSS、JS、SVG、MP3和下载指引PNG，无CDN、外部字体、fetch邻近文件依赖。AA-Program/XyCabin.html独立网页含两EXE；两DLL资源是同一无EXE快照。Website不是第三个WPS加载项，不产第三份EXE。

| 当前源码                                         | 职责                            |
| -------------------------------------------- | ----------------------------- |
| AA-Program/C3-XycabinSite_Configure.py       | 唯一品牌、文案、指南、版本、颜色、声音和下载配置      |
| src/site_core.py、site_template.html          | 校验/转义/序列化、单页布局及通用交互           |
| src/site_audio.js、site_downloads.js          | 音频状态机、下载解码/菜单/弹窗              |
| src/assets/tranquility.mp3、unblock-guide.png | 必需项目内音乐片段与下载提示图，普通重建不依赖开发机原文件 |
| rebuild.py、src/delivery.py、build_cache.py    | 固定快照、两包暂存、联合验证/事务发布/回滚、内容缓存   |
| validate.py、src/tools                        | 前端、目录/文档、资源、事务、无旧build集成与清理探针 |

三个输出区统一AA-Program（唯一HTML、配置与批处理）、AB-Graphic（当前PNG与图标）；文档统一在C-04-Devlog维护；build留工具、收据与证据。两加载项不保留网站模板副本；相关输入路径按脚本位置计算。

## 2. 携带源码复现

完整交付带三个同级项目、父verify_shared.py、两加载项lib、原证书私钥和项目内资产；仅重建单加载项也带完整Website。环境：Windows、可执行Python≥3.10、Node.js、.NET Framework4.8/C#5。先实际--version核解释器，子进程参数列表/UTF-8；普通构建不需网络或FFmpeg。

预览运行 `python AA-Program/C3-XycabinSite_Configure.py`，只改HTML；仅使用成功交付记录中版本/大小/摘要匹配的包，无记录或变更时默认下载不可用。不要手改生成HTML或把配置器移出固定目录。

联合交付直接运行本项目AA-Program批处理：1增量、2完整（均--background），或 `python rebuild.py --background --mode full`，不必先跑两单项目批处理。随后 `python validate.py --background`；回填审计后 `python src/tools/layout_probe.py --strict-evidence`核引用。安装、真实file下载/播放、WPS点击单列验收。

## 3. 当前状态与交付

### (1). 当前实现与待验收

本次日志收为四章，网站门禁和前端经验分开，目录在概览外；移除外部历史范例依赖，纠正“相同HTML资源”与“相同DLL”、构建摘要验证与浏览器长度验证的层次。加载项保持0.9.9.1及原发布者，指南依据两使用说明同步。

源码与正式产物内容指纹控制增量；两包、独立页和交付记录统一事务发布。后台前端/下载/音频替身不发声、不真实下载，不能代替file页面、听感、浏览器许可或安装效果。当前交付摘要在审计区，方法不重复历史流水。

### (2). 本轮交付证据

10月20日完成门禁清单化：B3 更名为 B3-Gate Control.md，条款按 R-xx（门禁自动核对）与 H-xx（人工核对）分类，门禁总览上移至文档首节，删除与门禁无关的实现叙述、格式规格并入 R-27/R-28，失败语义改为受控词表并新增执行点静态校验；rebuild.py 构建前/发布前文档指纹双检与逐门禁运行摘要，三个入口 bat 补收尾 validate.py --background，workflow_probe 锁步三维失败矩阵。网站配置与模板未改，随联合交付重新生成。

| 检查   | 结果                                                                                                                                                                    |
| ---- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 独立网页 | [XyCabin.html](../C-03-Website·Xycabin/AA-Program/XyCabin.html)，10,833,632 字节                                                                                         |
| 站点验证 | [site-validation.log](../C-03-Website·Xycabin/build/site-validation.log)、[site-downloads-validation.log](../C-03-Website·Xycabin/build/site-downloads-validation.log) |
| 联合交付 | [current-delivery.json](../C-03-Website·Xycabin/build/current-delivery.json)                                                                                          |
