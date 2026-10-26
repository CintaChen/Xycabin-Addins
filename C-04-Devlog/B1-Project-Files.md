**三项目档案与版本**

最后更新：2026-10-26 ｜ 当前版本：**0.9.9.3** ｜ 交付物：`Excel_Iterate_Setup.exe`、`WordUpdater_Setup.exe`、`XyCabin.html`

# 一. Agent及开发者接手概览

本仓库维护三个输出：WPS 文字加载项 WordUpdater、WPS 表格加载项 Excel-Iterate，以及两者共用的单文件离线网站 XyCabin。文档分为接手入口、使用说明、项目档案、开发经验、门禁五类。加载项主选项卡的“扩展→Agent扩展”显示 Agent-Evo，提供概览/全文、网站和自定配置。版本 0.9.9.3，构建不安装或操作用户 WPS。

## 1. 目录与阅读顺序

| 文档                                         | 内容                               |
| ------------------------------------------ | -------------------------------- |
| [B0-Entry.md](B0-Entry.md)                 | 入口与阅读顺序、项目关系、交付边界                |
| [A-UserGuide.md](A-UserGuide.md)           | 三产品使用说明                          |
| B1-Project-Files.md                        | 本章与三个项目档案：定位、版本摘要、命令按钮、源码地图、当前状态 |
| [B2-DevExperience.md](B2-DevExperience.md) | 共享开发经验与三个专属经验                    |
| \[B3-Gate Control.md\](B3-Gate Control.md) | 共享门禁与网站门禁                        |

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

版本来源与归并规则见\[门禁\](B3-Gate Control.md#1-文档编号与经验更新)，本次标记为0.9.9.3并沿用原固定发布者。

| 版本      | 重要变更                                                                                                                                                                                                 |
| ------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0.9.5.0 | 1. 首次提供WPS表格加载项，支持引用与从属追踪；2. 采用VSTO构建，提供单文件安装包。                                                                                                                                                      |
| 0.9.6.0 | 1. 完善引用与从属追踪；2. 统一窗口和图标；3. 增加运行库检查与固定证书签名，提升安装可靠性。                                                                                                                                                   |
| 0.9.7.0 | 1. 新增公式封装与计算、筛选增强、链接助手和工作表目录；2. 改善选区兼容与目录显示；3. 新增Agent-Evo选项卡、离线文档预览及网站入口。                                                                                                                           |
| 0.9.8.0 | 1. 新增文件整理与清单改名、链接清单保存替换，支持取消还原与确认清理；2. 改善筛选与反馈；3. Agent-Evo提供离线文档及树形尺寸设置；4. XyCabin提供离线下载、音乐和独立音量。                                                                                                   |
| 0.9.9.0 | 1. 文件整理、清单命名及多前缀OCR，按行断链、整簿恢复与编辑保护；2. 六列预览、布局回缩，即时配置及控件缩放，表头单行居中；3. 追踪引用新增显示错误来源：沿引用链定位错误源头格，可跳转标黄；4. 审计工具箱新增合并抵消：往来按单位对冲、坏账按比例冲回，收入成本逐行算未实现利润，列映射+分录预览+备份还原；5. 修复追踪弹窗锁死宿主：去除工具窗样式并对齐链接助手，根治抢焦点/失激活。 |
| 0.9.9.1 | 1. 合并抵消改净额法取数：资产端=债权方资产−负债、负债端=债务方负债−资产，债权方按净额较大方动态判定，净额保留真实符号；2. 公司对按实际往来自动配对，对方单位精确匹配，取消前缀匹配与别名归一；3. 差异表扩至15列，展示四类原始余额与净额过程，结果表跨表引用抵消金额；4. 修复未勾写数值时结果表金额引用错行。                                      |
| 0.9.9.2 | 1. 文档收敛为 C-04-Devlog 五份（B0 入口、A 使用说明、B1 档案、B2 经验、B3 门禁），项目内不留副本；2. 目录重构为 AA-Program/AB-Graphic，配置器更名 C1；3. 内嵌文档改按 C-04-Devlog 新路径读取，Agent-Evo 概览按章节标题与首段生成；4. 业务功能无变更。                               |
| 0.9.9.3 | 1. 开发/用户全文分别整份内嵌B2与A-UserGuide，导出文件名同源；2. 链接助手说明名称外链，按来源及作用域清理未使用名称后原生断链，保留整簿恢复；3. 断链指令改为“断开链接”，旧文本提示改填；4. 名称依赖预检区分A1地址与真实名称，保存后核对并准确报告失败状态；5. 名称外链状态改为“仅名称管理器定义引用（X 个名称）”。 |


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
| 链接助手    | Excel-Iterate·操作 | —                                                                                                                                               | 扫描并说明名称外链，C列填新路径或“断开链接”；名称清理先预检，整簿备份可恢复                                                         | OnLinkAssistant（Connect.cs:270）→ LinkAssistant + LinkBreakActions + LinkNameAnalysis                                                                                                                  |
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
| 链接/文件/OCR  | XlsxLinkPackage、LinkAssistant/BreakActions/Batch/OperationLog、LinkNameAnalysis、FileOperations/OperationBatch/RecycleFiles、FileOcr/Selection/CloudOcr、Connect.FileNames |
| 构建/安装      | build/rebuild/resign/gen_manifest/build_installer/verify_package/validate.py；src/installer及src/tools                                                  |
| 验证/网站      | src/testharness、src/tools；同级Website完整源码与父verify_shared.py                                                                                             |

单项目：`python rebuild.py --mode full` 后 `python validate.py --background`。联合文档/网页交付：Website `python rebuild.py --background --mode full`，随后 `python validate.py --background`。用已实际--version验证的解释器；新设备从[从零流程](B2-DevExperience.md#1-从零构建加载项)开始。

## 5. 当前状态与交付

### (1). 当前实现与待验收

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。新增表头匹配合并的接线、公式相对行规则与独立验收边界。

本轮修复链接助手的名称外链：来源可能仅存在于名称定义，扫描仍应列出并说明数量。LinkNameAnalysis按完整来源路径、外部引用顺序与名称作用域分析依赖；混合来源、仍被公式/对象使用或无法安全分析时在修改前停止。预检通过后沿用整簿备份，精确删除所选来源未使用名称，再调用宿主BreakLink并保存核对；不直接删除ZIP关系，不触碰其他来源。C列仅接受“断开链接”，旧文本提示改填。四列清单、旧批次及恢复边界沿用；实测结果、安装包及待验收项目见本轮交付证据。

名称外链扫描状态按用户要求改为“仅名称管理器定义引用（X 个名称）”，数量仍按实际名称定义显示；同一文案同步到离线/实机探针、使用说明及网站指南。已有清单中的旧文本在重新扫描时更新，安装与真实按钮验收单独记录。

补修短名称依赖误判：旧版将原始报表差异!Y$5误认为局部名称Y的使用，并把筛选范围$A$8:$Y$271误认为全局名称Y的使用。UsesName复用FormulaRefRewriter的A1地址跨度，核有效行列边界后排除地址中的名称字母；裸Y、Sheet!Y及超出合法地址边界的真实名称仍执行作用域与依赖保护。仅修改内部分析，公开命令、批次格式和恢复流程沿用；旧版失败、修复后离线与样本副本实机结果见本轮交付证据。

本批（0.9.9.0）修复追踪弹窗锁死宿主并经真实WPS实机验证：弹窗打开后可正常键入单元格、点开有子选项的功能区菜单；合并抵消由置灰占位转为可用：引擎/宿主/窗体三层拆分，写入前备份、可一键还原。收入成本抵消已改为矩阵式审计底稿（三维聚合+活公式+动态年度列，详见第五章(3)），离线探针144项锁定会计模型。底稿写入与还原待联机验收。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

### (2). 本轮交付证据

本轮单独更新Excel-Iterate 0.9.9.3的状态文案为“仅名称管理器定义引用（X 个名称）”，同步现有精确断言、实机探针及说明文档；其他状态、数量统计、依赖判断与断链逻辑沿用。本轮完整391项离线检查及24项后台任务通过，原证书重签、五载荷和内嵌资产校验通过。此前名称外链及A1误判修复的49项实机结果保留为历史证据，本轮文案调整未重复连接WPS；安装和真实按钮待人工验收。非审计正文在构建前固定，本节在全部验证后回填。

| 检查    | 结果                                                                                                                                                              |
| ----- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 安装器   | [Excel\_Iterate\_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，4,837,376字节，版本0.9.9.3，SHA-256前缀6280624B；[交付摘要](../C-01-ExcelIterate/build/name-status-label-delivery.json)记录安装包、DLL、两清单摘要、当前验证及边界 |
| 历史失败 | [离线旧DLL失败](../C-01-ExcelIterate/build/name-address-red.log)、[旧DLL复验](../C-01-ExcelIterate/build/name-address-red-confirm.log)及[实机旧DLL失败](../C-01-ExcelIterate/build/name-address-live-red.log)：此前实际生产会话在删除前报告相同的局部Y依赖，来源仍在；旧DLL与失败样本保留 |
| 历史定位 | [旧判断](../C-01-ExcelIterate/build/name-address-observation-red.log)记录局部Y 360处、全局Y 1处假依赖；[修复后](../C-01-ExcelIterate/build/name-address-observation-green.log)同一原件、同一内部分析均为0处。两份证据只读原件，本轮未重测 |
| 离线回归 | [background-validation.log](../C-01-ExcelIterate/build/background-validation.log)中的WorkbookToolsProbe为391 PASS，已使用新文案精确断言；包括此前的来源/作用域、地址/名称、指令、读取/保存失败、完整会话与恢复回归，本轮不新增重复测试 |
| 历史COM | [name-address-live-final.log](../C-01-ExcelIterate/build/name-address-live-final.log)：此前正式DLL 49项PASS，含审核2.XLS及其两个Y名称断链；其他来源、名称含筛选范围、数据表值/公式/已使用样式/几何均保留，保存重开不重现、取消逐字节还原，原件摘要未变。本轮仅改文案，未重复实机操作 |
| 工具入口 | [NameLinksLiveProbe.cs](../C-01-ExcelIterate/src/tools/NameLinksLiveProbe.cs)以LiveProbeSafety管理临时副本；[NameLinksProbe.cs](../C-01-ExcelIterate/src/testharness/NameLinksProbe.cs)加入既有WorkbookToolsProbe，原四列清单、LinkAction及旧批次XML兼容 |
| 构建与后台 | [name-status-label-build.log](../C-01-ExcelIterate/build/name-status-label-build.log)完整编译/重签/打包PASS；[name-status-label-validation.log](../C-01-ExcelIterate/build/name-status-label-validation.log)完整24项后台任务PASS；使用validate.py --background --mode full --site-html build/site-embedded.snapshot.html |
| 签名与载荷 | [name-status-label-package.log](../C-01-ExcelIterate/build/name-status-label-package.log)：固定发布者、两清单SHA-256双签名、五载荷、两份完整文档、14份图片与内嵌HTML全部PASS；交付脚本核对五份非审计正文在构建前后未变 |
| 前轮证据 | [名称残留交付](../C-01-ExcelIterate/build/name-link-delivery.json)、[地址误判交付](../C-01-ExcelIterate/build/name-address-delivery.json)、[地址修复后台原日志](../C-01-ExcelIterate/build/name-address-background-validation.log)及[原实机记录](../C-01-ExcelIterate/build/name-link-live-final.log)保留为历史证据；当前安装包及后台结果以name-status-label交付为准 |
| 文档与边界 | B1、B2、A使用说明及网站Excel摘要/指南已同步；最终按用户要求联合重建，两个安装包与网站正式产物一并更新，摘要以current-delivery.json为准。本次未自动安装，安装后真实按钮点击、焦点与弹窗外观待人工验收；不把实机COM或后台结果当作已完成安装验收 |

# 三. WordUpdater 项目档案

## 1. 定位与输出

WPS文字VSTO加载项：逐格链接Excel、更新/查看链接及选择正文顶层表格。曾用名WpsLinkUpdater，仅迁移自身旧身份。命名空间/程序集 `WordUpdater`，唯一安装器[AA-Program/WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02各保留完整公共源码，C-03是网站唯一源码及联合交付入口，不建立网站副本。

## 2. 版本摘要

版本来源与归并规则见\[门禁\](B3-Gate Control.md#1-文档编号与经验更新)，本次标记为0.9.9.3并沿用原固定发布者。

| 版本      | 重要变更                                                                                                                                      |
| ------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| 0.9.5.0 | 1. 首次提供WPS文字加载项，支持链接Excel数据；2. 采用VSTO构建，提供单文件安装包。                                                                                         |
| 0.9.6.0 | 1. 完善Excel链接、数据更新与查看功能；2. 统一窗口和图标；3. 增加运行库检查与固定证书签名，提升安装可靠性。                                                                              |
| 0.9.7.0 | 1. 按操作流程整理链接菜单，区分操作与反馈入口；2. 新增Agent-Evo选项卡；3. 支持离线阅读开发、使用文档及访问网站。                                                                         |
| 0.9.8.0 | 1. 完善链接建立、更新和查看；2. Agent-Evo提供离线开发/用户预览、全文应用选择和树形尺寸设置；3. XyCabin统一指南、离线下载、音乐和独立音量；4. 规范交付、文档与清理流程。                                        |
| 0.9.9.0 | 1. Agent-Evo同步排序、改名及控件缩放；2. 名称排序即时、格式尺寸下次开窗生效；3. 独立STA预览、窗口快照和需求记录；4. 通用确认隔离并完善Owner校验；5. 开发全文口令验证及离线配图缓存；6. 网格表头完整单行、居中且默认加粗。            |
| 0.9.9.1 | 1. 随联合交付重新打包，同步版本号与内嵌文档；2. WordUpdater 自身功能无变更。                                                                                           |
| 0.9.9.2 | 1. 查看链接窗体去除工具窗样式，修复 WPS 文字宿主单元格无法键入与功能区子菜单失效；2. 文档收敛为 C-04-Devlog 五份，内嵌文档与 Agent-Evo 概览同步新路径与新规则；3. 目录重构为 AA-Program/AB-Graphic，配置器更名 C2。 |
| 0.9.9.3 | 1. 开发/用户全文分别整份内嵌B2与A-UserGuide，导出文件名同源；2. 链接Excel仅校验所选规则矩形，允许选区外合并、拆分或嵌套；3. 保留实际目标Cell并在写入前复核，修复合并表头误拒与物理列号错位；4. 根据来源路径定位已打开的工作簿，兼容隐藏表格实例并恢复Ctrl多区域明细。 |

## 3. 当前命令按钮概览

本表以当前生成的功能区 XML 现状为准，逐行列出全部顶层命令按钮及菜单子项，唯一数据源为 `AA-Program/C2-Wordupdater_Configure.py` 的 `BUTTONS` 与 `EXTRA_TABS`；按钮、菜单或网址调整后须同步本表，Agent-Evo 为默认隐藏页，命令名称与顺序以用户自定配置为准。

| 命令按钮    | 选项卡与分组         | 子选项                                                                                                         | 功能摘要                                       | 实现来源                                                                                                                                               |
| ------- | -------------- | ----------------------------------------------------------------------------------------------------------- | ------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| 更新链接    | WordUpdater·链接 | —                                                                                                           | 一键批量更新所选范围内的全部链接                           | OnUpdateLinks（Connect.cs:236）→ UpdateEngine.cs                                                                                                     |
| 链接Excel | WordUpdater·链接 | —                                                                                                           | 按所选规则矩形逐格链接，允许选区外合并；按来源路径定位工作簿并读取Ctrl多选 | OnLinkToExcel（Connect.cs:498）→ LinkPasteEngine.cs + LinkPasteGeometry.cs + EtSelectionReader.cs                                                                            |
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

本轮修复“链接Excel”被选区外合并单元格阻止的问题：取消整表 Uniform/行列数拦截，按实际选中 Cell 与父表 XML 中对应单元格建立网格位置，选区自身仍须无合并、无嵌套且形成规则矩形；写入前复核同一批目标 Cell。离线失败回归与实机合并表头四格、左侧合并右侧四格已通过，字段引用、显示值、后续更新和选区外内容/格式均已核对。最终构建、签名与安装包证据见本轮交付证据；安装与真实按钮点击单独验收。

后续Ctrl多选反馈经新样本复现：默认Application实际指向Workbooks.Count=0的隐藏ET，来源工作簿在另一实例且只登记文件ROT；旧代码找到隐藏Application后漏查文件登记，窗口回读仍无来源。现无论是否已发现Application，都精确检查剪贴板来源路径对应的ROT登记，核对Workbook.FullName后通过Workbook.Application定位实际实例，再读全局/窗口选区；路径、表名、目标格数、精确外边界及不重叠约束继续生效，歧义拒绝。新CtrlSelectionLiveProbe由正常文件关联打开两个样本的临时副本，保留隐藏实例并存条件，覆盖第1/4行和第1/3行各6格到合并表头下2×3；原件不改动。安装与真实按钮点击仍单独记录。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

### (2). 本轮交付证据

10月24日本轮用用户新提供的两个样本副本复现并修复Ctrl来源实例漏查：上一版已安装文件摘要与前次交付一致，真实按钮仍报源12、目标6；默认COM Application为空的隐藏实例，而实际来源仅通过工作簿文件ROT可取得。修复后无论是否已有Application候选，都核对来源文件登记并取得真实Workbook.Application。沿用0.9.9.3、原证书及产品身份；最终交付按用户要求做联合重建，两个安装包与网站正式产物一并更新，Word 包因重新签名与打包导致摘要变化，已在最终摘要与本表中同步；项目档案、开发经验、使用说明与Word内嵌网站正文在构建前固定，构建后只回填本审计节。

| 检查    | 结果                                                                                                                                                      |
| ----- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 安装器   | [WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，4,235,776字节，版本0.9.9.3，SHA-256前缀039663FB；[最终摘要](../C-02-WordUpdater/build/source-instance-delivery.json)记录安装包、DLL、两清单完整摘要与验收状态 |
| 真实失败 | [source-instance-user-failure.json](../C-02-WordUpdater/build/source-instance-user-failure.json)：保留实际按钮12/6失败事件、安装DLL与旧交付一致的证据、样本摘要及根因排除过程；用户来源路径脱敏 |
| 失败回归 | [source-instance-red.log](../C-02-WordUpdater/build/source-instance-red.log)与[旧DLL复验](../C-02-WordUpdater/build/source-instance-red-confirm.log)：正常关联打开样本副本，默认实例0簿、真实实例1簿且选中6格；旧读取器来源候选0，完整引擎稳定返回SizeMismatch源12、目标6、写入0 |
| 离线检查 | [background-validation.log](../C-02-WordUpdater/build/background-validation.log)：映射67 PASS、读取及引用80 PASS、剪贴板边界19 PASS；新增来源工作簿归一、隐藏候选已存在、双反斜杠/错路径/接口失败及借用引用释放检查，保留原陈旧/错簿/错表/重叠/歧义与局部选区回归 |
| 新样本实机COM | [source-instance-live.log](../C-02-WordUpdater/build/source-instance-live.log)：正式DLL54项PASS；第1/4行及第1/3行分别形成6条LINK并全部刷新，逐格引用与显示值、后续更新、合并表头内容/格式和表格结构均通过，两个原件摘要未变；工具源码[CtrlSelectionLiveProbe.cs](../C-02-WordUpdater/src/testharness/CtrlSelectionLiveProbe.cs) |
| 既有实机回归 | [source-instance-live-regression.log](../C-02-WordUpdater/build/source-instance-live-regression.log)：正式DLL原112项全部PASS，涵盖整表、局部四格、Ctrl H1/H3、第2/4行8格、合并表头、左侧合并错列及更新；不以默认实例内创建样本代替本轮跨实例用例 |
| 构建及后台 | [source-instance-build.log](../C-02-WordUpdater/build/source-instance-build.log)完整构建/重签/打包PASS；[source-instance-validation-run.log](../C-02-WordUpdater/build/source-instance-validation-run.log)完整22项后台任务PASS；使用validate.py --background --mode full --site-html build/site-embedded.snapshot.html |
| 签名与载荷 | [source-instance-package.log](../C-02-WordUpdater/build/source-instance-package.log)：两清单SHA-256双签名、固定发布者、安装器五载荷、两份完整文档、14份图片与内嵌HTML全部PASS；交付记录核对构建前后的非审计正文摘要一致 |
| 验收边界 | 本轮新包未自动安装，安装后真实功能区按钮点击、键入焦点及弹窗外观待人工验收；正常文件关联打开副本后的实机COM通过，不冒充已完成新包按钮点击。本轮按用户要求联合交付，两个安装包与网站正式产物均已更新并同步为当前摘要 |

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

本次加载项内嵌文档改为整份导出：开发预览打开整份 B2-DevExperience.md，用户预览打开整份 A-UserGuide.md，不再分章裁剪、拼接或改写图片路径，导出文件名与源文件同名。加载项升至0.9.9.3并沿用原发布者；网站随联合交付重新生成，下载页携带 0.9.9.3 两包，站点版本摘要与 B1 版本表逐字一致，指南依据两使用说明同步。

源码与正式产物内容指纹控制增量；两包、独立页和交付记录统一事务发布。后台前端/下载/音频替身不发声、不真实下载，不能代替file页面、听感、浏览器许可或安装效果。当前交付摘要在审计区，方法不重复历史流水。

### (2). 本轮交付证据

10月20日发布 0.9.9.3：网站按新版本号重新生成，下载页携带 0.9.9.3 两包，站点版本摘要与 B1 版本表逐字一致；加载项内嵌文档改为整份导出，两个 DLL 网页同步重新打包。

| 检查   | 结果                                                                                                                                                                    |
| ---- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 独立网页 | [XyCabin.html](../C-03-Website·Xycabin/AA-Program/XyCabin.html)，13,321,994 字节                                                                                         |
| 站点验证 | [site-validation.log](../C-03-Website·Xycabin/build/site-validation.log)、[site-downloads-validation.log](../C-03-Website·Xycabin/build/site-downloads-validation.log) |
| 联合交付 | [current-delivery.json](../C-03-Website·Xycabin/build/current-delivery.json)                                                                                          |
