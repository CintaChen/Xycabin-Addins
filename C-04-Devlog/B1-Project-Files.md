**三项目档案与版本**

最后更新：2026-10-28 ｜ 当前版本：**0.9.9.3** ｜ 交付物：`Excel_Iterate_Setup.exe`、`WordUpdater_Setup.exe`、`XyCabin.html`

# 一. Agent及开发者接手概览

本仓库维护三个输出：WPS 文字加载项 WordUpdater、WPS 表格加载项 Excel-Iterate，以及两者共用的单文件离线网站 XyCabin。文档分为接手入口、使用说明、项目档案、开发经验、门禁、代码索引六类。加载项主选项卡的“扩展→Agent扩展”显示 Agent-Evo，提供概览/全文、网站和自定配置。版本 0.9.9.3，构建不安装或操作用户 WPS。

## 1. 目录与阅读顺序

| 文档                                         | 内容                               |
| ------------------------------------------ | -------------------------------- |
| [B0-Entry.md](B0-Entry.md)                 | 入口与阅读顺序、项目关系、交付边界                |
| [A-UserGuide.md](A-UserGuide.md)           | 三产品使用说明                          |
| B1-Project-Files.md                        | 本章与三个项目档案：定位、版本摘要、命令按钮、源码地图、当前状态 |
| [B2-DevExperience.md](B2-DevExperience.md) | 共享开发经验与三个专属经验                    |
| \[B3-Gate Control.md\](B3-Gate Control.md) | 共享门禁与网站门禁                        |
| [B4-Code Index.md](<B4-Code Index.md>) | 两个加载项全部末级功能的目录树、入口、核心实现与共用依赖 |

## 2. 项目与交付物

| 项目            | 程序集           | 交付物                     |
| ------------- | ------------- | ----------------------- |
| Excel-Iterate | Excel_Iterate | Excel_Iterate_Setup.exe |
| WordUpdater   | WordUpdater   | WordUpdater_Setup.exe   |
| XyCabin       | —             | XyCabin.html            |

## 3. 源码与资源目录

三个项目统一为 `AA-Program`（配置器、启动器、正式产物）与 `AB-Graphic`（界面截图、功能区图标、示意 SVG）；功能源码在产品 `AC-Original Code/`，按选项卡、分组和菜单叶子归档；加载框架、生成源码及工具仍在 `src/`。公共实现只在同级 `C-00-Shared`，依赖在 `lib/`，构建中间物与验证证据在各产品 `build/`。完整功能导航只在 [B4-Code Index](<B4-Code Index.md>) 维护。文档只在仓库根 `C-04-Devlog` 维护，项目内不再保留副本。

# 二. Excel-Iterate 项目档案

## 1. 定位与输出

WPS表格VSTO加载项：引用/从属追踪、公式、筛选、链接、目录、文件整理/命名、文件名称/OCR、表头匹配合并及合并抵消。命名空间/程序集 `Excel_Iterate`，唯一安装器[AA-Program/Excel\_Iterate\_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02直接编译引用C-00-Shared公共实现，C-03是网站唯一源码及联合交付入口，不建立网站副本。

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

本表以配置器 BUTTONS 与 EXTRA_TABS 为准，保留顶层功能概览。逐项源码路径、回调与共用依赖统一见 [B4-Code Index](<B4-Code Index.md>)；Agent-Evo 默认隐藏，显示名称和排序可由用户自定。

| 命令按钮 | 选项卡与分组 | 子选项 | 功能摘要 |
| --- | --- | --- | --- |
| 追踪引用 | Excel-Iterate·数据 | — | 解析当前单元格公式，逐项展开被引用单元格的详情；错误行红字，选中后可「显示错误来源」回溯定位源头格并跳转 |
| 追踪从属 | Excel-Iterate·数据 | — | 遍历所有已打开工作簿，找出哪些单元格的公式引用了当前单元格 |
| 筛选增强 | Excel-Iterate·数据 | 按列筛选、快捷高筛 | 按活动单元格所在表格的表头选择显示列 |
| 审计工具箱 | Excel-Iterate·数据 | 表头匹配、合并抵消（往来抵消、收入成本抵消） | 按表头列名批量匹配多个工作簿，把列数据合并到目标工作簿；往来抵消按列映射生成抵消分录，收入成本抵消按三维聚合生成矩阵式底稿（借贷恒等式保证平衡），两者均写入前备份、可一键还原 |
| 一键公式 | Excel-Iterate·操作 | 进行公式封装（IFERROR封装（默认返回空、自定义封装）、ROUND封装（默认保留2位、自定义封装）、ABS封装、自定义函数）、移除公式封装、数值计算（加（默认加、自定义计算）、减（默认减、自定义计算）、乘（默认乘、自定义计算）、除（默认除、自定义计算）） | 封装或移除当前选区公式的最外层函数 |
| 文件操作 | Excel-Iterate·操作 | 批量处理（文件批量整理、文件批量命名、文件提取名称）、文件转换与 PDF 工具（当前文档转 PDF、批量文件转 PDF、PDF 添加水印、PDF 拆分、PDF 合并） | 文件批量整理、命名、名称提取；当前文档或批量转 PDF，以及 PDF 水印、拆分和合并。 |
| 链接助手 | Excel-Iterate·操作 | — | 扫描并说明名称外链，C列填新路径或“断开链接”；名称清理先预检，整簿备份可恢复 |
| 生成目录 | Excel-Iterate·操作 | — | 链接全部工作表，隐藏表点击后显示并跳转 |
| 问题反馈 | Excel-Iterate·反馈 | — | 有BUG或新增功能请发送邮箱（打开 QQ 邮箱网页版） |
| Agent扩展 | Excel-Iterate·扩展 | — | 显示或隐藏 Agent-Evo 选项卡（开关，按用户记住选择） |
| 调用API | Agent-Evo·详情 | 智能云-API（百度智能云、阿里智能云）、Agent-API（国内Agent（WorkBuddy、TRAE、Qoder、DeepSeek、豆包、Kimi）、国外Agent（ChatGPT、Claude、Gemini）） | 四级嵌套菜单，打开对应云控制台或 AI 网站（无 screentip，url 按钮） |
| 开发流程 | Agent-Evo·详情 | 开发预览、用户预览、XyCabin | 打开内嵌开发/用户文档与网站首页 |
| 自定配置 | Agent-Evo·操作 | 自定排序、自定名称、自定格式、交互弹窗 | Agent 弹窗的自定义项 |
| 实验计划 | Agent-Evo·操作 | OCR识别、需求记录 | 实验性功能入口 |

## 4. 源码地图与运行入口

| 子系统 | 位置与入口 |
| --- | --- |
| 功能实现 | AC-Original Code；完整定位见 [B4-Code Index](<B4-Code Index.md>)，机器映射在 project.json 的 features |
| 加载与生成 | src/addin 的 ThisAddIn、Connect、AssemblyInfo 及生成文件；AA-Program 配置器和 src/configuration |
| 跨产品共用 | C-00-Shared；产品只保留功能薄入口，经 @shared/ 直接编译引用共用实现 |
| 构建与验证 | 根目录构建脚本；src/installer、src/pdf、src/tools、src/testharness；父目录 verify_shared.py |
| 依赖与证据 | lib 为必要编译依赖；build 为中间物、候选包、缓存与验证证据 |

单项目：运行 python rebuild.py --background --mode full，再运行 python validate.py --background。源码结构迁移验证使用 --installer-output 将候选包写入 build，并通过 validate.py 的 --installer 指定候选包。联合交付仍使用 Website 入口。

## 5. 当前状态与交付

### (1). 当前实现与待验收

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。新增表头匹配合并的接线、公式相对行规则与独立验收边界。

本轮修复链接助手的名称外链：来源可能仅存在于名称定义，扫描仍应列出并说明数量。LinkNameAnalysis按完整来源路径、外部引用顺序与名称作用域分析依赖；混合来源、仍被公式/对象使用或无法安全分析时在修改前停止。预检通过后沿用整簿备份，精确删除所选来源未使用名称，再调用宿主BreakLink并保存核对；不直接删除ZIP关系，不触碰其他来源。C列仅接受“断开链接”，旧文本提示改填。四列清单、旧批次及恢复边界沿用；实测结果、安装包及待验收项目见本轮交付证据。

名称外链扫描状态按用户要求改为“仅名称管理器定义引用（X 个名称）”，数量仍按实际名称定义显示；同一文案同步到离线/实机探针、使用说明及网站指南。已有清单中的旧文本在重新扫描时更新，安装与真实按钮验收单独记录。

补修短名称依赖误判：旧版将原始报表差异!Y$5误认为局部名称Y的使用，并把筛选范围$A$8:$Y$271误认为全局名称Y的使用。UsesName复用FormulaRefRewriter的A1地址跨度，核有效行列边界后排除地址中的名称字母；裸Y、Sheet!Y及超出合法地址边界的真实名称仍执行作用域与依赖保护。仅修改内部分析，公开命令、批次格式和恢复流程沿用；旧版失败、修复后离线与样本副本实机结果见本轮交付证据。

本批（0.9.9.0）修复追踪弹窗锁死宿主并经真实WPS实机验证：弹窗打开后可正常键入单元格、点开有子选项的功能区菜单；合并抵消由置灰占位转为可用：引擎/宿主/窗体三层拆分，写入前备份、可一键还原。收入成本抵消已改为矩阵式审计底稿（三维聚合+活公式+动态年度列，详见第五章(3)），离线探针144项锁定会计模型。底稿写入与还原待联机验收。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

本轮沿用版本与发布身份，公共源码迁至C-00-Shared，项目清单直接引用，生成文件与产品入口保留原位置。抵消窗口在首次可见前应用最大化，来源事件先订阅再回填；局部缓冲覆盖布局、网格和拖柄，尺寸预览保持普通窗口。默认布局按窗口及DPI延迟测量；从属扫描和抵消预检通过原宿主线程分块读取、后台纯数据运算，来源变化时中止。宿主不支持变化监听时回退同步读取并记录兼容路径。EliminationSession统一持有目标、快照和任务，大结果预览按需取值，仍完整测宽、支持复制；取消与失败重试保留尚未还原的目标。实际WPS黑影、键入、菜单、跨屏DPI与写入还原由用户验收，离线检查不能代替。

### (2). 本轮交付证据

本轮分类门禁实现提交 [3cff8c5](https://github.com/Xymircai/Xycabin-Addins/commit/3cff8c5856b6d32c2fc20a10994dd2228cb80892)；构建源码与 README 提交 [bfccdec](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)，原有四个未推送提交完整保留并一并推送。详见 [本轮验证与发布记录](../C-03-Website·Xycabin/build/github-release-r30/delivery-review.md) 与 [完整证据摘要](../C-03-Website·Xycabin/build/github-release-r30/delivery-summary.json)。

本产品 54 个末级入口通过 [B3 R-30](<B3-Gate Control.md#6-功能源码分类与代码索引>) 与 [B4 代码索引](<B4-Code Index.md>) 核验。完整与增量模式均只读检查；55 项正反例包含 16 次实际入口违规阻断，不自动修正目录或索引。

| 检查 | 结果 |
| --- | --- |
| 安装包 | [Excel_Iterate_Setup.exe](../C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe)，4,871,680 字节，版本 0.9.9.3；SHA-256：7a61efb06f15a02d5946dea233e22749407d12767d7b82432e1fe47c57cdee9d |
| 重建与后台 | [联合完整重建](../C-03-Website·Xycabin/build/last-run-summary.json)与[收尾独立验证](../C-03-Website·Xycabin/build/validate-run-summary.json)成功；本产品 26 项完整后台任务通过、缓存复用 0；[产品验证日志](../C-03-Website·Xycabin/build/Excel_Iterate-validation.log) |
| 分类与编译 | [验证收据](../C-01-ExcelIterate/build/validation-timings.json)记录 R-30=pass、cached=false；目录、实际回调、编译清单及 B4 一致 |
| 签名与载荷 | [重建及签名日志](../C-03-Website·Xycabin/build/Excel_Iterate-rebuild.log)与[五载荷和资源核对](../C-03-Website·Xycabin/build/Excel_Iterate-payload.log)通过；沿用原发布者、清单双签名，文档与图形资源匹配 |
| 接口兼容 | [公开接口对照](../C-03-Website·Xycabin/build/github-release-r30/public-api-summary.json)：680 条记录保持一致，新增 0、删除 0 |
| GitHub 与 Release | 已普通推送至 main；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)附件已替换并回下载，摘要匹配本地正式产物；HTML 内嵌 EXE 与独立附件逐字节一致 |
| 源码链接 | [本轮精确提交](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)及[对应源码 ZIP](https://github.com/Xymircai/Xycabin-Addins/archive/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae.zip)；既有版本标签保持历史指向，自动 Source code 附件仍是旧标签源码 |
| 验收边界 | 实际 WPS 点击、宿主交互、安装效果、焦点与跨屏待实机验收；本轮未安装、连接或关闭 WPS，未重置用户配置 |

# 三. WordUpdater 项目档案

## 1. 定位与输出

WPS文字VSTO加载项：逐格链接Excel、更新/查看链接及选择正文顶层表格。曾用名WpsLinkUpdater，仅迁移自身旧身份。命名空间/程序集 `WordUpdater`，唯一安装器[AA-Program/WordUpdater\_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，内嵌五载荷；用户需完整VSTO运行库及官方授权。Agent-Evo同一DLL提供网站、离线文档、四项自定配置和需求记录，实验OCR占位灰色。

AA-Program放配置器、批处理与正式程序，AB-Graphic放界面截图、功能区图标和示意SVG；文档统一在仓库根C-04-Devlog维护；build是工具、中间产物和验证证据。C-01/C-02直接编译引用C-00-Shared公共实现，C-03是网站唯一源码及联合交付入口，不建立网站副本。

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

本表以配置器 BUTTONS 与 EXTRA_TABS 为准，保留顶层功能概览。逐项源码路径、回调与共用依赖统一见 [B4-Code Index](<B4-Code Index.md>)；Agent-Evo 默认隐藏，显示名称和排序可由用户自定。

| 命令按钮 | 选项卡与分组 | 子选项 | 功能摘要 |
| --- | --- | --- | --- |
| 更新链接 | WordUpdater·链接 | — | 一键批量更新所选范围内的全部链接 |
| 链接Excel | WordUpdater·链接 | — | 按所选规则矩形逐格链接，允许选区外合并；按来源路径定位工作簿并读取Ctrl多选 |
| 查看链接 | WordUpdater·链接 | — | 查看本文档引用了哪些Excel文件及其工作表/区域 |
| 全选表格 | WordUpdater·操作 | — | 选中本文档中的全部表格（正文不受影响） |
| 文件操作 | WordUpdater·操作 | 文件转换与 PDF 工具（当前文档转 PDF、批量文件转 PDF、PDF 添加水印、PDF 拆分、PDF 合并） | 当前文档或批量转 PDF，以及 PDF 水印、拆分和合并。 |
| 问题反馈 | WordUpdater·反馈 | — | 有BUG或新增功能请发送邮箱（打开 QQ 邮箱网页版） |
| Agent扩展 | WordUpdater·扩展 | — | 显示或隐藏 Agent-Evo 选项卡（开关） |
| 调用API | Agent-Evo·详情 | 智能云-API（百度智能云、阿里智能云）、Agent-API（国内Agent（WorkBuddy、TRAE、Qoder、DeepSeek、豆包、Kimi）、国外Agent（ChatGPT、Claude、Gemini）） | 四级嵌套菜单，打开对应云控制台或 AI 网站（无 screentip，url 按钮） |
| 开发流程 | Agent-Evo·详情 | 开发预览、用户预览、XyCabin | 打开内嵌开发/用户文档与网站首页 |
| 自定配置 | Agent-Evo·操作 | 自定排序、自定名称、自定格式、交互弹窗 | Agent 弹窗的自定义项 |
| 实验计划 | Agent-Evo·操作 | OCR识别、需求记录 | 实验性功能入口 |

## 4. 源码地图与运行入口

| 子系统 | 位置与入口 |
| --- | --- |
| 功能实现 | AC-Original Code；完整定位见 [B4-Code Index](<B4-Code Index.md>)，机器映射在 project.json 的 features |
| 加载与生成 | src/addin 的 ThisAddIn、Connect、AssemblyInfo 及生成文件；AA-Program 配置器和 src/configuration |
| 跨产品共用 | C-00-Shared；产品只保留功能薄入口，经 @shared/ 直接编译引用共用实现 |
| 构建与验证 | 根目录构建脚本；src/installer、src/pdf、src/tools、src/testharness；父目录 verify_shared.py |
| 依赖与证据 | lib 为必要编译依赖；build 为中间物、候选包、缓存与验证证据 |

单项目：运行 python rebuild.py --background --mode full，再运行 python validate.py --background。源码结构迁移验证使用 --installer-output 将候选包写入 build，并通过 validate.py 的 --installer 指定候选包。联合交付仍使用 Website 入口。

## 5. 当前状态与交付

### (1). 当前实现与待验收

本次将文档整理为五部分，共享方法与门禁分开，删重复流水和外部历史范例依赖；保留宿主独有算法/恢复边界，修正控件与四种拖动术语、窗口过度因果和内嵌网页一致性表述。已审阅Excel专属合并、OCR与文件批次，只同步界面/维护方法，不移植工作表业务。

本轮修复“链接Excel”被选区外合并单元格阻止的问题：取消整表 Uniform/行列数拦截，按实际选中 Cell 与父表 XML 中对应单元格建立网格位置，选区自身仍须无合并、无嵌套且形成规则矩形；写入前复核同一批目标 Cell。离线失败回归与实机合并表头四格、左侧合并右侧四格已通过，字段引用、显示值、后续更新和选区外内容/格式均已核对。最终构建、签名与安装包证据见本轮交付证据；安装与真实按钮点击单独验收。

后续Ctrl多选反馈经新样本复现：默认Application实际指向Workbooks.Count=0的隐藏ET，来源工作簿在另一实例且只登记文件ROT；旧代码找到隐藏Application后漏查文件登记，窗口回读仍无来源。现无论是否已发现Application，都精确检查剪贴板来源路径对应的ROT登记，核对Workbook.FullName后通过Workbook.Application定位实际实例，再读全局/窗口选区；路径、表名、目标格数、精确外边界及不重叠约束继续生效，歧义拒绝。新CtrlSelectionLiveProbe由正常文件关联打开两个样本的临时副本，保留隐藏实例并存条件，覆盖第1/4行和第1/3行各6格到合并表头下2×3；原件不改动。安装与真实按钮点击仍单独记录。

原命令ID、用户设置格式、程序集与安装路径沿用。名称/排序即刻刷新；格式/控件/窗口尺寸下次开窗生效，旧窗口保留快照。共享确认隔离、字体所有权与安全预览方法见第三章，不将后台通过认作真实WPS键入/菜单/焦点问题已解决。本轮安装、宿主点击和云端调用单独验收。

本轮沿用版本、原证书及稳定命令，公共实现改为C-00-Shared单源引用。更新链接与链接查看统一使用WordHostStateScope：完整取得提示、刷新状态后才修改，设置失败时恢复已触碰状态，恢复失败明确报告；移除诊断前固定等待，不关闭非自有进程。共享窗口尺寸和网格测宽采用统一生命周期与缓存。原逐格链接、Ctrl来源定位、表格结构、数值变化标黄及配置继续回归；本轮不连接实际宿主、不自动安装或关闭WPS。

### (2). 本轮交付证据

本轮分类门禁实现提交 [3cff8c5](https://github.com/Xymircai/Xycabin-Addins/commit/3cff8c5856b6d32c2fc20a10994dd2228cb80892)；构建源码与 README 提交 [bfccdec](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)，原有四个未推送提交完整保留并一并推送。详见 [本轮验证与发布记录](../C-03-Website·Xycabin/build/github-release-r30/delivery-review.md) 与 [完整证据摘要](../C-03-Website·Xycabin/build/github-release-r30/delivery-summary.json)。

本产品 31 个末级入口通过 [B3 R-30](<B3-Gate Control.md#6-功能源码分类与代码索引>) 与 [B4 代码索引](<B4-Code Index.md>) 核验。完整与增量模式均只读检查；55 项正反例包含 16 次实际入口违规阻断，不自动修正目录或索引。

| 检查 | 结果 |
| --- | --- |
| 安装包 | [WordUpdater_Setup.exe](../C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe)，4,248,576 字节，版本 0.9.9.3；SHA-256：597d5b0fa1e841f5e410921d67a48289a9aff4ca33a108f7830216b6b5a7f291 |
| 重建与后台 | [联合完整重建](../C-03-Website·Xycabin/build/last-run-summary.json)与[收尾独立验证](../C-03-Website·Xycabin/build/validate-run-summary.json)成功；本产品 24 项完整后台任务通过、缓存复用 0；[产品验证日志](../C-03-Website·Xycabin/build/WordUpdater-validation.log) |
| 分类与编译 | [验证收据](../C-02-WordUpdater/build/validation-timings.json)记录 R-30=pass、cached=false；目录、实际回调、编译清单及 B4 一致 |
| 签名与载荷 | [重建及签名日志](../C-03-Website·Xycabin/build/WordUpdater-rebuild.log)与[五载荷和资源核对](../C-03-Website·Xycabin/build/WordUpdater-payload.log)通过；沿用原发布者、清单双签名，文档与图形资源匹配 |
| 接口兼容 | [公开接口对照](../C-03-Website·Xycabin/build/github-release-r30/public-api-summary.json)：327 条记录保持一致，新增 0、删除 0 |
| GitHub 与 Release | 已普通推送至 main；[现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)附件已替换并回下载，摘要匹配本地正式产物；HTML 内嵌 EXE 与独立附件逐字节一致 |
| 源码链接 | [本轮精确提交](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)及[对应源码 ZIP](https://github.com/Xymircai/Xycabin-Addins/archive/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae.zip)；既有版本标签保持历史指向，自动 Source code 附件仍是旧标签源码 |
| 验收边界 | 实际 WPS 点击、宿主交互、安装效果、焦点与跨屏待实机验收；本轮未安装、连接或关闭 WPS，未重置用户配置 |

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

本轮分类门禁实现提交 [3cff8c5](https://github.com/Xymircai/Xycabin-Addins/commit/3cff8c5856b6d32c2fc20a10994dd2228cb80892)；构建源码与 README 提交 [bfccdec](https://github.com/Xymircai/Xycabin-Addins/tree/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae)，原有四个未推送提交完整保留并一并推送。详见 [本轮验证与发布记录](../C-03-Website·Xycabin/build/github-release-r30/delivery-review.md) 与 [完整证据摘要](../C-03-Website·Xycabin/build/github-release-r30/delivery-summary.json)。

三份产物同批联合完整重建，并通过收尾独立验证。B1 此审计区在构建完成、推送与远端附件回下载核对成功后回填，不改变内嵌文档正文。

| 检查 | 结果 |
| --- | --- |
| 独立网页 | [XyCabin.html](../C-03-Website·Xycabin/AA-Program/XyCabin.html)，13,384,792 字节；SHA-256：9e300daa5991206ef91b4d85cdc22da5c3baac68e8d170f820a6522486e60bdc |
| 联合重建 | [运行摘要](../C-03-Website·Xycabin/build/last-run-summary.json)状态 success，14 个阶段返回 0，G-04、G-08、G-09、G-17、G-19、G-20、G-21、G-22 通过；R-30=pass、cached=false |
| 收尾验证 | [独立验证摘要](../C-03-Website·Xycabin/build/validate-run-summary.json)状态 success，两个加载项 26+24 项后台任务通过且不复用缓存；签名、五载荷及内嵌文档/图形资源通过 |
| 附件一致性 | [发布验证状态](../C-03-Website·Xycabin/build/github-release-r30/publication-state.json)为 PUBLISHED_AND_DOWNLOADED_VERIFIED；三个远端附件已重新下载，SHA-256 与本地匹配，HTML 内嵌两包与独立附件逐字节一致，去除安装包数据后的网页与两 DLL 快照一致 |
| 发布记录 | [现有 Release](https://github.com/Xymircai/Xycabin-Addins/releases/tag/v0.9.9.3)说明已更新本轮变化、精确源码提交、[源码 ZIP](https://github.com/Xymircai/Xycabin-Addins/archive/bfccdecb9bc3dab0af0030def01d1f9aac6a07ae.zip)与三附件完整 SHA-256；原说明和附件已备份，未移动标签或强制推送 |
| 版本与边界 | 保持 0.9.9.3、原发布者及既有标签；自动 Source code 对应历史标签。本轮未自动安装或关闭宿主，实际 WPS 按钮和安装效果仍待实机验收 |
