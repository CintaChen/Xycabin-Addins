**共享维护纪律与网站维护纪律**

# 一. Word与Excel共享维护纪律

## 1. 文档、编号与经验更新

**R-01** C-04-Devlog 恰好五份：B0-Entry、A-UserGuide、B1-Project-Files、B2-DevExperience、B3-Uphold-Discipline；三个项目内不得保留文档副本。仓库根 AGENTS.md 只放指向 B0 的指针，B0 只放阅读顺序与按角色分流。

**R-14** 三个项目只保留 AA-Program（配置器、启动器、正式产物）、AB-Graphic（PNG、SVG、README.txt、.gitkeep）、src 与 build；AP1-程序、AP2-美工、AP3-说明、icons、src/documentation 出现即失败。

**R-02** 标题最多五级，同级连续编号、新父级重置，编号后一个半角空格；代码围栏内不参与编号。层级为 `# 一.` → `## 1.` → `### (1).` → `#### a.` → `##### i.`，不跳级、不为凑层级拆短句。

**R-03** 顶部只允许一处“最后更新”日期，正文不得出现其他 ISO 日期；B1 必填。

**R-04** B1 四个一级章节；Excel-Iterate 与 WordUpdater 两章各含“版本摘要”，版本四段升序、同章不重复、每行 trim 后不超过 200 字。

**R-05** A-UserGuide 恰好三章，依次为 Excel-Iterate、WordUpdater、XyCabin。

**R-06** B2 四个一级章节；B3 两个一级章节。

**R-07** 一级章节标题决定范围：含 Excel-Iterate→excel、WordUpdater→word、XyCabin→site；其余必须是登记过的共享标题（B1“Agent及开发者接手概览”、B2“Word与Excel共享开发经验”、B3“Word与Excel共享维护纪律”）。未知标题或同时命中多个产品即失败。

**R-08** 概览取文档标题与日期之后的首段；A-UserGuide 取一级标题前的导言段，B1 取首章后的首段；为空即构建失败。

**R-09** 每个项目档案恰好一处“### (n). 本轮交付证据”；合成内嵌文档前整节剔除，B1 共三处。

**R-10** 文档不得出现 HTML 注释，也不再使用 scope、agent-overview、delivery-audit 标记。

**R-11** 文档内本地链接与标题锚点必须可达；审计证据在源码携带时可缺，正式 strict-evidence 要求齐全。

**R-12** 文档图片只允许三个项目 AB-Graphic 下的 PNG、SVG；禁远程地址、绝对路径、越界路径与重解析点。

**R-13** 内嵌资源清单、文档与图片摘要必须与实际字节一致；两份内嵌文档各一份，图片按摘要去重。

**R-16** C-01 与 C-02 的公共源码、构建脚本与启动器逐字节一致；共享经验在 B2、B3 各只保留一份，不再成对维护。

**R-17** 版本唯一来源是各项目配置器的 INSTALLER_VERSION，四段 0～65534，仅用户指定才修改；安装器、加载项与清单同源；新第三段主版本把前一主版本子版本并入末段 0。

**（人工）** 维护顺序：对照最终源码→提炼做法、边界、入口、证据→判断归属→更新正文与入口→跑门禁→固定正文重建→回填证据。无新增可迁移经验如实说明，不编造根因。

## 2. 重建、签名与交付

现有单项目入口只重建自己，无 `--background` 参数：`python rebuild.py --mode full` 后 `python validate.py --background`。两产品与网页共同交付在 Website 运行 `python rebuild.py --background --mode full`（或其 AA-Program 批处理选 2），再 `python validate.py --background`；增量选 1/`--mode incremental`，默认 full。构建不安装、不结束 WPS；`--open-installer-only` 只打开既有安装器，不重建。

**R-18** 链路为配置器（C1-Exceliterate_Configure.py、C2-Wordupdater_Configure.py）→build.py→resign.py→build_installer.py；每次编译后必须重签，csc 的新 MVID 会改变 DLL 摘要。

**R-19** 安装器内嵌主 DLL、Utilities、Office.Core、应用清单、部署清单共五载荷，版本与 DLL、清单一致。

**R-20** 清单保留 customHostSpecified、GAC 依赖与 vstav3/vstov4 结构；ThisAddIn 作为 vstav3 入口，FullTrust 由 VSTO 运行时授予，不在应用清单中声明。未强签依赖不带 token；签后读顶层 identity 重建 vsto 引用，记录 manifest 精确大小与 SHA-256 Base64 摘要；UTF-8 BOM、CRLF。SignCmi 采用 SHA-256，两清单分别有 StrongName 与 Authenticode 签名；ValidateManifest 解析并转储清单结构，ValidateXmlSig 校验整清单及 license 绑定。

统一入口固定空 installer-assets 快照→两包暂存与后台验证→独立 HTML 附本轮包→核对正文、资源、包字节→事务替换两 EXE、HTML 和 current-delivery.json。失败保留上轮正式产物，回滚未完保留恢复文件并报路径；只生成配置预览不表示完成联合发布。

**（人工）** 信任缺失与签名损坏须分别报告；缺校验器或载荷不符阻止交付。安装需完整 VSTO Runtime 和官方 VSTOInstaller：RuntimeLocator 查 HKLM 两视图及 Common Files，缺少先停止，不释放新载荷、不写信任。用户保存文件并完全退出 WPS 后运行安装器，在官方窗口授权；注销本产品旧订阅成功才替换或重装，真实退出 0 才补 Manifest、LoadBehavior=3、AddinsWL。取消或非零退出如实报告，不静默安装。卸载只清自身订阅、挂载、白名单、信任与载荷，保留 LastInstallDir。

**（人工）** 排障按“运行库→SetupLogs→五载荷、挂载、白名单、订阅信任→签名摘要→真实宿主”。Manifest 是绝对 file URL 加 `|vstolocal`，codebase 相对；移动目录须重装。同版本不代表同字节，查安装路径与实际加载 DLL 摘要。COMAddIns.Connect 可作连接证据，Object=null、模块扫描或文件释放不能单独定启用。SetupDoctor 在 build，只读诊断不随正式输出交付。

## 3. 验证、缓存与证据边界

### (1). 用COM启动、附加和验证WPS

在 STA 执行：表格用 KET.Application 并创建 Workbooks.Add，文字用 KWPS.Application 并创建 Documents.Add。只读诊断用 Marshal.GetActiveObject 附加现有实例，失败返回 SKIP/INCOMPLETE，不启动应用；已授权需要启动时使用 Type.GetTypeFromProgID 加 Activator.CreateInstance（PowerShell 等价为 New-Object -ComObject），再设置 Visible 并创建测试对象，不能只打开 WPS 首页。附加模式只限定不启动宿主，不自动使业务只读；只读诊断不创建或修改文档。以下为现有 .NET Framework 4.8 探针中 ProbeSession 的开窗接线示例，外围 Main 标记 [STAThread]，捕获异常并用 session.Finish 记录结果；它不包含具体业务验收断言。PowerShell 直接调用 Marshal.GetActiveObject 应使用 Windows PowerShell 5.1，PowerShell 7 须采用适用的 ROT/原生接口实现。

```csharp
bool isSheet = true;                         // false：WPS文字
bool attachOnly = true;                      // 显式启动时才改为false
string progId = isSheet ? "KET.Application" : "KWPS.Application";
dynamic app = session.Take(ProbeSession.Connect(new[] { progId }, attachOnly));
if (app == null) return session.Finish(2);
string exe = Path.Combine(Convert.ToString(app.Path), isSheet ? "et.exe" : "wps.exe");
FileVersionInfo info = FileVersionInfo.GetVersionInfo(exe);
if ((info.ProductName ?? "").IndexOf("WPS", StringComparison.OrdinalIgnoreCase) < 0)
    throw new InvalidOperationException("无法确认WPS宿主，停止写入。");
session.Override(() => (object)app.Visible, v => app.Visible = v, true, "Visible");
dynamic items = session.Take(isSheet ? (object)app.Workbooks : (object)app.Documents);
object sample = items.Add();
session.Own(sample, delegate {
    if (isSheet) ((dynamic)sample).Close(false); else ((dynamic)sample).Close(0);
});
// 仅对sample及独立临时目录执行写入、保存、读回和恢复断言；不操作用户文档。
return session.Finish(2);                    // 仅演示开窗；补齐业务断言后才可返回0
```

注册映射和进程核对另按上节执行；示例中的文件产品检查不能代替完整身份记录。ProgID 不存在、指向不符、未发现现有实例或身份不足时明确失败或未完成；以 WPS 为目标的验收不静默改用 Excel.Application 或 Word.Application。现有业务探针的兼容回退链不是严格 WPS 启动契约，运行前须核对实际连接；附加失败可检查 ROT 文件 moniker，但只能选择明确的测试文件，不能任取第一个文档。GetActiveObject 也不保证拿到用户眼前那个窗口，目标用实际文档或工作簿身份确认。

CreateInstance 可能复用已有 Application，不能据此认领或 Quit 宿主。仅关闭自己创建并持有的测试文档或簿；所有自取 COM 属性、集合、范围逐次登记并 ReleaseComObject，避免链式访问产生遗漏，不 FinalRelease 共享 RCW。读取原值成功后才临时设置 Visible、DisplayAlerts、ScreenUpdating、EnableEvents 等，并在 finally 恢复；无法读取则不改。原文档 Saved 只观察，不设 true 掩盖修改；按需恢复原活动文档或簿、选区和剪贴板。失败返回非零并保留样本及日志，不按全局新增 PID 退出或强杀进程。实现入口为两项目 src/tools/LiveProbeSafety.cs 及对应实机探针。

COM 可验证真实宿主中的业务读写、字段或公式、保存与还原；不能代替按钮点击、键入焦点、弹窗外观、鼠标拖动及实际跨屏 DPI 验收。COMAddIns.Connect 仅说明连接状态；本项目未提供 COMAddIns.Object 自动化桥，Object=null 须记 INCOMPLETE，不能用新建 Connect 实例冒充已加载 Ribbon 回调。COM/API 检查不需要 computer-use；可见交互尚未执行时保留人工待验收状态。启动与附加 API 区别参见[New-Object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/new-object)、[GetActiveObject](https://learn.microsoft.com/en-us/dotnet/fundamentals/runtime-libraries/system-runtime-interopservices-marshal-getactiveobject)。

### (2). 弹窗维护规范

交互弹窗设错窗口样式会把宿主锁死：弹窗打开后无法在单元格键入、点不开有子选项的功能区菜单，而宿主窗口 enabled 仍为 True、宿主线程仍活着（功能区主按钮可反复点击），极易误判为宿主被禁用或线程阻塞。真根因是工具窗样式：设 `ShowInTaskbar = false`、`MinimizeBox = false` 会加上 WS_EX_TOOLWINDOW，无主非模态 tool 窗在 WPS 宿主线程上抢焦点或失激活——前台被弹窗把持，键盘输入进不了文档、功能区下拉点不开，宿主 enabled 与线程状态都不受影响，表现为“假死”。曾两度误判：先删无主窗的 TopMost，再把 Load 内业务 BeginInvoke 异步化，均未生效；按打开前后 enabled 均 True 排除宿主被禁用、功能区主按钮可反复点排除线程阻塞、前台窗口等于弹窗指向激活或焦点问题三重排除后，才锁定与链接助手等正常弹窗的样式差异。Word 与 Excel 两侧的交互弹窗都按本节构建。

正确构建交互弹窗：不设 `ShowInTaskbar = false`、`MinimizeBox = false`（即不引入 WS_EX_TOOLWINDOW），不重写 ShowWithoutActivation，不向无主窗设 TopMost，与链接助手等既有正常弹窗保持同套默认样式；非模态显示统一 `IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));`，WPS 回调线程上 owner 校验常不成立、取到 0 时走普通 Show 即可，样式正确即不锁死，不为拿 owner 改样式；Load 只做界面装配，重业务移出 Show 栈（BeginInvoke 到消息泵 idle），内置打开前后 enabled、前台、owner 诊断日志，复发时按三重排除定位，不再次误判；新弹窗在 DialogSizes、DialogPreviewFactory、dialog_catalog_probe、GridHeadersProbe 登记尺寸与目录，探针补源码守卫（禁 TopMost=true 与工具窗样式）和预览构造断言（ShowInTaskbar/MinimizeBox 保持默认 true），把根因钉进门禁；验收必须含真实宿主交互：开窗后键入单元格、展开功能区子菜单、Esc 关闭，后台探针通过不等于焦点与菜单验收。

## 4. 冗余文件清理

**（人工）** 只有用户实际要求“启动冗余清理”才触发盘点或删除；讨论规则、改文档或构建不触发历史清理。默认范围三个项目，另指定从其指令；项目外原图、用户配置、缓存不自动纳入。

**（人工）** 冗余候选须同时满足：生命周期结束或有有效再生办法；不在现行依赖闭包（含动态加载、资源、文档与证据引用）；独有经验或恢复信息已归并；隔离排除候选后从保留输入无旧 build 重建和适用验证通过。旧日期、old/backup 文件名、build 目录或搜不到引用都不足以证明可删。

**（人工）** 分五类盘点：必要输入、当前产物、有效证据、可删历史、未确认。按实际字节从大到小核对用途与摘要；公共副本职责不同不能因同 SHA 删除；移动项目内不算释放。私钥、`.cer` 证书、必要依赖与未完成恢复资料始终保护，公共 CA 的 `.pem` 不按扩展名误作私钥。

**R-23** 执行顺序：记录范围、版本、用途、数量、字节、摘要→归并经验修引用→隔离排除候选完成重建、签名、载荷、资源验证→Website `src/tools/cleanup.py --dry-run` 生成 schema 3 五类清单→逐项审阅→`--apply` 读同一清单复查→复核正式产物与链接并报告实际删除、失败、保留及字节差。清单后的输入或目标变化须重新盘点。

**（人工）** fixture-state 必须真实结束；cleanup-review 绑定源码指纹、逐文件摘要、用途与经验结论、隔离成功证据及原日志摘要，禁止伪造状态。删除前核绝对路径在指定项目内，不删项目根、不跟随联接、不碰项目外文件；不完整恢复保留。已授权确认项按范围执行。探针成功仅释放自建临时文件，失败保留诊断与恢复资料。

## 5. 格式维护规范

### (1). 弹窗中数据网格的格式规范

#### a. 格式规范说明

##### i. 表头格式规范

所有新增及现有生产弹窗的数据网格表头完整单行显示、居中、默认加粗。用户在“自定格式”中明确关闭粗体时尊重其选择；字号、字体、颜色和表头文字仍可自定，不能因此开启换行或截断表头。单行表头高度随实际字体、字号和 DPI 重新测量，避免上下裁剪；正文样式、网格稳定身份和已有用户配置保持，不通过重置配置取得默认外观。

![弹窗数据网格表头格式的离线控件渲染](../C-01-ExcelIterate/AB-Graphic/格式规范-弹窗网格表头-离线渲染.png)

上图来自“设置OCR前缀”的离线生产控件渲染，展示完整、单行、居中加粗的表头；非真实 WPS 截图，不作为宿主交互验收证据。

##### ii. 列宽与正文显示规范

列宽下限按当前实际表头字体、文字、内边距和排序图标测量，包含 Fill 列、复选列、下拉列与按钮列。手动缩窄不得低于完整表头宽度；宽度总和超过窗口时通过横向滚动查看，不压缩或换行表头。

文件提取名称正文自动适配上限仍为当前字体 30 个“0”加内边距，通常下限 50 逻辑像素，完整表头需要更宽时优先。该上限不裁剪数据、不限制手动调宽。新扫描重启适配周期；本周期后续重试、核对、规则调整与窗口缩放尊重已手调列宽。

#### b. 代码复现方法

源码链接在完整项目目录中使用；加载项离线全文保留下列片段，不携带业务源码。完整实现以链接中的函数为准；在同一加载项源码程序集内复用现有内部类型，修改公共网格代码后同步两个项目，不另维护一份格式实现。

| 用途            | 可复用的现有实现                                                                                                                                                                                     | 调用位置与边界                                                                                                                                                                           |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 网格默认样式及完整表头测量 | [UiKit.cs](../C-01-ExcelIterate/src/addin/UiKit.cs)：Ui.StyleGrid、Ui.AutoFitColumns、Ui.EnsureGridHeaders                                                                                                        | 创建列后适配；StyleGrid 在 HandleCreated 复核 DPI；EnsureGridHeaders 按逐列表头字体、Padding、排序图标计算下限与高度，不强制覆盖显式的非粗体字体                                                                                    |
| 自定格式应用后的收敛    | [DialogFormatting.cs](../C-01-ExcelIterate/src/addin/DialogFormatting.cs)：DialogFormatting.Apply                                                                                                               | finally 阶段完成布局后重新测宽；普通网格调用 Ui.AutoFitColumns，尺寸配置调用专属 FitColumns，不再开启表头换行                                                                                                            |
| 保留局部列宽策略      | [FileNamesLayout.cs](../C-01-ExcelIterate/src/addin/FileNamesLayout.cs)：FileNamesForm.FitResultColumns；[DialogSizeEditor.cs](../C-01-ExcelIterate/src/addin/DialogSizeEditor.cs)：DialogSizesForm.FitColumns | 文件提取先在 fittingColumns 保护内调用 EnsureGridHeaders，再处理正文 30 个 0 上限和手动列宽；尺寸目录保留树缩进及自绘按钮测量。文件提取通过 ResultGrid/IGridColumnFit 接线，局部回调不反调 Ui.AutoFitColumns；尺寸目录由 DialogFormatting.Apply 直接调用 FitColumns |

**最小调用示例。** 在已经登记并使用 DialogSizing.Prepare 的业务窗中，先创建全部列，将 grid 加入业务窗的既有布局容器，再调用共享入口：

```csharp
var grid = new DataGridView {
    Dock = DockStyle.Fill, ScrollBars = ScrollBars.Both,
    AllowUserToAddRows = false, RowHeadersVisible = false
};
Ui.StyleGrid(grid);                         // 居中、单行、默认粗体
var column = new DataGridViewTextBoxColumn {
    Name = "result", HeaderText = "处理结果", // 保持稳定列ID
    AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
};
grid.Columns.Add(column);
Ui.AutoFitColumns(grid);                    // 包括Fill列的完整表头下限
```

以下是 EnsureGridHeaders 的关键测宽逻辑摘录，不是独立的完整函数；完整函数还负责重入保护和单行表头高度。graphics 是从该网格 CreateGraphics 取得并由 using 释放的 Graphics，不按字符串字符数换算像素：

```csharp
column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
column.HeaderCell.Style.WrapMode = DataGridViewTriState.False;
var style = column.HeaderCell.InheritedStyle;
Size text = TextRenderer.MeasureText(graphics, column.HeaderText ?? "",
    style.Font ?? grid.Font, Size.Empty,
    TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
int sorting = column.SortMode == DataGridViewColumnSortMode.NotSortable
    ? 0 : Ui.Px(grid, 18);
column.MinimumWidth = Math.Max(Ui.Px(grid, 50),
    text.Width + style.Padding.Horizontal + Ui.Px(grid, 12) + sorting);
```

**调用顺序与边界。** 创建列→加入既有布局→调用 Ui.AutoFitColumns；表头字体或文案调整复用 DialogFormatting，最终阶段重新测量。不要单独开启表头换行或用固定宽度压缩标题。文件提取的 ResultGrid/IGridColumnFit 保持局部适配回调，尺寸目录保留专属树形绘制及 FitColumns，不以普通适配覆盖它们。

生产网格按实例盘点，不把测试夹具计入：Excel 为追踪引用、追踪从属、错误来源、按列筛选、文件提取名称、OCR 前缀、OCR 字段核对、表头合并、合并抵消往来、合并抵消收入成本及交互弹窗尺寸共 11 张；Word 为查看链接及交互弹窗尺寸共 2 张。所有入口及专属网格均遵守完整表头下限。

**验证方法。** 两项目[GridHeadersProbe.cs](../C-01-ExcelIterate/src/testharness/GridHeadersProbe.cs)的 GridHeaderChecks 构造真实生产控件，比较 WinForms 原生表头所需宽高；工具先按 project.json 编译，再用 `Feature099Probe.exe --grid-headers` 单独运行。检查完整单行表头、居中、默认粗体及显式关闭粗体，覆盖旧自定配置、较大字号、排序图标、窄窗横滚、手动缩窄或调宽、反复缩放和 100/125/150/200% 字体模拟条件。后台控件检查不连接 WPS，字体模拟及离线渲染不证明真实 WPS 外观、鼠标拖动或跨屏 DPI；未执行的实机交互单独列待验收。

### (2). 程序所新生成工作表的格式规范

#### a. 格式规范说明

本节描述合并抵消等审计报告的格式；普通结果表沿用 WorkbookUi.StyleHeader 的浅底深字，不套用本节的填充色与字体。所有新增正式结果表必须遵循对应实现，不得引用不存在的样张文件。

##### i. 首行合计（SUBTOTAL）

合并抵消报告首行为合计行，每个金额列写 `=SUBTOTAL(9, <数据首行>:<工作表末行>)`（如 `=SUBTOTAL(9,G4:G1048576)`）。该行字体 Times New Roman 11、套会计数字格式（见 iii）、垂直居中、**不设边框**；合计行不计入表头行高 26 磅下限，按默认行高。

##### ii. 表头格式规范

合并抵消报告表头行统一设置：居中、加粗、不换行；填充色 `#4874CB`（主题色 accent1）、白字；字体 Times New Roman 11；行高按最终字体适配、保留 26 磅下限；空结果仍须完整显示表头。

![程序生成工作表的表头与全数据列宽适配构造示意](../C-01-ExcelIterate/AB-Graphic/格式规范-生成工作表-构造示意.png)

上图为构造示意，非真实 WPS 截图：表头单行、居中加粗，列宽依据表头和全部数据适配，最后一行的最长文本也参与测量。[可编辑SVG](../C-01-ExcelIterate/AB-Graphic/output-format-maintenance.svg)。

##### iii. 数据行格式规范

合并抵消报告数据行字体 Times New Roman 11、不加粗；金额列套会计数字格式 `_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ `（含千位分隔、负数带减号、零显短横）；垂直居中；备注、调整事项、会计科目等指定列水平居中。

##### iv. 边框规范

合并抵消报告表头与数据区统一边框：**最左列 left=medium（xlContinuous, xlMedium）、其余内部边框=slantDashDot（xlSlantDashDot=13）**——即表格左缘为粗实线、内部网格为斜点划线。**首行合计行（SUBTOTAL）不设边框**。表头顶边为 medium。

##### v. 列宽与正文显示规范

在数据、字体、数字格式及超链接设置完成后，对包含表头和全部写入数据的输出范围执行列 AutoFit，不只测表头，不再用固定列宽覆盖适配结果。最长值在最后一行或隐藏行也应参与；OCR 起始列不必为 A 列。为测量隐藏数据，可临时显露输出范围内的隐藏行，必须在 finally 逐一恢复原 Hidden，不改变 AutoFilter 及其条件；恢复失败须报告，不能记为完成。

极长正文达到宿主列宽上限仍显示不全时，只对本程序输出区域的正文换行并适配行高，不改数据内容或套用网格的 30 字符上限；表头保持单行。业务状态回写结束后适配本批结果，不持续监听用户后续编辑而覆盖手动列宽。

普通结果表沿用 WorkbookUi.StyleHeader；合并抵消报告使用本节规则。本次普通结果范围为 Excel 的“OCR提取结果”“文件批量命名”“外部链接”“高筛结果”和工作表目录，目录保留已有全表样式。已有合并目标表、业务临时条件表及 Word 文档表格不套用报告格式。测量范围限定为本程序输出区域；列宽和行高属于整列或整行属性，会影响同列或同行的其他单元格显示，不能据此改变其数据。

#### b. 代码复现方法

共享的是维护规范，具体工作表操作由 Excel 实现。源码链接在完整项目目录中使用；下列片段在同一加载项源码程序集内复用，依赖现有内部类型。OutputFormatSnapshot 和 UpdateReport 为私有实现，阅读其逻辑并从原业务入口接入，不作为外部直接调用接口。

| 用途           | 可复用的现有实现                                                                                                                                                                                                 | 调用位置与边界                                                                                                              |
| ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| 完整报告范围及隐藏行测量 | [WorkbookFeatures.cs](../C-01-ExcelIterate/src/addin/WorkbookFeatures.cs)：WorkbookUi.StyleHeader、FitOutput、ReadRowHeight、OutputFormatSnapshot                                                         | StyleHeader 规范表头；FitOutput 在最终数据、字体、数字格式后测完整输出范围、适配表头行高；超限正文换行。失败时按快照还原列宽、换行、行高，finally 恢复原 Hidden；隐藏行的 0 或 null 行高先安全读取潜在高度 |
| OCR写入与失败还原   | [Connect.FileNames.cs](../C-01-ExcelIterate/src/addin/Connect.FileNames.cs)：FileNameWorksheet.Write                                                                                                   | 写入前保留值、公式、格式、尺寸、隐藏状态，已有筛选保留，新增筛选后再 FitOutput；失败逐项还原并报告未恢复项                                                            |
| 状态回写与保存      | [LinkAssistant.cs](../C-01-ExcelIterate/src/addin/LinkAssistant.cs)：Scan、UpdateReport；[FileToolForms.cs](../C-01-ExcelIterate/src/addin/FileToolForms.cs)：RenameWorkbookSession.Create、WriteResult | 内容及状态回写完成→FitOutput→保存；断链清单保存后才登记 checkpoint.LastHash                                                                 |
| 目录及高筛结果      | [WorkbookFeatures.cs](../C-01-ExcelIterate/src/addin/WorkbookFeatures.cs)：DirectoryNavigation.Generate、QuickAdvancedFilter.Run                                                                        | 目录在 Hyperlinks.Add 和全表字体、对齐设置后适配；高筛在原列数字格式恢复和最终字体设置后适配，业务临时条件表不套报告格式                                                   |

**最小调用示例。** 以下完整接线片段演示新报告的格式调用，不代替现有写入、保存、恢复事务。rows 包含表头，startColumn 可不是 A 列，最后一个地址须包含最后一条数据；sheet 由业务入口提供：

```csharp
object[,] values = { { "名称", "说明" }, { "样例文件", "完整结果文字" } };
int startColumn = 2;                        // 从B列写入
int rows = values.GetLength(0), columns = values.GetLength(1);
using (var bag = new ComBag()) {
    string address = ComFormulaHost.Address(1, startColumn) + ":" +
        ComFormulaHost.Address(rows, startColumn + columns - 1);
    dynamic range = bag.Own(((dynamic)sheet).Range[address]);
    range.NumberFormat = "@";
    range.Value2 = values;
    WorkbookUi.FitOutput(sheet, startColumn, columns, rows);
    // FitOutput内设置表头、完整范围Columns.AutoFit及表头Rows.AutoFit；
    // 超限时只对正文WrapText及Rows.AutoFit，并恢复原Hidden。
}
// 如该业务需要保存，应在适配成功后调用原保存入口；不要额外自动保存OCR结果。
```

**调用顺序与边界。** 内容及状态回写→完成字体、数字格式、超链接及筛选设置→FitOutput→按原业务入口保存→登记提交摘要。适配后不回写固定列宽；OCR 结果沿用不自动保存的行为。外部链接清单和文件命名清单的格式适配在相关保存之前完成，含断链清单须在适配及保存后登记提交摘要，不能使恢复记录与实际文件字节不一致。

列宽、正文换行及受影响行高纳入原操作失败恢复；测量或格式写入失败须还原本次已尝试改动并报告未恢复项。格式快照和隐藏行恢复复用现有实现，保留原业务的内容、公式、筛选和恢复事务，不能用示例替换业务入口。

**验证方法。** 工具先按 Excel 的 project.json 编译；[Feature099Probe.cs](../C-01-ExcelIterate/src/testharness/Feature099Probe.cs)通过 `Feature099Probe.exe --output-format` 覆盖 OCR 及格式快照恢复，[WorkbookToolsProbe.cs](../C-01-ExcelIterate/src/testharness/WorkbookToolsProbe.cs)通过 `WorkbookToolsProbe.exe` 覆盖链接、命名、目录及高筛输出接线。检查五类输出、空结果、最后一行最长值、非 A 列起始、隐藏数据、超长正文、状态回写及适配中途失败后的恢复；测试模型应保存列宽、换行、行高与测量范围，不能只统计 AutoFit 调用次数。

这些后台 COM 模型检查不连接 WPS；高筛模型仅预置原生筛选结果，验证格式步骤，不验证宿主筛选算法。后台通过不证明真实 WPS 列宽、隐藏行测量或可见换行效果；尚未执行的实机交互单独列待验收。

# 二. XyCabin 网站维护纪律

## 1. 文档、版本与经验归属

**（人工）** 网站只维护前端经验（打开、指南、下载、音乐、帮助），不复制宿主公式或字段事务；改指南时核两个宿主是否适用；无新增经验如实说明，不新增维护副本。

文档结构、编号、日期、概览、交付证据、链接与图片规则直接适用第一章 R-01–R-14，不另立一套。

**R-22** PRODUCTS 版本新到旧，摘要与 B1 对应章节逐字一致，每版 trim 后不超过 200 字，每个功能按 1.、2. 连续编号。

## 2. 验证、编码与缓存

日常 rebuild/validate 显式 `--background`，只跑已核对不显示或激活窗口的路径；两加载项 DialogSizesProbe `--core` 仍必需，完整预览与拖动单列。`--skip-ui` 只跳截图不保证后台，透明窗不等后台，不抢焦点、不切用户页面、不自动打开截图。

| 工具 | 验证内容与边界 |
| --- | --- |
| devlog/markdown/layout 门禁 | 五份文档的章节范围、概览、交付证据、版本表、规则注册表，三输出及维护唯一；源码携带只允许审计历史 build 证据缺失，正式 strict-evidence 要求链接齐全 |
| site_probe | 配置非法输入、HTML 结构、资源、脚本语法、版本源、快照关系 |
| audio_probe/download_probe | 隔离 DOM 与 AudioContext、状态、字节、交互，不真实发声或保存文件 |
| workflow_probe/integration_probe | 真实 CMD 入口参数、提示、退出码、事务失败保护、无旧 build 及中文、空格、中点、其他 cwd 路径 |
| file_inventory/cleanup 探针 | 必要输入含父 verify_shared.py、候选引用、指纹与审阅清单，不能伪造隔离成功记录 |

**R-15** 三批处理 UTF-8 无 BOM、CRLF、chcp65001，Python 输出 UTF-8；Windows PowerShell 5 入口避免中文无 BOM 误读。网站末尾中文由 Python 的 ASCII 转义命令输出，子任务后重设 65001；编译诊断按实际编码，不把 GBK 推广至所有入口。

增量 `--mode incremental`、默认 full；build_cache 按源码、配置、依赖、规范快照、环境及产物证据摘要判断，缺失或变化重执行，首次无缓存完整。完整不复用；签名、载荷与网页包字节见第一章 R-19、R-20 与本章 R-21。锁内 stage 收据仅成功事务后提交，last-run-summary 及 validation-timings 记录耗时，复用保留原检查时间，不声称本轮重测。**（人工）** 固定图片样例不读取桌面原图。

## 3. 统一交付与失败恢复

固定无 EXE 快照→以 `--site-html`、`--installer-output` 重建两暂存包→业务、文档、原证书双签名、五载荷验证→attach_installers 附本轮 EXE→verify_installers 对实际包逐字节及版本、名称、长度、SHA-256 核对→delivery.publish 事务替换两 EXE、HTML 和 current-delivery.json→正式资源复核。

**R-21** 两 DLL 内嵌 WpsShared.Agent.Home.html 必须逐字节相同，移除独立页 installer-assets 后比较网页正文；两个 DLL 含不同宿主业务，不要求 DLL 字节相同。编译只接受空包数据区，防安装包包含自身；网站产品按独立身份映射，不将 Excel 宿主等同 Excel-Iterate 产品。

**（人工）** 缺源码、资产或原私钥停止，不回退旧 HTML 或换发布者；发布失败恢复上轮正式文件，回滚未完保留恢复文件并报路径。最终审计集中本日志，宿主详细结果链接对应日志；内部 current-delivery.json 是机器交付记录，不是另建维护文档。用户保存并退出 WPS 后分别安装才更新加载项内网页；缓存与浏览器接线归加载项共享经验。

## 4. 冗余文件清理

网站与两个加载项共用第一章第 4 节的清理纪律（R-23 与人工条款），默认范围为三个项目；网站入口为 `src/tools/cleanup.py --dry-run` 与 `--apply`，清单、fixture-state、绝对路径边界与保护集合同第一章执行。
