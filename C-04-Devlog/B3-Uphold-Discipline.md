**共享维护纪律与网站维护纪律**

# 一. Word与Excel共享维护纪律

## 1. 文档、编号与经验更新

仓库根 AGENTS.md 只放指向 B0-Entry.md 的指针；B0-Entry.md 只放阅读顺序与按角色分流；开发日志保存可执行方法、边界、入口和证据；使用说明保存用户步骤/结果/限制。源码文档为知识来源，不依赖旧聊天或隐藏记忆。不额外建立维护副本。

标题体系如下，最多五级，不为凑层级拆短句；同级连续编号、新父级重置子级，编号后一个半角空格。新增经验归现有主题，必要时用第五级细分，代码围栏内不参与编号：

```markdown
# 一. 标题
## 1. 标题
### (1). 标题
#### a. 标题
##### i. 标题
##### ii. 标题
```

维护顺序：对照最终源码→提炼做法/边界/入口/证据→判断共享→更新对应主题、指南/地图/目录/顶部日期→文档/锚点/共享门禁→固定正文重建→回填审计。共同经验及公共源码成对同步、逐字/字节一致；宿主算法留第五章，姊妹缺失写待同步，不声称成对通过。无新增可迁移经验如实说明，不编造根因。

版本唯一来源INSTALLER_VERSION，四段0～65534，仅用户指定才修改；B1中每个加载项章节的版本表升序，每版一行最多200字，新第三段主版本将前一主版本子版本并入末段0。安装器/加载项/清单版本同源，配置导入不生成文件/字节码，值不变不回写；网站PRODUCTS从新到旧但摘要与B1对应章节一致。

## 2. 重建、签名与交付

现有单项目入口只重建自己，无--background参数：`python rebuild.py --mode full` 后 `python validate.py --background`。两产品文档/网页共同交付直接在Website运行 `python rebuild.py --background --mode full`（或其AA-Program批处理选2），再 `python validate.py --background`；增量选1/--mode incremental，无mode默认full。携带完整Website；两加载项均在时带父verify_shared.py。构建不安装、不结束WPS。

链路为各项目配置器（C1-Exceliterate_Configure.py、C2-Wordupdater_Configure.py）→build.py→resign.py→build_installer.py；每次编译后必须重签，csc的新MVID会改变DLL摘要。固定证书缺失/无私钥/过期或错指纹即停，不切换发布者。安装器内嵌主DLL、Utilities、Office.Core、应用清单、部署清单共五载荷。

清单保留FullTrust/customHostSpecified/GAC与vstav3/vstov4结构；未强签DLL依赖不带token。应用清单签后读顶层identity重建vsto引用，计算的是manifest文件本身的精确大小与SHA-256 Base64摘要；UTF-8 BOM/CRLF。SignCmi采用SHA-256，两清单分别有StrongName/Authenticode签名；ValidateManifest查结构，ValidateXmlSig查整清单及license绑定。信任缺失与签名损坏分别报告，缺校验器/载荷不符阻止交付。

统一入口固定空installer-assets快照→两包暂存与后台验证→独立HTML附本轮包→核对正文/资源/包字节→事务替换两EXE、HTML和current-delivery.json。失败保留上轮正式产物，回滚未完保留恢复文件；只生成配置预览不表示完成联合发布。两个DLL的HTML资源相同，不要求DLL相同；发布后再次单独打包须重跑联合入口以恢复下载一致性。

安装需完整VSTO Runtime和官方VSTOInstaller：RuntimeLocator查HKLM两视图及Common Files，缺少先停止，不释放新载荷/写信任。用户保存文件并完全退出WPS后运行安装器，在官方窗口授权；注销本产品旧订阅成功才替换/重装，真实退出0才补Manifest、LoadBehavior=3、AddinsWL。取消/非零退出如实报告，不静默安装。卸载只清自身订阅/挂载/白名单/信任/载荷，保留LastInstallDir。

排障按“运行库→SetupLogs→五载荷/挂载/白名单/订阅信任→签名摘要→真实宿主”。Manifest是绝对file URL+|vstolocal，codebase相对；移动目录须重装。同版本不代表同字节，查安装路径与实际加载DLL摘要。COMAddIns.Connect可作连接证据，Object=null/模块扫描/文件释放不能单独定启用。SetupDoctor在build，只读诊断不随正式输出交付。

## 3. 验证、缓存与证据边界

### (1). 用COM启动、附加和验证WPS

在STA执行：表格用KET.Application并创建Workbooks.Add，文字用KWPS.Application并创建Documents.Add。只读诊断用Marshal.GetActiveObject附加现有实例，失败返回SKIP/INCOMPLETE，不启动应用；已授权需要启动时使用Type.GetTypeFromProgID加Activator.CreateInstance（PowerShell等价为New-Object -ComObject），再设置Visible并创建测试对象，不能只打开WPS首页。附加模式只限定不启动宿主，不自动使业务只读；只读诊断不创建或修改文档。以下为现有.NET Framework 4.8探针中ProbeSession的开窗接线示例，外围Main标记[STAThread]，捕获异常并用session.Finish记录结果；它不包含具体业务验收断言。PowerShell直接调用Marshal.GetActiveObject应使用Windows PowerShell 5.1，PowerShell 7须采用适用的ROT/原生接口实现。

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

注册映射和进程核对另按上节执行；示例中的文件产品检查不能代替完整身份记录。ProgID不存在、指向不符、未发现现有实例或身份不足时明确失败/未完成；以WPS为目标的验收不静默改用Excel.Application/Word.Application。现有业务探针的兼容回退链不是严格WPS启动契约，运行前须核对实际连接；附加失败可检查ROT文件moniker，但只能选择明确的测试文件，不能任取第一个文档。GetActiveObject也不保证拿到用户眼前那个窗口，目标用实际文档/工作簿身份确认。

CreateInstance可能复用已有Application，不能据此认领或Quit宿主。仅关闭自己创建并持有的测试文档/簿；所有自取COM属性/集合/范围逐次登记并ReleaseComObject，避免链式访问产生遗漏，不FinalRelease共享RCW。读取原值成功后才临时设置Visible、DisplayAlerts、ScreenUpdating、EnableEvents等，并在finally恢复；无法读取则不改。原文档Saved只观察，不设true掩盖修改；按需恢复原活动文档/簿、选区和剪贴板。失败返回非零并保留样本及日志，不按全局新增PID退出或强杀进程。实现入口为两项目src/tools/LiveProbeSafety.cs及对应实机探针。

COM可验证真实宿主中的业务读写、字段/公式、保存与还原；不能代替按钮点击、键入焦点、弹窗外观、鼠标拖动及实际跨屏DPI验收。COMAddIns.Connect仅说明连接状态；本项目未提供COMAddIns.Object自动化桥，Object=null须记INCOMPLETE，不能用新建Connect实例冒充已加载Ribbon回调。COM/API检查不需要computer-use；可见交互尚未执行时保留人工待验收状态。启动与附加API区别参见[New-Object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/new-object)、[GetActiveObject](https://learn.microsoft.com/en-us/dotnet/fundamentals/runtime-libraries/system-runtime-interopservices-marshal-getactiveobject)。

### (2). 弹窗维护规范

交互弹窗设错窗口样式会把宿主锁死：弹窗打开后无法在单元格键入、点不开有子选项的功能区菜单，而宿主窗口 enabled 仍为 True、宿主线程仍活着（功能区主按钮可反复点击），极易误判为宿主被禁用或线程阻塞。真根因是工具窗样式：设 `ShowInTaskbar = false`、`MinimizeBox = false` 会加上 WS_EX_TOOLWINDOW，无主非模态 tool 窗在 WPS 宿主线程上抢焦点/失激活——前台被弹窗把持，键盘输入进不了文档、功能区下拉点不开，宿主 enabled 与线程状态都不受影响，表现为“假死”。曾两度误判：先删无主窗的 TopMost，再把 Load 内业务 BeginInvoke 异步化，均未生效；按打开前后 enabled 均 True 排除宿主被禁用、功能区主按钮可反复点排除线程阻塞、前台窗口=弹窗指向激活/焦点问题三重排除后，才锁定与链接助手等正常弹窗的样式差异。Word 与 Excel 两侧的交互弹窗都按本节构建。

正确构建交互弹窗：不设 `ShowInTaskbar = false`、`MinimizeBox = false`（即不引入 WS_EX_TOOLWINDOW），不重写 ShowWithoutActivation，不向无主窗设 TopMost，与链接助手等既有正常弹窗保持同套默认样式；非模态显示统一 `IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));`，WPS 回调线程上 owner 校验常不成立、取到 0 时走普通 Show 即可，样式正确即不锁死，不为拿 owner 改样式；Load 只做界面装配，重业务移出 Show 栈（BeginInvoke 到消息泵 idle），内置打开前后 enabled/前台/owner 诊断日志，复发时按三重排除定位，不再次误判；新弹窗在 DialogSizes、DialogPreviewFactory、dialog_catalog_probe、GridHeadersProbe 登记尺寸与目录，探针补源码守卫（禁 TopMost=true 与工具窗样式）和预览构造断言（ShowInTaskbar/MinimizeBox 保持默认 true），把根因钉进门禁；验收必须含真实宿主交互：开窗后键入单元格、展开功能区子菜单、Esc 关闭，后台探针通过不等于焦点与菜单验收。

## 4. 冗余文件清理

只有用户实际要求“启动冗余清理”才触发盘点/删除；讨论规则、改文档或构建不触发历史清理。默认范围三个项目，另指定从其指令；项目外原图、用户配置/缓存不自动纳入。

冗余候选须同时满足：生命周期结束/有有效再生办法；不在现行依赖闭包（含动态加载、资源、文档/证据引用）；独有经验/恢复信息已归并；隔离排除候选后从保留输入无旧build重建和适用验证通过。旧日期、old/backup文件名、build目录或搜不到引用都不足以证明可删。

分五类盘点：必要输入、当前产物、有效证据、可删历史、未确认。按实际字节从大到小核对用途/摘要，公共副本职责不同不能因同SHA删除；移动项目内不算释放。私钥/必要依赖/未完成恢复资料始终保护，.cer和公共CA .pem不按扩展名误作私钥。

执行顺序：记录范围/版本/用途/数量/字节/摘要→归并经验修引用→隔离排除候选完成重建/签名/载荷/资源验证→Website `src/tools/cleanup.py --dry-run`生成schema 3五类清单→逐项审阅→`--apply`读同一清单复查→复核正式产物/链接并报告实际删除/失败/保留及字节差。清单后的输入/目标变化须重新盘点。

fixture-state必须真实结束；cleanup-review绑定源码指纹、逐文件摘要、用途/经验结论、隔离成功证据及原日志摘要，禁止伪造状态。删除前核绝对路径在指定项目内，不删项目根、不跟随联接、不碰项目外文件；不完整恢复保留。已授权确认项按范围执行。探针成功仅释放自建临时文件，失败保留诊断与恢复资料。

## 5. 格式维护规范

### (1). 弹窗中数据网格的格式规范

#### a. 格式规范说明

##### i. 表头格式规范

所有新增及现有生产弹窗的数据网格表头完整单行显示、居中、默认加粗。用户在“自定格式”中明确关闭粗体时尊重其选择；字号、字体、颜色和表头文字仍可自定，不能因此开启换行或截断表头。单行表头高度随实际字体、字号和DPI重新测量，避免上下裁剪；正文样式、网格稳定身份和已有用户配置保持，不通过重置配置取得默认外观。

![弹窗数据网格表头格式的离线控件渲染](../C-01-ExcelIterate/AB-Graphic/格式规范-弹窗网格表头-离线渲染.png)

上图来自“设置OCR前缀”的离线生产控件渲染，展示完整、单行、居中加粗的表头；非真实WPS截图，不作为宿主交互验收证据。

##### ii. 列宽与正文显示规范

列宽下限按当前实际表头字体、文字、内边距和排序图标测量，包含Fill列、复选列、下拉列与按钮列。手动缩窄不得低于完整表头宽度；宽度总和超过窗口时通过横向滚动查看，不压缩或换行表头。

文件提取名称正文自动适配上限仍为当前字体30个“0”加内边距，通常下限50逻辑像素，完整表头需要更宽时优先。该上限不裁剪数据、不限制手动调宽。新扫描重启适配周期；本周期后续重试、核对、规则调整与窗口缩放尊重已手调列宽。

#### b. 代码复现方法

源码链接在完整项目目录中使用；加载项离线全文保留下列片段，不携带业务源码。完整实现以链接中的函数为准；在同一加载项源码程序集内复用现有内部类型，修改公共网格代码后同步两个项目，不另维护一份格式实现。

| 用途            | 可复用的现有实现                                                                                                                                                                                     | 调用位置与边界                                                                                                                                                                           |
| ------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 网格默认样式及完整表头测量 | [UiKit.cs](../C-01-ExcelIterate/src/addin/UiKit.cs)：Ui.StyleGrid、Ui.AutoFitColumns、Ui.EnsureGridHeaders                                                                                                        | 创建列后适配；StyleGrid在HandleCreated复核DPI；EnsureGridHeaders按逐列表头字体／Padding／排序图标计算下限与高度，不强制覆盖显式的非粗体字体                                                                                    |
| 自定格式应用后的收敛    | [DialogFormatting.cs](../C-01-ExcelIterate/src/addin/DialogFormatting.cs)：DialogFormatting.Apply                                                                                                               | finally阶段完成布局后重新测宽；普通网格调用Ui.AutoFitColumns，尺寸配置调用专属FitColumns，不再开启表头换行                                                                                                            |
| 保留局部列宽策略      | [FileNamesLayout.cs](../C-01-ExcelIterate/src/addin/FileNamesLayout.cs)：FileNamesForm.FitResultColumns；[DialogSizeEditor.cs](../C-01-ExcelIterate/src/addin/DialogSizeEditor.cs)：DialogSizesForm.FitColumns | 文件提取先在fittingColumns保护内调用EnsureGridHeaders，再处理正文30个0上限和手动列宽；尺寸目录保留树缩进及自绘按钮测量。文件提取通过ResultGrid／IGridColumnFit接线，局部回调不反调Ui.AutoFitColumns；尺寸目录由DialogFormatting.Apply直接调用FitColumns |

**最小调用示例。** 在已经登记并使用DialogSizing.Prepare的业务窗中，先创建全部列，将grid加入业务窗的既有布局容器，再调用共享入口：

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

以下是EnsureGridHeaders的关键测宽逻辑摘录，不是独立的完整函数；完整函数还负责重入保护和单行表头高度。graphics是从该网格CreateGraphics取得并由using释放的Graphics，不按字符串字符数换算像素：

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

**调用顺序与边界。** 创建列→加入既有布局→调用Ui.AutoFitColumns；表头字体或文案调整复用DialogFormatting，最终阶段重新测量。不要单独开启表头换行或用固定宽度压缩标题。文件提取的ResultGrid／IGridColumnFit保持局部适配回调，尺寸目录保留专属树形绘制及FitColumns，不以普通适配覆盖它们。

生产网格按实例盘点，不把测试夹具计入：Excel为追踪结果、按列筛选、文件提取名称、OCR前缀、OCR字段核对、合并的来源／映射／样本及交互弹窗尺寸共9张；Word为查看链接的来源／明细及交互弹窗尺寸共3张。所有入口及专属网格均遵守完整表头下限。

**验证方法。** 两项目[GridHeadersProbe.cs](../C-01-ExcelIterate/src/testharness/GridHeadersProbe.cs)的GridHeaderChecks构造真实生产控件，比较WinForms原生表头所需宽高；工具先按project.json编译，再用`Feature099Probe.exe --grid-headers`单独运行。检查完整单行表头、居中、默认粗体及显式关闭粗体，覆盖旧自定配置、较大字号、排序图标、窄窗横滚、手动缩窄／调宽、反复缩放和100／125／150／200%字体模拟条件。后台控件检查不连接WPS，字体模拟及离线渲染不证明真实WPS外观、鼠标拖动或跨屏DPI；未执行的实机交互单独列待验收。

### (2). 程序所新生成工作表的格式规范

#### a. 格式规范说明

本节格式规范以 `AB-Graphic/格式维护.xlsx`（桌面样张 `工作簿1.xlsx` 的「格式维护」工作表同源）为权威样张，所有新增正式结果表必须严格遵循。

##### i. 首行合计（SUBTOTAL）

正式结果表首行为合计行，每个金额列写 `=SUBTOTAL(9, <数据首行>:<工作表末行>)`（如 `=SUBTOTAL(9,G4:G1048576)`）。该行字体 Times New Roman 11、套会计数字格式（见 iii）、垂直居中、**不设边框**（样张 R1 border 为空）；合计行不计入表头行高 26 磅下限，按默认行高。

##### ii. 表头格式规范

表头行统一设置：居中、加粗、不换行；填充色 `#4874CB`（主题色 accent1）、白字；字体 Times New Roman 11；行高按最终字体适配、保留 26 磅下限；空结果仍须完整显示表头。

![程序生成工作表的表头与全数据列宽适配构造示意](../C-01-ExcelIterate/AB-Graphic/格式规范-生成工作表-构造示意.png)

上图为构造示意，非真实WPS截图：表头单行、居中加粗，列宽依据表头和全部数据适配，最后一行的最长文本也参与测量。[可编辑SVG](../C-01-ExcelIterate/AB-Graphic/output-format-maintenance.svg)。

##### iii. 数据行格式规范

数据行字体 Times New Roman 11、不加粗；金额列套会计数字格式 `_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ `（含千位分隔、负数带减号、零显短横）；垂直居中（valign=center）；备注/调整事项/会计科目等指定列水平居中。

##### iv. 边框规范

表头与数据区统一边框：**最左列 left=medium（xlContinuous, xlMedium）、其余内部边框=slantDashDot（xlSlantDashDot=13）**——即表格左缘为粗实线、内部网格为斜点划线。**首行合计行（SUBTOTAL）不设边框**。表头顶边为 medium。

##### v. 列宽与正文显示规范

在数据、字体、数字格式及超链接设置完成后，对包含表头和全部写入数据的输出范围执行列AutoFit，不只测表头，不再用固定列宽覆盖适配结果。最长值在最后一行或隐藏行也应参与；OCR起始列不必为A列。为测量隐藏数据，可临时显露输出范围内的隐藏行，必须在finally逐一恢复原Hidden，不改变AutoFilter及其条件；恢复失败须报告，不能记为完成。

极长正文达到宿主列宽上限仍显示不全时，只对本程序输出区域的正文换行并适配行高，不改数据内容或套用网格的30字符上限；表头保持单行。业务状态回写结束后适配本批结果，不持续监听用户后续编辑而覆盖手动列宽。

所有新增正式结果表沿用此规则；本次范围为Excel的"OCR提取结果""文件批量命名""外部链接""高筛结果"和工作表目录，目录保留已有全表样式。已有合并目标表、业务临时条件表及Word文档表格不套用报告格式。测量范围限定为本程序输出区域；列宽和行高属于整列／整行属性，会影响同列／同行的其他单元格显示，不能据此改变其数据。

#### b. 代码复现方法

共享的是维护规范，具体工作表操作由Excel实现。源码链接在完整项目目录中使用；下列片段在同一加载项源码程序集内复用，依赖现有内部类型。OutputFormatSnapshot和UpdateReport为私有实现，阅读其逻辑并从原业务入口接入，不作为外部直接调用接口。

| 用途           | 可复用的现有实现                                                                                                                                                                                                 | 调用位置与边界                                                                                                              |
| ------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------- |
| 完整报告范围及隐藏行测量 | [WorkbookFeatures.cs](../C-01-ExcelIterate/src/addin/WorkbookFeatures.cs)：WorkbookUi.StyleHeader、FitOutput、ReadRowHeight、OutputFormatSnapshot                                                         | StyleHeader规范表头；FitOutput在最终数据／字体／数字格式后测完整输出范围、适配表头行高；超限正文换行。失败时按快照还原列宽／换行／行高，finally恢复原Hidden；隐藏行的0／null行高先安全读取潜在高度 |
| OCR写入与失败还原   | [Connect.FileNames.cs](../C-01-ExcelIterate/src/addin/Connect.FileNames.cs)：FileNameWorksheet.Write                                                                                                   | 写入前保留值／公式／格式／尺寸／隐藏状态，已有筛选保留，新增筛选后再FitOutput；失败逐项还原并报告未恢复项                                                            |
| 状态回写与保存      | [LinkAssistant.cs](../C-01-ExcelIterate/src/addin/LinkAssistant.cs)：Scan、UpdateReport；[FileToolForms.cs](../C-01-ExcelIterate/src/addin/FileToolForms.cs)：RenameWorkbookSession.Create、WriteResult | 内容及状态回写完成→FitOutput→保存；断链清单保存后才登记checkpoint.LastHash                                                                 |
| 目录及高筛结果      | [WorkbookFeatures.cs](../C-01-ExcelIterate/src/addin/WorkbookFeatures.cs)：DirectoryNavigation.Generate、QuickAdvancedFilter.Run                                                                        | 目录在Hyperlinks.Add和全表字体／对齐设置后适配；高筛在原列数字格式恢复和最终字体设置后适配，业务临时条件表不套报告格式                                                   |

**最小调用示例。** 以下完整接线片段演示新报告的格式调用，不代替现有写入／保存／恢复事务。rows包含表头，startColumn可不是A列，最后一个地址须包含最后一条数据；sheet由业务入口提供：

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

**调用顺序与边界。** 内容及状态回写→完成字体、数字格式、超链接及筛选设置→FitOutput→按原业务入口保存→登记提交摘要。适配后不回写固定列宽；OCR结果沿用不自动保存的行为。外部链接清单和文件命名清单的格式适配在相关保存之前完成，含断链清单须在适配及保存后登记提交摘要，不能使恢复记录与实际文件字节不一致。

列宽、正文换行及受影响行高纳入原操作失败恢复；测量或格式写入失败须还原本次已尝试改动并报告未恢复项。格式快照和隐藏行恢复复用现有实现，保留原业务的内容、公式、筛选和恢复事务，不能用示例替换业务入口。

**验证方法。** 工具先按Excel的project.json编译；[Feature099Probe.cs](../C-01-ExcelIterate/src/testharness/Feature099Probe.cs)通过`Feature099Probe.exe --output-format`覆盖OCR及格式快照恢复，[WorkbookToolsProbe.cs](../C-01-ExcelIterate/src/testharness/WorkbookToolsProbe.cs)通过`WorkbookToolsProbe.exe`覆盖链接、命名、目录及高筛输出接线。检查五类输出、空结果、最后一行最长值、非A列起始、隐藏数据、超长正文、状态回写及适配中途失败后的恢复；测试模型应保存列宽、换行、行高与测量范围，不能只统计AutoFit调用次数。

这些后台COM模型检查不连接WPS；高筛模型仅预置原生筛选结果，验证格式步骤，不验证宿主筛选算法。后台通过不证明真实WPS列宽、隐藏行测量或可见换行效果；尚未执行的实机交互单独列待验收。

# 二. XyCabin 网站维护纪律

## 1. 文档、版本与经验归属

B0-Entry.md 是阅读入口，AGENTS.md 只放指针；使用说明是打开、指南、下载、音乐和帮助；开发日志保存实现方法、边界、入口和证据。网站只维护前端经验，不复制宿主公式/字段事务。新经验归对应主题，保留独有失败/恢复结论，改指南时核两宿主适用性；无新增经验如实说明。不新增维护MD副本。

三份文档最多五级，编号后一个半角空格：`# 一. 标题`→`## 1. 标题`→`### (1). 标题`→`#### a. 标题`→`##### i. 标题`；同级连续、新父重置，五级i./ii./iii.小写罗马，不跳级、不强凑五层，代码围栏排除。顶部仅一个更新日期，目录/锚点同步；概览与delivery-audit标记唯一有序，目录在概览外。门禁同时查三项目九文档。

版本从两加载项INSTALLER_VERSION读取，用户未指定不递增。PRODUCTS版本新到旧、摘要与对应日志逐字一致，每版最多150字，每个功能按1.、2.编号。公共方法在两加载项第三、四章逐字同步，公共源码字节同步；网站指南以对应使用说明为准。

维护顺序：对照最终实现→提炼做法/边界/入口/证据→更新正文、指南与配置→文档/锚点/共享门禁→固定非审计正文→联合重建→回填真实审计。正式结果与旧证据区分，不将猜测写为已证实根因；回填审计不能改变DLL内嵌正文。

## 2. 验证、编码与缓存

日常rebuild/validate显式--background，只跑已核对不显示/激活窗口的路径；两加载项DialogSizesProbe --core仍必需，完整预览/拖动单列。--skip-ui只跳截图不保证后台，透明窗不等后台，不抢焦点/切用户页面/自动打开截图。

| 工具 | 验证内容与边界 |
| --- | --- |
| devlog/markdown/layout门禁 | 四章、三份编号/标记/锚点、三输出及维护唯一；源码携带只允许审计历史build证据缺失，正式strict-evidence要求链接齐全 |
| site_probe | 配置非法输入、HTML结构/资源/脚本语法、版本源、快照关系 |
| audio_probe/download_probe | 隔离DOM与AudioContext、状态/字节/交互，不真实发声/保存文件 |
| workflow_probe/integration_probe | 真实CMD入口参数/提示/退出码、事务失败保护、无旧build及中文/空格/中点/其他cwd路径 |
| file_inventory/cleanup探针 | 必要输入含父verify_shared.py、候选引用/指纹与审阅清单，不能伪造隔离成功记录 |

三批处理UTF-8无BOM/CRLF/chcp65001，Python输出UTF-8；WindowsPowerShell5入口避免中文无BOM误读。网站末尾中文由Python的ASCII转义命令输出，子任务后重设65001；编译诊断按实际编码，不把GBK推广至所有入口。

增量--mode incremental、默认full；build_cache按源码/配置/依赖/规范快照/环境及产物证据摘要判断，缺失或变化重执行、首次无缓存完整。完整不复用；固定私钥/清单双签名/五载荷/网页包字节每轮核。锁内stage收据仅成功事务后提交，last-run-summary及validation-timings记录耗时，复用保留原检查时间，不声称本轮重测。固定图片样例不读取桌面原图。

## 3. 统一交付与失败恢复

固定无EXE快照→以--site-html/--installer-output重建两暂存包→业务/文档/原证书双签名/五载荷验证→attach_installers附本轮EXE→verify_installers对实际包逐字节及版本/名称/长度/SHA-256核对→delivery.publish事务替换两EXE、HTML和current-delivery.json→正式资源复核。

两DLL内嵌WpsShared.Agent.Home.html必须逐字节相同，移除独立页installer-assets后比较网页正文；两个DLL含不同宿主业务，不要求DLL字节相同。编译只接受空包数据区，防安装包包含自身；网站产品按独立身份映射，不将Excel宿主等同Excel-Iterate产品。

缺源码/资产/原私钥停止，不回退旧HTML或换发布者；发布失败恢复上轮正式文件，回滚未完保留恢复文件并报路径。最终审计集中本日志，宿主详细结果链接对应日志；内部current-delivery.json是机器交付记录，不是另建维护文档。用户保存/退出WPS后分别安装才更新加载项内网页；缓存/浏览器接线归加载项共享经验。

## 4. 冗余文件清理

用户实际要求“启动冗余清理”才执行，讨论规则/文档修改/普通构建不触发。默认范围三个项目，外部原图/用户配置/缓存不自动纳入。冗余须生命周期结束且可证明不影响功能、从零重建、签名、必要验证/恢复；同时核依赖闭包、经验归并、隔离排除候选的无旧build成功证据，不能靠文件名/年龄/build目录或无搜索引用判删。

盘点必要输入、当前产物、有效证据、可删历史、未确认五类；按字节从大到小，核摘要与职责，公共副本不可因同摘要删除，项目内移动不算释放。证书私钥/必要输入/未完恢复始终保护，公开CA .pem不按后缀误作私钥。

记录范围/版本/引用/字节/摘要→归并经验修引用→隔离排除候选完整重建与适用验证→`src/tools/cleanup.py --dry-run`生成schema 3清单→审阅→`--apply`读同清单复查→复核保留输入/产物/链接并报实际删除失败及前后字节。目标或源码含父检查器变更需重盘；fixture-state真实结束，cleanup-review绑定源码/目标/独有经验/隔离证据及原日志摘要，不伪造状态。

删除前核绝对路径在明确项目内，不删根、不跟随联接、不碰项目外；旧记录缺恢复信息先保留，失败不降级删保护文件。已授权确认项按范围执行。成功探针仅释放其自建临时资料，失败留诊断恢复。规则共用于联合项目清理，加载项日志保持同等边界。
