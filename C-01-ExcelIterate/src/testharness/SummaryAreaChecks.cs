using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Excel_Iterate;

internal static class SummaryAreaChecks
{
    internal static void Run(Action<string, bool, string> check)
    {
        Action<Action, string> rejects = delegate(Action action, string name) { bool rejected = false; try { action(); } catch (InvalidOperationException) { rejected = true; } check(name, rejected, "invalid input accepted"); };
        check("年度期间解析",SummaryPeriodOrder.Parse("2027年度","2027年度",false)==new DateTime(2027,1,1),"");
        check("点分日期解析",SummaryPeriodOrder.Parse("2027.2.3","2027.2.3",false)==new DateTime(2027,2,3),"");
        check("真实日期序列解析",SummaryPeriodOrder.Parse(46388.0,"1/1/2027",false)==DateTime.FromOADate(46388),"");
        check("1904日期序列解析",SummaryPeriodOrder.Parse(40000.0,"unformatted",true)==DateTime.FromOADate(41462),"");
        var periodOrder=SummaryPeriodOrder.Descending(new[]{new DateTime(2025,1,1),new DateTime(2027,1,1),new DateTime(2026,1,1)});
        check("期间倒序保存原来源索引",periodOrder[0]==1&&periodOrder[1]==2&&periodOrder[2]==0,"");
        rejects(delegate{SummaryPeriodOrder.Parse("not a date","未知期间",false);},"未知期间阻断");
        rejects(delegate{SummaryPeriodOrder.Parse("2027.2.30","2027.2.30",false);},"无效日期阻断");
        rejects(delegate{SummaryPeriodOrder.Descending(new[]{new DateTime(2027,1,1),new DateTime(2027,1,1)});},"重复时间阻断");
        var singleCompanies=SummaryAreaEngine.LeadRows(new[]{new SummaryAreaPlan{Sheet="Lead",Row=3,Column=16,Rows=2,Columns=1}});
        check("Lead单列同时用于展示与文件名",singleCompanies.Count==2&&singleCompanies[0].CodeColumn==16&&singleCompanies[0].NameColumn==16,"");
        using(var form=new SingleSummaryForm(true)){
            var ui=(SummaryAreaView)form.Controls[0];var fakeApp=new SummaryRecalcApp();
            using(var host=new SummaryAreaHost(fakeApp,form,ui)){
                var method=typeof(SummaryAreaHost).GetMethod("Recalculate",BindingFlags.Instance|BindingFlags.NonPublic);
                var region=new SummaryRecalcRange();method.Invoke(host,new object[]{region});check("区域Calculate失败后使用工作表",region.Worksheet.Calls==1&&fakeApp.Calls==0,"");
                region.Worksheet.Fail=true;method.Invoke(host,new object[]{region});check("工作表Calculate失败后使用应用",fakeApp.Calls==1,"");
            }
        }
        using(var groupForm=new SingleSummaryForm(true)){
            var ui=(SummaryAreaView)groupForm.Controls[0];ui.AddGroup("repeat","repeat_1");ui.AddGroup("repeat");ui.AddGroup("year");
            check("恢复后添加组不会覆盖旧组",ui.ExtraGroups.Count==3&&ui.ExtraGroups[1]!="repeat_1","");
            string removed=ui.ExtraGroups[1];ui.RemoveGroup(removed);check("删除额外组更新配置清单",ui.ExtraGroups.Count==2&&!ui.ExtraGroups.Contains(removed),"");ui.ClearGroups();check("清空分组清理全部配置",ui.ExtraGroups.Count==0,"");
        }
        var friendly=typeof(SummaryAreaHost).GetMethod("Explain",BindingFlags.Static|BindingFlags.NonPublic);
        check("合并跨组提示提供处理方法",((string)friendly.Invoke(null,new object[]{"公司名称合并区域跨组：G2:I2"})).Contains("请缩小合并范围"),"");
        check("数据General使用千分位",!SummaryAreaHost.PreserveNumberFormat("General"),"");
        check("日期格式保留",SummaryAreaHost.PreserveNumberFormat("yyyy.m.d"),"");
        check("百分比格式保留",SummaryAreaHost.PreserveNumberFormat("0.00%"),"");
        check("前导零序号格式保留",SummaryAreaHost.PreserveNumberFormat("0000"),"");
        check("颜色标记不误判为日期",!SummaryAreaHost.PreserveNumberFormat("0.00;[Red]-0.00"),"");
        using(var formatForm=new SingleSummaryForm(true)){
            var ui=(SummaryAreaView)formatForm.Controls[0];check("表头默认启用且没有可见开关",ui.HeadersEnabled.Checked&&ui.HeadersEnabled.Parent==null,"");
            var controls=new Queue<Control>();controls.Enqueue(ui);bool hasRead=false;while(controls.Count>0){var current=controls.Dequeue();if(current.Text=="读取当前选区")hasRead=true;foreach(Control child in current.Controls)controls.Enqueue(child);}check("所有选区只保留原生选择按钮",!hasRead,"");
        }
        var fragments=new List<SummaryAreaPlan>{new SummaryAreaPlan{Sheet="Lead",Row=3,Column=15,Rows=1,Columns=1},new SummaryAreaPlan{Sheet="Lead",Row=5,Column=15,Rows=1,Columns=1},new SummaryAreaPlan{Sheet="Lead",Row=3,Column=17,Rows=1,Columns=1},new SummaryAreaPlan{Sheet="Lead",Row=5,Column=17,Rows=1,Columns=1}};
        check("重复变量Ctrl分段按真实行配对",SummaryAreaHost.VariableColumns(fragments,5).Count==2&&SummaryAreaHost.VariableColumns(fragments,3)[1]==17,"");
        rejects(delegate{SummaryAreaHost.VariableColumns(fragments,4);},"未选公司行不得从相邻行取变量");fragments.Add(fragments[0]);rejects(delegate{SummaryAreaHost.VariableColumns(fragments,3);},"重复变量重叠选择阻断");
        var areas = new List<SummaryAreaPlan> {
            new SummaryAreaPlan { Sheet="汇总", Row=3,Column=12,Rows=4,Columns=5 },
            new SummaryAreaPlan { Sheet="汇总", Row=20,Column=12,Rows=11,Columns=5 } };
        var targets = SummaryAreaEngine.Targets(areas, 3, 17, 2);
        check("Ctrl 第一家公司 Q3:U6", SummaryAreaEngine.Address(targets[0]) == "Q3:U6", SummaryAreaEngine.Address(targets[0]));
        check("Ctrl 竖向空隙保持 Q20:U30", SummaryAreaEngine.Address(targets[1]) == "Q20:U30", SummaryAreaEngine.Address(targets[1]));
        check("Ctrl 第二家公司 V3:Z6", SummaryAreaEngine.Address(targets[2]) == "V3:Z6", SummaryAreaEngine.Address(targets[2]));
        check("Ctrl 第二片 V20:Z30", SummaryAreaEngine.Address(targets[3]) == "V20:Z30", SummaryAreaEngine.Address(targets[3]));
        check("计划从未包含间隔行", targets.TrueForAll(a => a.Row == 3 || a.Row == 20), "gap copied");
        targets = SummaryAreaEngine.Targets(areas, 50, 20, 1);
        check("行列均平移且保留17行偏移", targets[1].Row == 67 && targets[1].Column == 20, "offset wrong");
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,12,1); }, "模板重叠拒绝");
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,16383,1); }, "目标列越界拒绝");
        rejects(delegate { SummaryAreaEngine.Targets(areas,1048570,17,1); }, "目标行越界拒绝");
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,17,0); }, "未选公司拒绝");
        areas.Add(new SummaryAreaPlan { Sheet="其他",Row=40,Column=12,Rows=1,Columns=1 });
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,17,1); }, "跨表模板拒绝"); areas.RemoveAt(2);
        areas.Add(new SummaryAreaPlan { Sheet="汇总",Row=4,Column=13,Rows=1,Columns=1 });
        rejects(delegate { SummaryAreaEngine.Targets(areas,3,17,1); }, "重叠 Area 拒绝"); areas.RemoveAt(2);
        string formula = SummaryTextReference.External(@"N:\EASDATA\项目'一\", "！01-示例（集团）", "BS!$L$17");
        check("文件名称全角字符及固定源地址保留", formula == "='N:\\EASDATA\\项目''一\\[！01-示例（集团）.xlsx]BS'!$L$17", formula);
        formula = SummaryTextReference.External(@"C:\项目", "公司甲", "ADJ'!$G$1");
        check("单引号工作表转义", formula == "='C:\\项目\\[公司甲.xlsx]ADJ'''!$G$1", formula);
        formula = SummaryTextReference.External(@"C:\项目", "公司甲", "'工作 表''一'!G1");
        check("带引号空格中文工作表", formula.EndsWith("]工作 表''一'!G1", StringComparison.Ordinal), formula);
        foreach (string invalid in new[] { "标题文字", "BS!XFE1", "BS!A1048577", "BS!A0", "BS!A1:B2", "=SUM(A1)", "BS!A1+1" })
        { string text=invalid; rejects(delegate { SummaryTextReference.External(@"C:\项目", "甲", text); }, "无效引用阻断 "+text); }
        rejects(delegate { SummaryTextReference.External(@"C:\项目", "甲/乙", "BS!A1"); }, "非法公司名拒绝而不裁剪");
        rejects(delegate { SummaryTextReference.External("relative", "甲", "BS!A1"); }, "相对文件夹拒绝");
        SummaryAreaEngine.ValidateShift("=ROUND(F76-F121,2)", 20, 5);
        check("ROUND 原生公式允许2D平移", true, "");
        rejects(delegate { SummaryAreaEngine.NativeFormula("=INDIRECT(A1)"); }, "INDIRECT 阻断");
        rejects(delegate { SummaryAreaEngine.NativeFormula("='C:\\[a.xlsx]BS'!A1"); }, "不明确已有外链阻断");
        rejects(delegate { SummaryAreaEngine.NativeFormula("=CustomName+1"); }, "未确认名称阻断");
        rejects(delegate { SummaryAreaEngine.ValidateShift("=XFD1",0,1); }, "原生相对引用越界预检阻断");
        var leadParts = new List<SummaryAreaPlan> {
            new SummaryAreaPlan { Sheet="Lead",Row=7,Column=1,Rows=2,Columns=2 },
            new SummaryAreaPlan { Sheet="Lead",Row=3,Column=1,Rows=1,Columns=2 } };
        var leadRows = SummaryAreaEngine.LeadRows(leadParts);
        check("Lead 分段按Area和区域内行序读取",leadRows.Count==3&&leadRows[0].Row==7&&leadRows[1].Row==8&&leadRows[2].Row==3,"row order changed");
        check("Lead 未选中间行不进入名单",leadRows.TrueForAll(r=>r.Row==7||r.Row==8||r.Row==3),"gap rows included");
        leadParts = new List<SummaryAreaPlan> {
            new SummaryAreaPlan { Sheet="Lead",Row=3,Column=4,Rows=2,Columns=1 },
            new SummaryAreaPlan { Sheet="Lead",Row=3,Column=1,Rows=2,Columns=1 } };
        leadRows=SummaryAreaEngine.LeadRows(leadParts);
        check("Lead 不相邻两列按实际行配对且左编码右全称",leadRows.Count==2&&leadRows[0].CodeColumn==1&&leadRows[0].NameColumn==4&&leadRows[1].Row==4,"column pairing shifted");
        leadParts[1].Rows=1; rejects(delegate { SummaryAreaEngine.LeadRows(leadParts); },"Lead 缺少同行配对阻断而不串行");leadParts[1].Rows=2;
        leadParts.Add(new SummaryAreaPlan { Sheet="Lead",Row=3,Column=6,Rows=2,Columns=1 }); rejects(delegate { SummaryAreaEngine.LeadRows(leadParts); },"Lead 超过两列阻断");leadParts.RemoveAt(2);
        leadParts.Add(new SummaryAreaPlan { Sheet="Lead",Row=3,Column=1,Rows=1,Columns=1 });rejects(delegate { SummaryAreaEngine.LeadRows(leadParts); },"Lead 重叠选区阻断");leadParts.RemoveAt(2);
        leadParts[1].Sheet="其他";rejects(delegate { SummaryAreaEngine.LeadRows(leadParts); },"Lead 跨表选区阻断");
        var compiled=SummaryExistingFormula.Read("=ROUND('N:\\旧目录\\[公司甲.xlsx]B S'!L17+A1-$B$2,2)");
        compiled.Decisions[@"N:\旧目录\[公司甲.xlsx]"]=true;
        string rewritten=compiled.Render(@"C:\新目录","公司乙",2,5);
        check("已有外链更换路径文件且源地址不平移",rewritten=="=ROUND('C:\\新目录\\[公司乙.xlsx]B S'!L17+F3-$B$2,2)",rewritten);
        var mixed=SummaryExistingFormula.Read("='C:\\旧\\[甲.xlsx]BS'!$A$1+'C:\\公共\\[参数.xlsx]BS'!B2+A1");
        rejects(delegate { mixed.Render(@"C:\新","乙",0,5); },"未确认来源阻断公式生成");
        mixed.Decisions[@"C:\旧\[甲.xlsx]"]=true;mixed.Decisions[@"C:\公共\[参数.xlsx]"]=false;
        rewritten=mixed.Render(@"C:\新","乙",0,5);
        check("公共引用及源相对地址保持原文",rewritten=="='C:\\新\\[乙.xlsx]BS'!$A$1+'C:\\公共\\[参数.xlsx]BS'!B2+F1",rewritten);
        rejects(delegate {SummaryExistingFormula.Read("='C:\\[甲.xlsx]BS'!A1+UnknownName");},"已有外链仍阻断未知名称");
        var quoted=SummaryExistingFormula.Read("='C:\\O''Brien\\[甲.xlsx]表''一'!A1");quoted.Decisions[@"C:\O'Brien\[甲.xlsx]"]=true;
        check("已有引用单引号正确转义",quoted.Render(@"D:\O'Brien","乙",0,1)=="='D:\\O''Brien\\[乙.xlsx]表''一'!A1","quote mismatch");
        var literalFormula=SummaryExistingFormula.Read("=IF(A1=1,\"[假.xlsx]BS!A1\",0)");
        check("公式字符串不作为外链识别",literalFormula.Spans.TrueForAll(x=>x.Kind==RefKind.Local),"literal parsed as external");
        var cross=SummaryExistingFormula.Read("='公共表'!A1+A1");cross.Decisions["本工作簿!公共表"]=false;
        check("明确保留本簿跨表来源",cross.Render(@"C:\新","乙",1,2)=="='公共表'!A1+C2","cross-sheet shifted");
        check("竖向年度直接转置",SummaryHeaderFormula.Years(new SummaryAreaPlan{Sheet="Lead",Row=2,Column=14,Rows=5,Columns=1})=="=TRANSPOSE('Lead'!$N$2:$N$6)","transpose mismatch");
        check("公司公式直接关联编码简称列",SummaryHeaderFormula.Cell("Lead",3,15)=="='Lead'!$O$3","company link mismatch");
        check("表名引号正确转义",SummaryHeaderFormula.Cell("公'司",3,15)=="='公''司'!$O$3","quote mismatch");
        var preset = new BatchPreset { Feature="summary", Name="区域测试", Context="test.xlsx", Values=new Dictionary<string,string> {
            {"mode","area-mvp"},{"templateSheet","汇总"},{"templateAreas","3,12,4,5;20,12,11,5"},{"pasteSheet","汇总"},{"pasteArea","3,17,1,1"},{"leadSheet","Lead"},{"leadAreas","3,1,2,1;3,4,2,1"},{"projectFolder",@"C:\项目"},{"checkFiles","False"} } };
        var settings=SummaryAreaSettings.Read(BatchPresetStore.Decode(BatchPresetStore.Encode(preset),"summary"));
        check("区域预设完整保存Ctrl模板与Lead字段",settings.Template.Count==2&&settings.Lead.Count==2&&settings.Paste.Column==17&&settings.Template[1].RowOffset==17,"preset lost area metadata");
        check("预设保留路径及默认文件检查选项",settings.Folder==@"C:\项目"&&!settings.CheckFiles,"preset changed options");
        check("旧区域预设表头默认关闭",!settings.HeadersEnabled&&settings.CompanyHeader==null,"old preset enabled headers");
        preset.Values["headersEnabled"]="True"; rejects(delegate { SummaryAreaSettings.Read(preset); },"启用表头预设缺少选区阻断");
        preset.Values["companyHeader"]="2,17,1,1";preset.Values["companyHeaderSheet"]="汇总";preset.Values["yearSource"]="1,7,1,5";preset.Values["yearSourceSheet"]="Lead";preset.Values["yearTarget"]="7,17,1,1";preset.Values["yearTargetSheet"]="汇总";
        settings=SummaryAreaSettings.Read(preset);check("表头三组选区与开关可保存",settings.HeadersEnabled&&settings.YearSource.Columns==5&&settings.CompanyHeader.Row==2,"header preset lost");
        preset.Values["yearSource"]="1,7,2,5";rejects(delegate { SummaryAreaSettings.Read(preset); },"年度来源预设多行阻断");preset.Values["yearSource"]="1,7,1,5";
        preset.Values["yearTargetSheet"]="其他";rejects(delegate { SummaryAreaSettings.Read(preset); },"表头目标预设跨表阻断");preset.Values["yearTargetSheet"]="汇总";
        preset.Values.Remove("mode");rejects(delegate { SummaryAreaSettings.Read(preset); },"旧整列预设不猜测转换为区域预设");preset.Values["mode"]="area-mvp";
        preset.Values["pasteArea"]="3,12,1,1";rejects(delegate { SummaryAreaSettings.Read(preset); },"预设重叠模板在应用前阻断");preset.Values["pasteArea"]="3,17,1,1";
        preset.Values["leadAreas"]="3,1,2,1;3,4,1,1";rejects(delegate { SummaryAreaSettings.Read(preset); },"预设Lead同行缺失阻断");
        var matrixArea=new SummaryAreaPlan { Sheet="汇总",Row=3,Column=12,Rows=2,Columns=2 };
        var block=new List<SummaryAreaCell> { new SummaryAreaCell { Row=3,Column=12,Text="BS!$A$1" },new SummaryAreaCell { Row=3,Column=13,Formula=true,R1C1="=ROUND(RC[-1]-RC[-1],2)" } };
        var matrix=SummaryAreaEngine.FormulaMatrix(new object[,]{{"BS!$A$1","=ROUND(Q3-Q3,2)"},{null,null}},matrixArea,block,@"C:\项目","甲");
        check("批量矩阵保留空白与原生复制公式",matrix[1,0]==null&&Convert.ToString(matrix[0,1])=="=ROUND(Q3-Q3,2)","bulk write changed blanks/native formula");
        SummaryAreaEngine.VerifyMatrix(matrix,new object[,]{{null,"=ROUND(RC[-1]-RC[-1],2)"},{null,null}},matrixArea,block,@"C:\项目","甲");
        matrix[0,0]="=WRONG";rejects(delegate { SummaryAreaEngine.VerifyMatrix(matrix,new object[,]{{null,"=ROUND(RC[-1]-RC[-1],2)"},{null,null}},matrixArea,block,@"C:\项目","甲"); },"批量公式读回不符阻断");
        check("批量读取识别Excel HRESULT错误码",SummaryAreaEngine.IsError(unchecked((int)0x800A07D7))&&!SummaryAreaEngine.IsError(-25.0),"error values confused with ordinary negative numbers");
        string xml="<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'><Styles><Style ss:ID='Default'><Font ss:Color='#000000'/></Style></Styles><Worksheet><Table><Row><Cell ss:Formula='=1+1'><Data ss:Type='Number'>2</Data><Comment><Data>批注</Data></Comment></Cell></Row></Table></Worksheet></Workbook>";
        check("整块状态忽略公式缓存变化",SummaryAreaEngine.XmlState(xml)==SummaryAreaEngine.XmlState(xml.Replace(">2</Data>",">3</Data>")),"formula cache treated as edit");
        check("整块状态保留格式变化",SummaryAreaEngine.XmlState(xml)!=SummaryAreaEngine.XmlState(xml.Replace("#000000","#123456")),"font change lost");
        check("整块状态保留批注变化",SummaryAreaEngine.XmlState(xml)!=SummaryAreaEngine.XmlState(xml.Replace("批注","另一批注")),"comment change lost");
        using (var form = new SingleSummaryForm(true))
        {
            check("Ribbon MVP使用四组参数视图", form.Controls.Count==1 && form.Controls[0] is SummaryAreaView, "view missing");
            var view=(SummaryAreaView)form.Controls[0]; check("来源文件检查默认关闭",!view.CheckFiles.Checked,"default checked");
            check("项目文件夹允许手动输入粘贴",!view.Folder.ReadOnly&&view.Folder.ShortcutsEnabled&&view.Folder.Enabled,"folder locked");
            view.Folder.Text="  \"C:\\项目文件夹\"  ";check("手动粘贴路径支持外层引号",view.ProjectFolder==@"C:\项目文件夹","quoted folder not normalized");
            int changes=0;view.Changed+=delegate { changes++; };view.Folder.Text=@"C:\另一文件夹";check("修改文件夹触发预检失效",changes==1,"path edit kept previous preview");
            var app=new SummaryAreaFakeApp();
            using(var host=new SummaryAreaHost(app,form,view))
            {
                check("MVP复用原保存预设齿轮",view.Controls.Find("batchPresetSettings",true).Length==1,"preset gear absent");
                var flags=BindingFlags.Instance|BindingFlags.NonPublic;
                var folderCell = new SummaryFolderCell { Value2 = "  C:\\底稿路径  " };
                app.Selection = folderCell;
                var pickFolder = typeof(SummaryAreaHost).GetMethod("Pick", flags);
                pickFolder.Invoke(host, new object[] { "folder", false });
                check("项目文件夹从单元格文字读取", view.ProjectFolder == @"C:\底稿路径", "cell path missing");
                check("路径来源不绑定业务工作簿", typeof(SummaryAreaHost).GetField("book",flags).GetValue(host)==null, "folder cell bound workbook");
                foreach (object invalid in new object[] { "", 123, DBNull.Value })
                {
                    folderCell.Value2 = invalid; pickFolder.Invoke(host, new object[] { "folder", false });
                    check("无效路径文字保留旧值", view.ProjectFolder == @"C:\底稿路径", "invalid cell overwrote folder");
                }
                folderCell.Value2 = @"C:\其他路径"; folderCell.Width = 2;
                pickFolder.Invoke(host, new object[] { "folder", false });
                check("路径拒绝多单元格选区", view.ProjectFolder == @"C:\底稿路径", "multiple cells accepted");
                app.Selection = null; pickFolder.Invoke(host, new object[] { "folder", false });
                check("取消路径选区保留旧值", view.ProjectFolder == @"C:\底稿路径", "cancel overwrote folder");
                Action<string,object> field=(name,value)=>typeof(SummaryAreaHost).GetField(name,flags).SetValue(host,value);
                field("book",app.Book);field("bookName","test.xlsx");field("templateSheet","汇总");field("backup","备份");field("targets",SummaryAreaEngine.Targets(areas,3,17,2));
                check("溢出空状态不等同于True",SummaryAreaHost.FormulaFlag(DBNull.Value,"test","HasSpill")==-1&&SummaryAreaHost.FormulaFlag(null,"test","HasSpill")==-1,"null misclassified");
                check("溢出布尔及VARIANT_BOOL解析",SummaryAreaHost.FormulaFlag(false,"test","HasSpill")==0&&SummaryAreaHost.FormulaFlag(true,"test","HasSpill")==1&&SummaryAreaHost.FormulaFlag(-1,"test","HasSpill")==1,"boolean wrong");
                rejects(delegate {SummaryAreaHost.FormulaFlag(unchecked((int)0x800A07D7),"test","HasSpill");},"溢出错误码不得强转True");
                var guard=typeof(SummaryAreaHost).GetMethod("RejectArraySpill",flags);
                using(var guardBag=new ComBag())
                {
                    guard.Invoke(host,new object[]{new SummaryArrayOnlyRange(),"legacy",guardBag});check("旧宿主缺少HasSpill成员兼容",true,"");
                    guard.Invoke(host,new object[]{new SummaryMixedSpillRange(false,true),"merged",guardBag});check("合并区域混合状态检查左上角后通过",true,"");
                    bool blockedSpill=false;try{guard.Invoke(host,new object[]{new SummaryMixedSpillRange(true,false),"spill",guardBag});}catch(TargetInvocationException e){blockedSpill=e.InnerException is InvalidOperationException;}check("混合区域中的真实溢出仍阻断",blockedSpill,"actual spill allowed");
                    blockedSpill=false;try{guard.Invoke(host,new object[]{new SummaryMixedSpillRange(null,false),"unknown",guardBag});}catch(TargetInvocationException e){blockedSpill=e.InnerException is InvalidOperationException;}check("单格状态未知仍阻断且不宣称溢出",blockedSpill,"unknown state allowed");
                }
                var restore=typeof(SummaryAreaHost).GetMethod("Restore",flags);
                restore.Invoke(host,new object[]{4});
                check("生产回滚只访问四个Area",app.Book.Target.Copied.Count==4 && app.Book.Target.Copied[0]=="Q3:U6" && app.Book.Target.Copied[1]=="Q20:U30" && app.Book.Target.Copied[3]=="V20:Z30","rollback touched gaps");
                check("回滚恢复宿主状态",app.EnableEvents&&app.ScreenUpdating&&app.AskToUpdateLinks&&app.DisplayAlerts,"host state leaked");
                app.Book.Target.Fail=true;bool failed=false;try{restore.Invoke(host,new object[]{1});}catch(TargetInvocationException){failed=true;}
                check("回滚写入失败明确抛出",failed,"restore failure swallowed");
                check("失败仍恢复宿主状态",app.EnableEvents&&app.ScreenUpdating&&app.AskToUpdateLinks&&app.DisplayAlerts,"failed host state leaked");
                field("pending",true);field("failed",true);field("step",4);
                bool blocked=false;try{typeof(SummaryAreaHost).GetMethod("Confirm",flags).Invoke(host,null);}catch(TargetInvocationException){blocked=true;}
                check("失败批次不能确认丢失恢复信息",blocked,"failed session confirmed");
            }
        }
    }
}
public sealed class SummaryAreaFakeApp
{
    public object Selection { get; set; }
    public int Calculation { get; set; }
    public bool EnableEvents{get;set;} public bool ScreenUpdating{get;set;} public bool AskToUpdateLinks{get;set;} public bool DisplayAlerts{get;set;}
    public SummaryAreaFakeBook Book=new SummaryAreaFakeBook();
    public SummaryAreaFakeApp(){Calculation=-4105;EnableEvents=ScreenUpdating=AskToUpdateLinks=DisplayAlerts=true;}
    public SummaryAreaFakeBooks Workbooks{get{return new SummaryAreaFakeBooks(Book);}}
}
public sealed class SummaryFolderCell
{
    public object Value2 { get; set; }
    public object Parent { get { return new object(); } }
    public int Row { get { return 1; } } public int Column { get { return 1; } }
    public int Width = 1;
    public string Address { get { return "$A$1"; } }
    public SummarySpillCount Rows { get { return new SummarySpillCount(1); } }
    public SummarySpillCount Columns { get { return new SummarySpillCount(Width); } }
    public SummarySpillCount Areas { get { return new SummarySpillCount(1); } }
}
public sealed class SummaryAreaFakeBooks
{
    private readonly SummaryAreaFakeBook book;public SummaryAreaFakeBooks(SummaryAreaFakeBook book){this.book=book;}public int Count{get{return 1;}}public object Item(int index){return book;}
}
public sealed class SummaryAreaFakeBook
{
    public string FullName{get{return "test.xlsx";}}public SummaryAreaFakeSheet Target=new SummaryAreaFakeSheet();public SummaryAreaFakeSheet Original=new SummaryAreaFakeSheet();public SummaryAreaFakeSheets Worksheets{get{return new SummaryAreaFakeSheets(this);}}
}
public sealed class SummaryAreaFakeSheets
{
    private readonly SummaryAreaFakeBook book;public SummaryAreaFakeSheets(SummaryAreaFakeBook b){book=b;}public object Item(string name){return name=="备份"?book.Original:book.Target;}
}
public sealed class SummaryAreaFakeSheet
{
    public readonly List<string> Copied=new List<string>();public bool Fail;public SummaryAreaFakeRanges Range{get{return new SummaryAreaFakeRanges(this);}}
}
public sealed class SummaryAreaFakeRanges
{
    private readonly SummaryAreaFakeSheet sheet;public SummaryAreaFakeRanges(SummaryAreaFakeSheet s){sheet=s;}public SummaryAreaFakeRange this[string address]{get{return new SummaryAreaFakeRange(sheet,address);}}
}
public sealed class SummaryAreaFakeRange
{
    private readonly SummaryAreaFakeSheet sheet;private readonly string address;public SummaryAreaFakeRange(SummaryAreaFakeSheet s,string a){sheet=s;address=a;}public void Clear(){if(sheet.Fail)throw new InvalidOperationException("injected restore failure");}public void Copy(object target){var r=(SummaryAreaFakeRange)target;r.sheet.Copied.Add(r.address);}
}

// Managed host fixture: a mixed aggregate and explicit per-cell states.
public sealed class SummaryMixedSpillRange
{
    private readonly object spill;private readonly bool merged;private readonly bool single;
    public SummaryMixedSpillRange(object value,bool merge):this(value,merge,false){}
    private SummaryMixedSpillRange(object value,bool merge,bool one){spill=value;merged=merge;single=one;}
    public object HasArray{get{return false;}} public object HasSpill{get{return single?spill:DBNull.Value;}} public bool MergeCells{get{return merged;}}
    public int Row{get{return 2;}}public int Column{get{return 17;}}public string Address{get{return "$Q$2";}}
    public SummarySpillCount Rows{get{return new SummarySpillCount(1);}}public SummarySpillCount Columns{get{return new SummarySpillCount(single?1:2);}}
    public SummaryMixedSpillRange MergeArea{get{return new SummaryMixedSpillRange(spill,false,false);}}
    public SummarySpillCells Cells{get{return new SummarySpillCells(new SummaryMixedSpillRange(spill,merged,true));}}
}
public sealed class SummarySpillCount{public int Count{get;private set;}public SummarySpillCount(int count){Count=count;}}
public sealed class SummarySpillCells{private readonly object cell;public SummarySpillCells(object value){cell=value;}public object Item(int row,int column){return cell;}}

public sealed class SummaryArrayOnlyRange{public bool HasArray{get{return false;}}}

public sealed class SummaryRecalcRange {public readonly SummaryRecalcSheet Worksheet=new SummaryRecalcSheet();public void Calculate(){throw new InvalidOperationException("Range Calculate unavailable");}}
public sealed class SummaryRecalcSheet {public int Calls;public bool Fail;public void Calculate(){Calls++;if(Fail)throw new InvalidOperationException("Sheet Calculate unavailable");}}
public sealed class SummaryRecalcApp {public int Calls;public void Calculate(){Calls++;}}
