using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 文件操作 → 批量处理 → 文件提取名称
// 控件：btnFileNames
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_btnFileNames(object control) {return GetButtonImageCore("btnFileNames");}

        public void OnFileNames(object control) {
            if(fileNames!=null&&!fileNames.IsDisposed){fileNames.Activate();return;}
            FileNameWorksheet target=null;try {if(!EnsureApplication())throw new InvalidOperationException("请先打开 WPS 表格。");target=new FileNameWorksheet(_app);fileNames=new FileNamesForm(target);target=null;var form=fileNames;form.FormClosed+=delegate {fileNames=null;form.Dispose();};ShowFileForm(form);}catch(Exception ex){if(target!=null)target.Dispose();UiDialog.Show(ex.Message,"文件提取名称",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        }
    }
}
