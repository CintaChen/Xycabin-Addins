using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Windows.Forms;
#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // Latest disk data is separate from the immutable appearance owned by each form.
    internal static class FormatSession
    {
        private static readonly object sync=new object();
        private static readonly ConditionalWeakTable<Form,CustomizationData> forms=new ConditionalWeakTable<Form,CustomizationData>();
        private static CustomizationData current=new CustomizationData();private static string fingerprint,source;
        internal static string SourceOverride;
        internal static string PathForUser {get{return SourceOverride??RibbonCustomization.Store.Path;}}
        internal static string Digest(string path){if(!File.Exists(path))return "missing";using(var stream=File.Open(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite|FileShare.Delete))using(var hash=SHA256.Create())return Convert.ToBase64String(hash.ComputeHash(stream));}
        internal static CustomizationData Data {get {lock(sync){
            string path=PathForUser;
            try {
                string stamp=Digest(path)+"|"+Digest(DialogSizeStore.PathForUser);
                if(path!=source||stamp!=fingerprint){
                    var store=new LocalCustomization(path);var sizes=DialogSizeStore.Read();
                    if(store.IsReadable&&sizes.Writable){var next=store.Data.Copy();next.WindowSizes=sizes.Values.ToDictionary(p=>p.Key,p=>new[]{p.Value.Width,p.Value.Height});current=next;source=path;fingerprint=stamp;}
                    else UiDiagnostics.Record("configuration-read",new InvalidDataException());
                }
            }catch(Exception error){UiDiagnostics.Record("configuration-read",error);}
            return current.Copy();
        }}}
        internal static CustomizationData Capture(Form form){lock(sync){CustomizationData value;if(!forms.TryGetValue(form,out value)){value=Data;forms.Add(form,value);}return value;}}
        internal static void Initialize(){var ignored=Data;}
    }
    internal static class LiveCustomization
    {
        private static Control dispatcher;private static FileSystemWatcher watcher;private static Timer debounce;private static string menuStamp;
        private static string MenuStamp(){var data=FormatSession.Data;data.Formats.Clear();data.WindowSizes=null;return Convert.ToBase64String(LocalCustomization.Encode(data));}
        internal static void Start(){
            if(dispatcher!=null)return;dispatcher=new Control();var handle=dispatcher.Handle;menuStamp=MenuStamp();
            debounce=new Timer {Interval=150};debounce.Tick+=delegate{debounce.Stop();Refresh();};
            string wanted=System.IO.Path.GetFullPath(FormatSession.PathForUser),root=System.IO.Path.GetDirectoryName(wanted);while(!Directory.Exists(root))root=System.IO.Path.GetDirectoryName(root);
            watcher=new FileSystemWatcher(root,System.IO.Path.GetFileName(wanted)){IncludeSubdirectories=true,NotifyFilter=NotifyFilters.LastWrite|NotifyFilters.FileName|NotifyFilters.Size};
            FileSystemEventHandler changed=delegate(object sender,FileSystemEventArgs e){if(string.Equals(e.FullPath,wanted,StringComparison.OrdinalIgnoreCase))Notify();};watcher.Changed+=changed;watcher.Created+=changed;watcher.Deleted+=changed;watcher.Renamed+=delegate(object sender,RenamedEventArgs e){changed(sender,e);};watcher.EnableRaisingEvents=true;
        }
        internal static void Notify(){var target=dispatcher;if(target==null||target.IsDisposed)return;try{target.BeginInvoke((MethodInvoker)delegate{if(debounce!=null){debounce.Stop();debounce.Start();}});}catch(InvalidOperationException){}}
        private static void Refresh(){try{string next=MenuStamp();if(next==menuStamp)return;RibbonCustomization.Refresh();menuStamp=next;}catch(Exception error){UiDiagnostics.Record("ribbon-refresh",error);}}
        internal static void Stop(){if(watcher!=null){watcher.Dispose();watcher=null;}if(debounce!=null){debounce.Dispose();debounce=null;}if(dispatcher!=null){dispatcher.Dispose();dispatcher=null;}menuStamp=null;}
    }
}
