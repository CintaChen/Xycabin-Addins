using System;
using System.Collections.Generic;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace WpsShared
{
    public sealed class RestartBook
    {
        public string Name, Path, Error;
        public bool ReadOnly, Active, Reopened;
    }
    public sealed class RestartTicket
    {
#if WPS_WORD
        private const string ProductName="WordUpdater";
#else
        private const string ProductName="Excel_Iterate";
#endif
        public string Product=ProductName;
        public int Version=1, ProcessId, SessionId;
        public long Started;
        public string Id, User, Executable, State="Prepared", Error;
        public List<int> ExistingProcesses=new List<int>();
        public List<RestartBook> Books=new List<RestartBook>();
        public static string Root {get{return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),ProductName,"Restart");}}
        public static string FileFor(string id) {Guid value;if(!Guid.TryParseExact(id,"N",out value))throw new InvalidDataException("重启任务标识无效。");return System.IO.Path.Combine(Root,id+".xml");}
        public static void PrepareDirectory() {
            var acl=new DirectorySecurity();acl.SetAccessRuleProtection(true,false);acl.AddAccessRule(new FileSystemAccessRule(WindowsIdentity.GetCurrent().User,FileSystemRights.FullControl,InheritanceFlags.ContainerInherit|InheritanceFlags.ObjectInherit,PropagationFlags.None,AccessControlType.Allow));
            Directory.CreateDirectory(Root,acl);
        }
        public void Save() {
            PrepareDirectory();string file=FileFor(Id),temp=file+"."+Guid.NewGuid().ToString("N")+".tmp";
            try {using(var writer=XmlWriter.Create(temp,new XmlWriterSettings{Encoding=new UTF8Encoding(false),Indent=true}))new XmlSerializer(typeof(RestartTicket)).Serialize(writer,this);if(File.Exists(file))File.Replace(temp,file,null);else File.Move(temp,file);}
            finally {if(File.Exists(temp))File.Delete(temp);}
        }
        public static RestartTicket Load(string id) {
            using(var reader=XmlReader.Create(FileFor(id),new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=1024*1024})) {
                var ticket=(RestartTicket)new XmlSerializer(typeof(RestartTicket)).Deserialize(reader);
                if(ticket.Version!=1||ticket.Product!=ProductName||ticket.Id!=id||ticket.User!=WindowsIdentity.GetCurrent().User.Value||ticket.SessionId!=System.Diagnostics.Process.GetCurrentProcess().SessionId)throw new InvalidDataException("重启任务与当前用户或登录会话不匹配。");return ticket;
            }
        }
    }
}
