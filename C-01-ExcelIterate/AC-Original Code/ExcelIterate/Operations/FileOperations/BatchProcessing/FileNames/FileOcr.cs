using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_Iterate
{
    [DataContract] internal sealed class OcrSegment {
        [DataMember(Name="text")] internal string Text;
        [DataMember(Name="left")] internal double Left;
        [DataMember(Name="top")] internal double Top;
        [DataMember(Name="width")] internal double Width;
        [DataMember(Name="height")] internal double Height;
        [DataMember(Name="score")] internal double? Score;
        internal int Page,Sequence;
        internal bool NoLocation;
        internal string LineId;
    }
    internal interface IFileOcr:IDisposable {Task<List<OcrSegment>> RecognizeAsync(string image,int page,CancellationToken token);}
    internal sealed class OcrFailure:Exception {internal readonly bool Pause;internal OcrFailure(string message,bool pause):base(message){Pause=pause;}}
    internal static class OcrDiagnostics {
        internal static string RootOverride;
        private static string assemblyDigest;
        internal static string Safe(string value){return new string((value??"-").Where(c=>((c>='a'&&c<='z')||(c>='A'&&c<='Z')||(c>='0'&&c<='9'))||c=='.'||c=='_'||c=='-').Take(100).ToArray());}
        internal static void Cloud(string stage,string provider,string method,string code,string requestId){try{Directory.CreateDirectory(Path.GetDirectoryName(LogPath));lock(typeof(OcrDiagnostics)){if(File.Exists(LogPath)&&new FileInfo(LogPath).Length>256*1024)File.WriteAllText(LogPath,"",Encoding.UTF8);File.AppendAllText(LogPath,DateTime.UtcNow.ToString("o")+" stage="+Safe(stage)+" provider="+Safe(provider)+" interface="+Safe(method)+" code="+Safe(code)+" request="+Safe(requestId)+Environment.NewLine,Encoding.UTF8);}}catch{}}
        internal static string AssemblyDigest {get {if(assemblyDigest==null)try {using(var stream=File.OpenRead(typeof(Connect).Assembly.Location))using(var sha=SHA256.Create())assemblyDigest=BitConverter.ToString(sha.ComputeHash(stream)).Replace("-","").ToLowerInvariant();}catch {assemblyDigest="unavailable";}return assemblyDigest;}}
        internal static string ErrorCode(Exception error) {var win32=error as Win32Exception;return error.GetType().Name+"，HRESULT="+error.HResult.ToString("X8")+(win32==null?"":"，Win32="+win32.NativeErrorCode);}
        internal static string LogPath {get {return Path.Combine(RootOverride??Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Excel_Iterate","OcrDiagnostics.log");}}
        internal static void Record(string stage,Exception error,int? exit=null) {
            // Never record messages, paths, requests, recognized text or credentials.
            try {Directory.CreateDirectory(Path.GetDirectoryName(LogPath));lock(typeof(OcrDiagnostics)) {if(File.Exists(LogPath)&&new FileInfo(LogPath).Length>256*1024)File.WriteAllText(LogPath,"",Encoding.UTF8);var native=error as Win32Exception;File.AppendAllText(LogPath,DateTime.UtcNow.ToString("o")+" stage="+stage+" type="+(error==null?"none":error.GetType().Name)+" hresult="+(error==null?"0":error.HResult.ToString("X8"))+" win32="+(native==null?"-":native.NativeErrorCode.ToString())+" exit="+(exit.HasValue?exit.Value.ToString():"-")+" host="+(Environment.Is64BitProcess?"x64":"x86")+" pid="+Process.GetCurrentProcess().Id+" trusted="+AppDomain.CurrentDomain.IsFullyTrusted+" assembly="+AssemblyDigest+Environment.NewLine,Encoding.UTF8);}}
            catch {}
        }
    }
    internal static class OcrJson {
        internal static T Read<T>(byte[] bytes) {using(var m=new MemoryStream(bytes))return (T)new DataContractJsonSerializer(typeof(T)).ReadObject(m);}
    }
    internal static class OcrFeatures {
        internal static string Normalize(string text) {var result=new StringBuilder();foreach(char c in text??""){if(char.IsWhiteSpace(c))continue;result.Append(c>=0xff01&&c<=0xff5e&&!char.IsLetterOrDigit(c)?(char)(c-0xfee0):c);}return result.ToString();}
        internal static List<string> Extract(IEnumerable<OcrSegment> segments,string prefix) {
            prefix=Normalize(prefix);if(prefix.Length==0)throw new ArgumentException("请填写需要提取的前缀。");
            var result=new List<string>();
            var material=segments.ToList();
            foreach(var line in material.Where(s=>s.NoLocation).OrderBy(s=>s.Page).ThenBy(s=>s.Sequence))foreach(var text in (line.Text??"").Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries))ExtractText(text,prefix,result);
            var rows=new List<List<OcrSegment>>();
            foreach(var s in material.Where(s=>!s.NoLocation).OrderBy(s=>s.Page).ThenBy(s=>s.Top).ThenBy(s=>s.Left)) {
                var row=rows.LastOrDefault(r=>r[0].Page==s.Page&&r.Any(p=>Math.Min(p.Top+p.Height,s.Top+s.Height)-Math.Max(p.Top,s.Top)>=Math.Min(p.Height,s.Height)*.5));
                if(row==null){row=new List<OcrSegment>();rows.Add(row);}row.Add(s);
            }
            foreach(var row in rows) {
                var runs=new List<List<OcrSegment>>();var ordered=row.OrderBy(s=>s.Left).ToList();
                double spacing=Math.Max(1,ordered.Select(s=>s.Height).OrderBy(h=>h).ElementAt(ordered.Count/2))*1.5;
                foreach(var s in ordered){var run=runs.LastOrDefault();var prior=run==null?null:run.Last();if(prior==null||s.Left-(prior.Left+prior.Width)>spacing||(Normalize(s.Text).Contains(":")&&!Normalize(s.Text).StartsWith(prefix,StringComparison.Ordinal)&&Normalize(string.Concat(run.Select(x=>x.Text))).Contains(prefix))){run=new List<OcrSegment>();runs.Add(run);}run.Add(s);}
                foreach(var run in runs) {
                    ExtractText(string.Concat(run.Select(s=>s.Text)),prefix,result);
                }
            }
            return result;
        }
        private static void ExtractText(string text,string prefix,List<string> result) {
            text=Normalize(text);int from=0;
            while(from<text.Length){int at=text.IndexOf(prefix,from,StringComparison.Ordinal);if(at<0)break;int start=at+prefix.Length;int next=text.IndexOf(prefix,start,StringComparison.Ordinal);int valueStart=start;if(!prefix.EndsWith(":",StringComparison.Ordinal)&&valueStart<text.Length&&text[valueStart]==':')valueStart++;string value=text.Substring(valueStart,(next<0?text.Length:next)-valueStart);if(value.Length>0&&!result.Contains(value))result.Add(value);from=start;}
        }

    }
    internal static class OcrCredentials {
        internal static string RootOverride;
        internal static string PathForUser {get{return CredentialPath(false);}}
        internal static string CredentialPath(bool aliyun){return Path.Combine(RootOverride??Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Excel_Iterate",aliyun?"OcrCredentials.Aliyun.bin":"OcrCredentials.bin");}
        internal static string[] Read(){return Read(false);}
        internal static string[] Read(bool aliyun) {
            string PathForUser=CredentialPath(aliyun);
            try {if(!File.Exists(PathForUser))return new[]{"",""};var text=Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(PathForUser),null,DataProtectionScope.CurrentUser));var parts=text.Split('\n');return parts.Length==2?parts:new[]{"",""};}catch {return new[]{"",""};}
        }
        internal static void Clear(){Clear(false);}
        internal static void Clear(bool aliyun){string path=CredentialPath(aliyun);if(File.Exists(path))File.Delete(path);}
        internal static void Save(string api,string secret){Save(false,api,secret);}
        internal static void Save(bool aliyun,string api,string secret) {
            string PathForUser=CredentialPath(aliyun);
            Directory.CreateDirectory(Path.GetDirectoryName(PathForUser));string temp=PathForUser+"."+Guid.NewGuid().ToString("N")+".tmp";
            try {File.WriteAllBytes(temp,ProtectedData.Protect(Encoding.UTF8.GetBytes(api+"\n"+secret),null,DataProtectionScope.CurrentUser));if(File.Exists(PathForUser))File.Replace(temp,PathForUser,null);else File.Move(temp,PathForUser);}finally {if(File.Exists(temp))File.Delete(temp);}
        }
    }
    [DataContract] internal sealed class BaiduTokenReply {
        [DataMember(Name="access_token")] internal string Token;
        [DataMember(Name="expires_in")] internal long Expires;
        [DataMember(Name="error")] internal string Error;
    }
    [DataContract] internal sealed class BaiduWords {
        [DataMember(Name="words")] internal string Text;
        [DataMember(Name="location")] internal BaiduLocation Location;
        [DataMember(Name="probability")] internal BaiduProbability Probability;
    }
    [DataContract] internal sealed class BaiduLocation {
        [DataMember(Name="left")] internal double Left;
        [DataMember(Name="top")] internal double Top;
        [DataMember(Name="width")] internal double Width;
        [DataMember(Name="height")] internal double Height;
    }
    [DataContract] internal sealed class BaiduProbability {[DataMember(Name="average")]internal double Average;}
    [DataContract] internal sealed class BaiduOcrReply {
        [DataMember(Name="words_result")] internal List<BaiduWords> Words;
        [DataMember(Name="error_code")] internal int Error;
        [DataMember(Name="log_id")] internal ulong LogId;
    }
    internal sealed class BaiduFileOcr:IFileOcr {
        private readonly HttpClient http;
        private readonly string api,secret,method;
        private readonly OcrMode mode;
        private string access;private DateTime expires;
        internal BaiduFileOcr(string api,string secret,bool accurate):this(api,secret,accurate,null){}
        internal BaiduFileOcr(string api,string secret,bool accurate,HttpMessageHandler handler):this(api,secret,accurate?OcrMode.Accurate:OcrMode.General,handler){}
        internal BaiduFileOcr(string api,string secret,OcrMode mode):this(api,secret,mode,null){}
        internal BaiduFileOcr(string api,string secret,OcrMode mode,HttpMessageHandler handler) {
            if(string.IsNullOrWhiteSpace(api)||string.IsNullOrWhiteSpace(secret))throw new OcrFailure("请填写百度云应用的 API Key 和 Secret Key。它们用于验证身份，并使用你账号下的 OCR 权限与配额；未填写无法调用百度云识别。",true);
            if(mode==OcrMode.AliyunAdvanced)throw new ArgumentOutOfRangeException("mode");this.mode=mode;this.api=api;this.secret=secret;method=OcrModes.Method(mode);http=handler==null?new HttpClient():new HttpClient(handler);http.Timeout=TimeSpan.FromSeconds(120);
        }
        internal static string ImageBody(byte[] bytes) {return "image="+Encode(bytes);}
        internal static string PdfBody(byte[] bytes,int page) {return "pdf_file="+Encode(bytes)+"&pdf_file_num="+page.ToString(CultureInfo.InvariantCulture);}
        private static string Encode(byte[] bytes) {return Convert.ToBase64String(bytes).Replace("+","%2B").Replace("/","%2F").Replace("=","%3D");}
        private static string RequestBody(string file,int page,int limit) {
            if(!OcrImages.IsPdf(file)) {
                string image=ImageBody(File.ReadAllBytes(file));
                if(image.Length>limit)throw new OcrFailure("图片编码后超过所选接口限制，请压缩图片后重试。",false);
                return image;
            }
            if(!OcrImages.LooksLikePdf(file))throw new OcrFailure("文件后缀为 PDF，但内容不是有效的 PDF 文件，请确认未被改名或损坏。",false);
            string pdf=PdfBody(File.ReadAllBytes(file),page);
            if(pdf.Length>limit)throw new OcrFailure("PDF 编码后超过所选接口限制，请压缩 PDF 后重试。",false);
            return pdf;
        }
        private async Task<string> Access(CancellationToken cancel) {
            if(!string.IsNullOrEmpty(access)&&DateTime.UtcNow<expires)return access;
            using(var body=new FormUrlEncodedContent(new Dictionary<string,string>{{"grant_type","client_credentials"},{"client_id",api},{"client_secret",secret}}))
            using(var response=await http.PostAsync("https://aip.baidubce.com/oauth/2.0/token",body,cancel).ConfigureAwait(false)) {
                if(!response.IsSuccessStatusCode){OcrDiagnostics.Cloud("auth","baidu",method,((int)response.StatusCode).ToString(),null);throw new OcrFailure("百度云鉴权失败，请核对密钥和网络。",true);}
                BaiduTokenReply result;try {result=OcrJson.Read<BaiduTokenReply>(await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false));}catch {throw new OcrFailure("百度云鉴权结果无法读取。",true);}
                if(string.IsNullOrWhiteSpace(result.Token)||result.Token.Trim()=="None"||result.Expires<=0||!string.IsNullOrEmpty(result.Error)){OcrDiagnostics.Cloud("auth","baidu",method,result.Error??"InvalidToken",null);throw new OcrFailure("百度云鉴权失败，请检查 API Key 和 Secret Key 是否正确。",true);}
                access=result.Token;expires=DateTime.UtcNow.AddSeconds(Math.Min(result.Expires,31536000)-Math.Min(60,result.Expires/10));return access;
            }
        }
        public async Task<List<OcrSegment>> RecognizeAsync(string file,int page,CancellationToken cancel) {
            CancellationToken requested=cancel;
            using(var deadline=CancellationTokenSource.CreateLinkedTokenSource(cancel)) {
            deadline.CancelAfter(TimeSpan.FromSeconds(120));cancel=deadline.Token;
            try {
                string body=RequestBody(file,page,OcrModes.Limit(mode));
                body+="&language_type=CHN_ENG&detect_direction=true&probability=true";if(OcrModes.Positioned(mode))body+="&recognize_granularity=big";
                for(int attempt=0;attempt<2;attempt++) {
                    string token=await Access(cancel).ConfigureAwait(false);
                    using(var content=new StringContent(body,Encoding.UTF8,"application/x-www-form-urlencoded"))
                    using(var response=await http.PostAsync("https://aip.baidubce.com/rest/2.0/ocr/v1/"+method+"?access_token="+token,content,cancel).ConfigureAwait(false)) {
                        if(!response.IsSuccessStatusCode){OcrDiagnostics.Cloud("recognize","baidu",method,((int)response.StatusCode).ToString(),null);throw new OcrFailure("百度云请求失败（HTTP "+(int)response.StatusCode+"），请检查网络与服务状态。",new[]{401,403,429}.Contains((int)response.StatusCode));}
                        BaiduOcrReply result;try {result=OcrJson.Read<BaiduOcrReply>(await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false));}catch {throw new OcrFailure("百度云识别结果无法读取。",false);}
                        OcrDiagnostics.Cloud("recognize","baidu",method,result.Error.ToString(),result.LogId.ToString());
                        if((result.Error==110||result.Error==111)&&attempt==0){access=null;continue;}
                        if(result.Error!=0)throw new OcrFailure("百度云识别失败（错误码 "+result.Error+"）"+(new[]{6,14,17,18,19,110,111,216100}.Contains(result.Error)?"，请检查应用权限、配额或密钥。":"，请检查文件或稍后手动重试。"),new[]{6,14,17,18,19,110,111}.Contains(result.Error));
                        if(result.Words==null)throw new OcrFailure("百度云未返回有效的文字结果。",false);
                        return result.Words.Select((w,i)=>new OcrSegment {Text=w.Text,Left=w.Location==null?0:w.Location.Left,Top=w.Location==null?0:w.Location.Top,Width=w.Location==null?0:w.Location.Width,Height=w.Location==null?0:w.Location.Height,NoLocation=!OcrModes.Positioned(mode)||w.Location==null,Sequence=i,LineId=(!OcrModes.Positioned(mode)||w.Location==null)?i.ToString():null,Score=w.Probability==null?(double?)null:w.Probability.Average,Page=page}).ToList();
                    }
                }
                throw new OcrFailure("百度云 Token 失效，请重新填写密钥后重试。",true);
            }catch(OperationCanceledException) {if(requested.IsCancellationRequested)throw;OcrDiagnostics.Cloud("request","baidu",method,"Timeout",null);throw new OcrFailure("识别超时；未自动重复请求，请手动重试。",false);}
            catch(HttpRequestException) {OcrDiagnostics.Cloud("request","baidu",method,"Network",null);throw new OcrFailure("百度云网络请求失败，请检查网络连接后手动重试。",false);}
            }
        }
        public void Dispose() {access=null;http.Dispose();}
    }
    internal static class OcrImages {
        internal static readonly string[] Extensions={".png",".jpg",".jpeg",".bmp",".tif",".tiff"};
        internal const string PdfExtension=".pdf";
        internal static bool Supported(string path) {return IsImage(path)||IsPdf(path);}
        internal static bool IsImage(string path) {return Extensions.Contains(Path.GetExtension(path).ToLowerInvariant());}
        internal static bool IsPdf(string path) {return Path.GetExtension(path).ToLowerInvariant()==PdfExtension;}
        // A PDF travels to the cloud unchanged, so only its header can be verified locally.
        internal static bool LooksLikePdf(string path) {
            try {using(var stream=File.OpenRead(path)) {byte[] head=new byte[1024];int read=stream.Read(head,0,head.Length);return read>=5&&Encoding.ASCII.GetString(head,0,read).IndexOf("%PDF",StringComparison.Ordinal)>=0;}}
            catch(IOException){return false;}catch(UnauthorizedAccessException){return false;}
        }
        internal static int Pages(string path) {using(var image=Image.FromFile(path))return image.FrameDimensionsList.Contains(FrameDimension.Page.Guid)?image.GetFrameCount(FrameDimension.Page):1;}
        internal static string Prepare(string source,int page,int longest,string directory) {
            return Prepare(source,page,longest,directory,0);
        }
        internal static string Prepare(string source,int page,OcrMode mode,string directory){return PrepareCore(source,page,OcrModes.Longest(mode),directory,OcrModes.Limit(mode),mode==OcrMode.AliyunAdvanced);}
        internal static string Prepare(string source,int page,int longest,string directory,int encodedLimit){return PrepareCore(source,page,longest,directory,encodedLimit,false);}
        private static string PrepareCore(string source,int page,int longest,string directory,int encodedLimit,bool binary) {
            using(var image=Image.FromFile(source)) {
                if(image.FrameDimensionsList.Contains(FrameDimension.Page.Guid))image.SelectActiveFrame(FrameDimension.Page,page-1);
                int minimum=binary?16:15;if(binary&&(double)Math.Max(image.Width,image.Height)/Math.Min(image.Width,image.Height)>=50)throw new OcrFailure("阿里云要求图片长宽比小于 50，未上传。",false);
                double factor=Math.Min(1,(double)longest/Math.Max(image.Width,image.Height));string target=Path.Combine(directory,Guid.NewGuid().ToString("N")+".png");
                try {for(int attempt=0;attempt<6;attempt++) {
                    int w=(int)Math.Round(image.Width*factor),h=(int)Math.Round(image.Height*factor);
                    if(w<minimum||h<minimum)throw new OcrFailure("图片处理后最短边不足 "+minimum+" 像素，请使用清晰的原始图片。",false);
                    using(var bitmap=new Bitmap(w,h))using(var g=Graphics.FromImage(bitmap)){g.Clear(Color.White);g.DrawImage(image,0,0,w,h);bitmap.Save(target,ImageFormat.Png);}
                    if(encodedLimit<=0)return target;int length=binary?(int)new FileInfo(target).Length:BaiduFileOcr.ImageBody(File.ReadAllBytes(target)).Length;if(length<=encodedLimit)return target;
                    factor*=Math.Min(.85,Math.Sqrt((double)encodedLimit/length)*.95);
                }throw new OcrFailure("图片压缩后仍超过所选接口大小限制，未上传。",false);
                }catch {if(File.Exists(target))File.Delete(target);throw;}
            }
        }
    }
    // PDF 统一页码设置的可解析规格：0 不扫描、ALL 全部页、单数字扫描该页、起-止 范围扫描多页。
    // 解析在 UI 线程完成；实际页码在 OCR 阶段对每个 PDF 单独求解。
    internal sealed class PdfPageSpec {
        internal readonly bool All,None;
        internal readonly List<int> ExplicitPages;
        private PdfPageSpec(bool all,bool none,List<int> pages){All=all;None=none;ExplicitPages=pages??new List<int>();}
        internal static PdfPageSpec Parse(string text) {
            string trimmed=(text??"").Trim();
            if(trimmed.Length==0)throw new InvalidOperationException("请填写 PDF 页码设置：0 不扫描、ALL 扫描全部页、单数字扫描该页、或 起始-结束 扫描多页。");
            if(string.Equals(trimmed,"ALL",StringComparison.OrdinalIgnoreCase))return new PdfPageSpec(true,false,null);
            int single;
            if(int.TryParse(trimmed,NumberStyles.Integer,CultureInfo.InvariantCulture,out single)) {
                if(single==0)return new PdfPageSpec(false,true,null);
                if(single<0)throw new InvalidOperationException("PDF 页码必须为正数；0 不扫描、ALL 扫描全部页、或输入 起始-结束。");
                return new PdfPageSpec(false,false,new List<int>{single});
            }
            string[] range=trimmed.Split('-');int rangeFrom,rangeTo;
            if(range.Length==2&&int.TryParse(range[0].Trim(),NumberStyles.Integer,CultureInfo.InvariantCulture,out rangeFrom)
                &&int.TryParse(range[1].Trim(),NumberStyles.Integer,CultureInfo.InvariantCulture,out rangeTo)) {
                if(rangeFrom<1||rangeTo<rangeFrom)throw new InvalidOperationException("PDF 页码范围格式有误：起始页应为正数且不大于结束页。");
                var list=new List<int>();for(int p=rangeFrom;p<=rangeTo;p++)list.Add(p);return new PdfPageSpec(false,false,list);
            }
            throw new InvalidOperationException("无法解析 PDF 页码设置。可输入：0（不扫描）、ALL（全部页）、单数字或 起始-结束 范围。");
        }
        // documentPages<=0 表示页数未知：ALL 回退到第 1 页，范围则信任用户输入不裁剪。
        internal List<int> PagesFor(int documentPages) {
            if(None)return new List<int>();
            if(All){var list=new List<int>();int bound=documentPages>0?documentPages:1;for(int p=1;p<=bound;p++)list.Add(p);return list;}
            var result=new List<int>();
            foreach(int p in ExplicitPages)if(documentPages<=0||p<=documentPages)result.Add(p);
            return result;
        }
        internal string Description {get {if(None)return "0（不扫描）";if(All)return "全部页";return string.Join(",",ExplicitPages);}}
    }
    // 本地求解 PDF 页数：扫描文件中 /Type /Page（排除 /Pages）的出现次数。
    // 加密 PDF 的对象字典 key 不加密，因此此法通常有效；解析失败返回 0。
    internal static class PdfPages {
        private static readonly Regex PagePattern=new Regex(@"/Type\s*/Page(?![sS])",RegexOptions.CultureInvariant);
        internal static int Count(string path) {
            try {
                var info=new FileInfo(path);if(!info.Exists)return 0;
                int limit=(int)Math.Min(info.Length,32*1024*1024);
                using(var stream=File.OpenRead(path)) {
                    var buffer=new byte[limit];int read=0;
                    while(read<limit){int n=stream.Read(buffer,read,limit-read);if(n<=0)break;read+=n;}
                    string text=Encoding.ASCII.GetString(buffer,0,read);
                    return PagePattern.Matches(text).Count;
                }
            }catch{return 0;}
        }
    }
}
