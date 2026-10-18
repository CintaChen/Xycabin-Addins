using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_Iterate
{
    internal enum OcrMode { General, Accurate, GeneralBasic, AccurateBasic, AliyunAdvanced }
    internal static class OcrModes
    {
        internal static readonly string[] BaiduNames={"标准含位置版","高精度含位置版","标准版","高精度版"};
        internal static string Method(OcrMode mode) {switch(mode){case OcrMode.General:return "general";case OcrMode.Accurate:return "accurate";case OcrMode.GeneralBasic:return "general_basic";case OcrMode.AccurateBasic:return "accurate_basic";case OcrMode.AliyunAdvanced:return "RecognizeAdvanced";default:throw new ArgumentOutOfRangeException("mode");}}
        internal static bool Accurate(OcrMode mode){return mode==OcrMode.Accurate||mode==OcrMode.AccurateBasic;}
        internal static bool Positioned(OcrMode mode){return mode==OcrMode.General||mode==OcrMode.Accurate;}
        internal static int Longest(OcrMode mode){return mode==OcrMode.AliyunAdvanced?8191:Accurate(mode)?8192:4096;}
        internal static int Limit(OcrMode mode){return (mode==OcrMode.AliyunAdvanced||Accurate(mode)?10:8)*1024*1024;}
    }
    [DataContract] internal sealed class AliyunReply
    {
        [DataMember] internal string Data;
        [DataMember] internal string RequestId;
        [DataMember] internal string Code;
    }
    [DataContract] internal sealed class AliyunData
    {
        [DataMember(Name="prism_rowsInfo")] internal List<AliyunRow> Rows;
        [DataMember(Name="prism_wordsInfo")] internal List<AliyunWord> Words;
    }
    [DataContract] internal sealed class AliyunRow
    {
        [DataMember(Name="rowId")] internal int Id;
        [DataMember(Name="word")] internal string Text;
    }
    [DataContract] internal sealed class AliyunWord
    {
        [DataMember(Name="word")] internal string Text;
        [DataMember(Name="pos")] internal List<AliyunPoint> Points;
        [DataMember(Name="prob")] internal double? Score;
    }
    [DataContract] internal sealed class AliyunPoint
    {
        [DataMember(Name="x")] internal double X;
        [DataMember(Name="y")] internal double Y;
    }
    internal sealed class AliyunFileOcr:IFileOcr
    {
        internal const string Host="ocr-api.cn-hangzhou.aliyuncs.com";
        private readonly string key,secret;
        private readonly HttpClient http;
        internal AliyunFileOcr(string key,string secret):this(key,secret,null){}
        internal AliyunFileOcr(string key,string secret,HttpMessageHandler handler)
        {
            if(string.IsNullOrWhiteSpace(key)||string.IsNullOrWhiteSpace(secret))throw new OcrFailure("请填写阿里云的 AccessKey ID 和 AccessKey Secret，并确认账号已开通文字识别服务及调用权限。",true);
            this.key=key;this.secret=secret;http=handler==null?new HttpClient():new HttpClient(handler);http.Timeout=TimeSpan.FromSeconds(120);
        }
        private static string Hex(byte[] bytes){return BitConverter.ToString(bytes).Replace("-","").ToLowerInvariant();}
        internal static string Hash(byte[] bytes){using(var sha=SHA256.Create())return Hex(sha.ComputeHash(bytes));}
        internal static string Authorization(string key,string secret,string query,byte[] body,string date,string nonce)
        {
            const string signed="content-type;host;x-acs-action;x-acs-content-sha256;x-acs-date;x-acs-signature-nonce;x-acs-version";
            string payload=Hash(body);
            string headers="content-type:application/octet-stream\nhost:"+Host+"\nx-acs-action:RecognizeAdvanced\nx-acs-content-sha256:"+payload+"\nx-acs-date:"+date+"\nx-acs-signature-nonce:"+nonce+"\nx-acs-version:2021-07-07\n";
            string canonical="POST\n/\n"+query+"\n"+headers+"\n"+signed+"\n"+payload;
            string toSign="ACS3-HMAC-SHA256\n"+Hash(Encoding.UTF8.GetBytes(canonical));
            using(var hmac=new HMACSHA256(Encoding.UTF8.GetBytes(secret)))return "ACS3-HMAC-SHA256 Credential="+key+",SignedHeaders="+signed+",Signature="+Hex(hmac.ComputeHash(Encoding.UTF8.GetBytes(toSign)));
        }
        internal static HttpRequestMessage Request(string key,string secret,byte[] bytes,string date,string nonce)
        {
            // These RPC options are query parameters. The body contains only image bytes.
            const string query="NeedRotate=true&NeedSortPage=true&Row=true";
            var request=new HttpRequestMessage(HttpMethod.Post,"https://"+Host+"/?"+query);
            request.Content=new ByteArrayContent(bytes);request.Content.Headers.ContentType=new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
            request.Headers.Host=Host;request.Headers.Add("x-acs-action","RecognizeAdvanced");request.Headers.Add("x-acs-version","2021-07-07");request.Headers.Add("x-acs-date",date);request.Headers.Add("x-acs-signature-nonce",nonce);request.Headers.Add("x-acs-content-sha256",Hash(bytes));request.Headers.TryAddWithoutValidation("Authorization",Authorization(key,secret,query,bytes,date,nonce));request.Headers.Accept.ParseAdd("application/json");return request;
        }
        internal static List<OcrSegment> Parse(string json,int page)
        {
            var data=OcrJson.Read<AliyunData>(Encoding.UTF8.GetBytes(json));var output=new List<OcrSegment>();int sequence=0;
            if(data.Rows!=null&&data.Rows.Any(r=>r!=null&&!string.IsNullOrWhiteSpace(r.Text)))
            {
                foreach(var row in data.Rows.Where(r=>r!=null&&!string.IsNullOrWhiteSpace(r.Text)))output.Add(new OcrSegment {Text=row.Text,Page=page,NoLocation=true,LineId=row.Id.ToString(CultureInfo.InvariantCulture),Sequence=sequence++});
                return output;
            }
            if(data.Words!=null)foreach(var word in data.Words.Where(w=>w!=null&&!string.IsNullOrWhiteSpace(w.Text)))
            {
                var points=word.Points;
                if(points==null||points.Count!=4){output.Add(new OcrSegment {Text=word.Text,Page=page,NoLocation=true,Sequence=sequence++});continue;}
                double x=points.Min(p=>p.X),y=points.Min(p=>p.Y),width=points.Max(p=>p.X)-x,height=points.Max(p=>p.Y)-y;
                output.Add(new OcrSegment {Text=word.Text,Page=page,Left=x,Top=y,Width=width,Height=height,NoLocation=width<=0||height<=0,Score=word.Score.HasValue?(double?)(word.Score.Value/100):null,Sequence=sequence++});
            }
            if(data.Rows==null&&data.Words==null)throw new OcrFailure("阿里云未返回有效的行或文字块结果。",false);
            return output;
        }
        internal static bool Pauses(string code)
        {
            string value=(code??"").ToLowerInvariant();return value.Contains("accesskey")||value.Contains("signature")||value.Contains("permission")||value.Contains("forbidden")||value.Contains("unauthorized")||value.Contains("quota")||value.Contains("throttl")||value.Contains("balance")||value.Contains("arrear")||value.Contains("unpaid")||value.Contains("notpurchased")||value.Contains("notenabl")||value.Contains("notopened");
        }
        public async Task<List<OcrSegment>> RecognizeAsync(string image,int page,CancellationToken cancel)
        {
            var requested=cancel;using(var deadline=CancellationTokenSource.CreateLinkedTokenSource(cancel))
            {
                deadline.CancelAfter(TimeSpan.FromSeconds(120));cancel=deadline.Token;
                try
                {
                    cancel.ThrowIfCancellationRequested();
                    // RecognizeAdvanced documents no PDF support; images are the only accepted body.
                    if(OcrImages.IsPdf(image))throw new OcrFailure("阿里云全文识别高精版不支持 PDF，请改用百度云识别 PDF 文件。",false);
                    var bytes=File.ReadAllBytes(image);if(bytes.Length>10*1024*1024)throw new OcrFailure("图片超过阿里云 10MB 限制，未上传。",false);
                    using(var request=Request(key,secret,bytes,DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ",CultureInfo.InvariantCulture),Guid.NewGuid().ToString("N")))
                    using(var response=await http.SendAsync(request,cancel).ConfigureAwait(false))
                    {
                        AliyunReply reply=null;try{reply=OcrJson.Read<AliyunReply>(await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false));}catch{}
                        string code=reply==null?"InvalidResponse":reply.Code;
                        OcrDiagnostics.Cloud("recognize","aliyun","RecognizeAdvanced",string.IsNullOrEmpty(code)?((int)response.StatusCode).ToString():code,reply==null?null:reply.RequestId);
                        if(!response.IsSuccessStatusCode||!string.IsNullOrEmpty(code))throw new OcrFailure("阿里云识别失败（HTTP "+(int)response.StatusCode+"，错误码 "+OcrDiagnostics.Safe(code)+"），请检查服务权限、配额、密钥或图片后重试。",Pauses(code)||new[]{401,403,429}.Contains((int)response.StatusCode));
                        if(reply==null||string.IsNullOrWhiteSpace(reply.Data))throw new OcrFailure("阿里云识别结果无法读取。",false);
                        try{return Parse(reply.Data,page);}catch(OcrFailure){throw;}catch{throw new OcrFailure("阿里云识别结果格式无效。",false);}
                    }
                }
                catch(OperationCanceledException){if(requested.IsCancellationRequested)throw;OcrDiagnostics.Cloud("request","aliyun","RecognizeAdvanced","Timeout",null);throw new OcrFailure("识别超时；未自动重复请求，请手动重试。",false);}
                catch(HttpRequestException){OcrDiagnostics.Cloud("request","aliyun","RecognizeAdvanced","Network",null);throw new OcrFailure("阿里云网络请求失败，请检查网络连接后手动重试。",false);}
            }
        }
        public void Dispose(){http.Dispose();}
    }
}
