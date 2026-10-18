"use strict";
const fs=require("fs"),vm=require("vm"),crypto=require("crypto"),path=require("path");
const source=fs.readFileSync(path.join(__dirname,"../site_downloads.js"),"utf8");
let passed=0;
function check(ok,label){if(!ok)throw new Error(label);passed++}
function fixture(packs={},channels={word:{baidu:"",quark:""},excel:{baidu:"",quark:""}}){
 const nodes={},events={},clicks=[],notices=[],timers=[],revoked=[],blobs=[];let current="word",fail=false,decodeCalls=0,maxDecode=0;
 const doc={activeElement:null,documentElement:{clientWidth:360},addEventListener:(k,f)=>events[k]=f,getElementById:k=>nodes[k],createElement:k=>element(k),body:null};
 function element(tag){return {tag,hidden:true,disabled:false,children:[],dataset:{},attrs:{},listeners:{},textContent:"",setAttribute(k,v){this.attrs[k]=v},addEventListener(k,f){this.listeners[k]=f},appendChild(n){this.children.push(n);n.parent=this},replaceChildren(){this.children=[]},remove(){if(this.parent)this.parent.children=this.parent.children.filter(n=>n!==this)},focus(){doc.activeElement=this},closest(k){return this.inside?k===".download-split"?this:null:null},click(){if(this.tag==="a"){if(fail)throw new Error("failed");clicks.push({href:this.href,download:this.download,target:this.target,rel:this.rel})}else if(!this.disabled&&this.listeners.click)this.listeners.click({target:this})}}}
 doc.body=element("body");["download-main","download-toggle","download-menu","installer-assets"].forEach(k=>{nodes[k]=element(k);nodes[k].inside=true});nodes["installer-assets"].textContent=JSON.stringify({packages:packs});
 const bounds={left:-97,right:148};nodes["download-menu"].style={};nodes["download-menu"].getBoundingClientRect=()=>bounds;
 const APIURL=class extends URL{};APIURL.createObjectURL=b=>{blobs.push(b);return "blob:test/"+blobs.length};APIURL.revokeObjectURL=u=>revoked.push(u);
 const win={URL:APIURL,Blob,atob:(value)=>{decodeCalls++;maxDecode=Math.max(maxDecode,value.length);return atob(value)},addEventListener:(k,f)=>events["window:"+k]=f,setTimeout:(f,ms)=>{timers.push(f);check(ms===60000,"object URL has delayed cleanup")}};
 const text={download:"点我下载",download_methods:"选择下载方式",download_default:"默认下载",download_baidu:"百度网盘",download_quark:"夸克网盘",download_unavailable_short:"此页面不可用",download_missing:"未配置",download_unconfigured:"未配置",download_unavailable:"不可用",download_requested:"已请求",download_failure:"失败"};
 const context={Uint8Array,Array,Number};vm.createContext(context);vm.runInContext(source,context);
 const api=context.createCabinDownloads({text,products:[{id:"word",name:"WordUpdater"},{id:"excel",name:"Excel-Iterate"}],downloads:channels},doc,win,()=>current,m=>notices.push(m));
 return {nodes,doc,events,clicks,notices,timers,revoked,blobs,api,bounds,get decodeCalls(){return decodeCalls},get maxDecode(){return maxDecode},set(value){current=value;api.render()},fail(){fail=true}};
}
function pack(name,body){return {filename:name+"_Setup.exe",version:"0.9.9.0",bytes:body.length,sha256:crypto.createHash("sha256").update(body).digest("hex"),base64:body.toString("base64")}}
(async()=>{
 const word=Buffer.from("MZword-exe"),excel=Buffer.from("MZexcel-exe"),assets={word:pack("WordUpdater",word),excel:pack("Excel_Iterate",excel)};
 let f=fixture(assets);check(f.nodes["download-menu"].children.length===3,"three download choices");check(!f.nodes["download-menu"].children[0].disabled&&f.nodes["download-menu"].children[1].disabled,"asset default enabled, empty cloud disabled");
 f.nodes["download-main"].click();check(f.clicks[0].download==="WordUpdater_Setup.exe","Word project filename");check(Buffer.from(await f.blobs[0].arrayBuffer()).equals(word),"Word Blob bytes exact");check(f.notices[0]==="已请求","request does not claim download saved");
 f.nodes["download-toggle"].click();check(!f.nodes["download-menu"].hidden&&f.doc.activeElement===f.nodes["download-menu"].children[0],"menu opens and focuses enabled choice");
 check(f.nodes["download-menu"].style.transform==="translateX(105px)","left-overflowing mobile menu moves inside viewport");
 f.bounds.left=200;f.bounds.right=445;f.events["window:resize"]();check(f.nodes["download-menu"].style.transform==="translateX(-93px)","right overflow is corrected on resize");
 f.bounds.left=20;f.bounds.right=265;f.events["window:resize"]();check(f.nodes["download-menu"].style.transform==="translateX(0px)","resize removes old correction when menu fits");
 check(f.nodes["download-menu"].style.maxWidth==="344px","menu width leaves viewport edge space");
 f.set("excel");check(f.nodes["download-menu"].hidden&&f.nodes["download-toggle"].attrs["aria-expanded"]==="false","project switch closes menu");f.nodes["download-main"].click();check(f.clicks[1].download==="Excel_Iterate_Setup.exe","Excel switch cannot download stale Word");check(Buffer.from(await f.blobs[1].arrayBuffer()).equals(excel),"Excel Blob bytes exact");
 f.timers[0]();check(f.revoked.length===1,"delayed URL release");f.events["window:pagehide"]();check(f.revoked.length===2,"remaining URLs released on pagehide");
 f=fixture({}, {word:{baidu:"https://pan.baidu.com/s/test?pwd=1234",quark:"https://pan.quark.cn/s/test"},excel:{baidu:"",quark:"https://pan.quark.cn/s/excel"}});
 check(f.nodes["download-menu"].children[0].disabled,"embedded site default extraction unavailable");f.nodes["download-main"].click();check(f.clicks[0].href.includes("baidu")&&f.clicks[0].target==="_blank"&&f.clicks[0].rel==="noopener noreferrer","embedded main uses Baidu safely");
 f.set("excel");f.nodes["download-main"].click();check(f.clicks[1].href.includes("/excel"),"Quark fallback uses selected project");f.set("word");f.nodes["download-menu"].children[2].click();check(f.clicks[2].href.includes("quark"),"explicit Quark choice");
 f.nodes["download-toggle"].click();let prevented=false;f.events.keydown({key:"End",target:f.nodes["download-toggle"],preventDefault(){prevented=true}});check(prevented&&f.doc.activeElement===f.nodes["download-menu"].children[2],"keyboard End selection");
 f.events.keydown({key:"Escape",target:f.doc.activeElement,preventDefault(){}});check(f.nodes["download-menu"].hidden&&f.doc.activeElement===f.nodes["download-toggle"],"Esc closes and returns focus");f.nodes["download-toggle"].click();f.events.click({target:{closest(){return null}}});check(f.nodes["download-menu"].hidden,"outside click closes");
 f=fixture();f.nodes["download-main"].click();check(f.notices[0]==="未配置"&&f.clicks.length===0,"missing links tell user");
 for(const url of ["http://example.org/","javascript:alert(1)","https://user:pass@example.org/","https://example.org/\n"]){f=fixture({}, {word:{baidu:url,quark:""},excel:{baidu:"",quark:""}});f.nodes["download-main"].click();check(f.clicks.length===0,"unsafe cloud URL rejected")}
 f=fixture(assets);f.fail();f.nodes["download-main"].click();check(f.notices[0]==="失败"&&!f.nodes["download-main"].disabled,"failed browser request leaves controls usable");
 check(f.revoked.length===1&&f.doc.body.children.length===0,"failed request releases temporary URL and anchor immediately");
 f=fixture({word:{...assets.word,bytes:999}});f.nodes["download-main"].click();check(f.notices[0]==="失败"&&f.clicks.length===0,"damaged asset rejected");
 const large=crypto.randomBytes(1024*1024+7);f=fixture({excel:pack("Excel_Iterate",large)});f.set("excel");
 for(let i=0;i<2;i++){f.nodes["download-main"].click();const extracted=Buffer.from(await f.blobs[i].arrayBuffer());check(extracted.equals(large),"large package repeated download bytes exact "+i);check(crypto.createHash("sha256").update(extracted).digest("hex")===crypto.createHash("sha256").update(large).digest("hex"),"large package repeated digest exact "+i)}
 check(f.decodeCalls>2&&f.maxDecode<=87380,"Base64 decode is bounded per chunk");f.events["window:pagehide"]();check(f.revoked.length===2,"repeated large download resources released");
 if(process.argv[2]){
  const html=fs.readFileSync(process.argv[2],"utf8"),data=JSON.parse(html.match(/<script id="installer-assets" type="application\/json">(.*?)<\/script>/s)[1]);
  for(const [id,p] of Object.entries(data.packages)){const bytes=Buffer.from(p.base64,"base64");check(bytes.length===p.bytes&&crypto.createHash("sha256").update(bytes).digest("hex")===p.sha256,"delivered "+id+" hash exact");const real=fixture(data.packages);real.set(id);for(let i=0;i<2;i++){real.nodes["download-main"].click();check(Buffer.from(await real.blobs[i].arrayBuffer()).equals(bytes),"delivered "+id+" repeated extraction exact "+i)}check(real.maxDecode<=87380,"delivered chunk bound "+id);real.events["window:pagehide"]();check(real.revoked.length===2,"delivered repeated resources freed "+id)}
 }
 console.log("DownloadProbe: "+passed+" PASS, 0 FAIL");
})().catch(e=>{console.error(e);process.exit(1)});
