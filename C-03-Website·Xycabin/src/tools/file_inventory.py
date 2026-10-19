"""Conservative file/reference inventory and explicit lifecycle cleanup."""
from pathlib import Path
import hashlib
import json
import re
import shutil
import stat
import ssl
import sys
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src'))
from document_assets import collect as collect_document_assets, referenced_paths
NAMES = ('C-01-ExcelIterate', 'C-02-WordUpdater', 'C-03-Website·Xycabin')
INPUT, PRODUCT, EVIDENCE, HISTORY, UNKNOWN = '必要输入', '当前产物', '有效证据', '可删除历史文件', '未确认文件'
CURRENT_LOGS = {'offline-validation.log','ui-validation.log','installer-payload-verification.log','standalone-validation.log',
 'site-validation.log','shared-source.log','workflow-validation.log','integration-validation.log','devlog-validation.log',
 'audio-controller-validation.log','current-delivery.json','cleanup-report.json','cleanup-validation.log',
 'maintenance-final-rebuild.log','maintenance-second-rebuild.log','maintenance-gates.log','maintenance-integration.log','maintenance-audio-asset.json'}
SECRET = {'.pfx','.p12','.key','.pem','.snk','.pvk','.p8','.jks','.keystore','.der'}

def linked(path):
 info=path.lstat()
 return path.is_symlink() or bool(getattr(info,'st_file_attributes',0)&stat.FILE_ATTRIBUTE_REPARSE_POINT)

def files_without_links(path):
 if linked(path):raise ValueError('不跟随重解析点：'+str(path))
 if path.is_file():return [path]
 result=[]
 for item in sorted(path.iterdir()):result.extend(files_without_links(item))
 return result

def assert_target(base,target):
 base,target=Path(base).absolute(),Path(target).absolute()
 if linked(base):raise ValueError('项目根为重解析点：'+str(base))
 resolved=target.resolve()
 if resolved==base.resolve() or not resolved.is_relative_to(base.resolve()):raise ValueError('清理目标越出项目：'+str(target))
 cursor=target
 while cursor!=base:
  if linked(cursor):raise ValueError('清理路径含重解析点：'+str(cursor))
  cursor=cursor.parent
 return resolved

def digest(path):
 value=hashlib.sha256()
 with Path(path).open('rb') as stream:
  for block in iter(lambda:stream.read(1024*1024),b''):value.update(block)
 return value.hexdigest()

def collect(parent):
 files,opaque=[],[]
 def walk(path,base):
  if linked(path):opaque.append({'path':str(path),'base':str(base),'category':UNKNOWN,'reason':'重解析点，不扫描或删除','references':[],'bytes':0,'sha256':None,'status':'retained'})
  elif path.is_dir():
   for child in sorted(path.iterdir()):walk(child,base)
  else:files.append((path,base))
 for name in NAMES:
  base=Path(parent)/name
  if base.exists():walk(base,base)
 # The parent checker is a required build input, not an optional project artifact.
 checker=Path(parent)/'verify_shared.py'
 if checker.exists():walk(checker,Path(parent))
 return files,opaque

def document_input_paths(files):
 return {asset for base in {base for _,base in files} if base.name in NAMES[:2] and (base.parent/'C-04-Devlog/B1-Project-Files.md').is_file() for asset in referenced_paths(base)}

def input_file(path,base,document_paths=None):
 rel=path.relative_to(base)
 if rel.parts[0]=='AB-Graphic' and path.suffix.lower() in ('.png','.svg') and base.name in NAMES[:2]:
  if document_paths is None:document_paths=set(referenced_paths(base))
  if path.resolve() in document_paths:return True
 return rel.parts[0] in ('src','lib') or (len(rel.parts)==1 and path.suffix.lower() in ('.py','.json')) or (rel.parts[0]=='AA-Program' and path.suffix.lower() in ('.py','.bat')) or path.suffix.lower() in SECRET

def references(files):
 # Positive references protect files; absence is never deletion authorization.
 known={str(p.resolve()).casefold():p for p,_ in files};refs={};document_paths=document_input_paths(files)
 def add(target,source):
  key=str(target.resolve()).casefold()
  if key in known and target.resolve()!=source.resolve():refs.setdefault(key,set()).add(str(source))
 for path,base in files:
  if not input_file(path,base,document_paths) or path.suffix.lower() not in ('.py','.cs','.json','.md','.html','.js','.bat'):continue
  text=path.read_text(encoding='utf-8-sig',errors='replace')
  values=re.findall(r'''["'`]([^"'`\n]{1,260}\.(?:py|cs|dll|exe|json|md|html|js|png|mp3|xml|log|vsto|manifest))["'`]''',text,re.I)
  if path.suffix.lower()=='.md':values += [a or b for a,b in re.findall(r'\]\((?:<([^>]+)>|([^\s)]+))\)',text)]
  for value in values:
   value=value.replace('\\\\','\\').split('#',1)[0]
   if '://' in value:continue
   for origin in (path.parent,base,base/'src/addin',base/'build'):add(origin/value,path)
  if path.name=='project.json':
   config=json.loads(text)
   for value in config.get('sources',[]):add(base/'src/addin'/value,path)
   for tool in config.get('tools',[]):add(base/tool['source'],path);add(base/'build'/(tool['name']+'.exe'),path)
 for base in {base for _,base in files}:
  if base.name not in NAMES[:2] or not (base.parent/'C-04-Devlog/B1-Project-Files.md').is_file():continue
  for document in collect_document_assets(base)['documents']:
   for asset in document['assets']:add(asset['source'],base.parent/document['path'])
 return refs

def lifecycle(path,base):
 rel=path.relative_to(base)
 if rel.parts[0]!='build' or len(rel.parts)<3:return None
 batch=base/'build'/rel.parts[1]
 meta=batch/'fixture-state.json'
 if meta.is_file() and not linked(meta):
  try:
   state=json.loads(meta.read_text(encoding='utf-8-sig'))
   if state.get('status')=='complete' and state.get('experience_merged') is True and state.get('kind') in ('test','install-validation','obsolete-maintenance'):return batch,'批次记录证明已结束且经验已归并'
  except (OSError,ValueError):pass
 return None

def source_fingerprint(files,excluded=()):
 pairs=[];document_paths=document_input_paths(files)
 for path,base in files:
  if input_file(path,base,document_paths) and not any(path.is_relative_to(target) for target in excluded):
   pairs.append((str(path),digest(path)))
 return hashlib.sha256(json.dumps(pairs,ensure_ascii=False,sort_keys=True).encode()).hexdigest()


def public_certificate_bundle(path):
 if path.suffix.lower()!='.pem':return False
 try:
  content=path.read_text(encoding='ascii')
  labels=re.findall(r'-----BEGIN ([A-Z0-9 ]+)-----',content)
  if not labels or set(labels)!={'CERTIFICATE'}:return False
  ssl.create_default_context(cafile=str(path))
  return True
 except (OSError,ValueError,ssl.SSLError):return False


def reviewed_candidates(parent,files,refs):
 """Only measured obsolete build artifacts with a successful excluded-source build qualify."""
 parent=Path(parent).resolve();review_path=parent/NAMES[2]/'build/cleanup-review.json'
 if not review_path.is_file():return {},[]
 review=json.loads(review_path.read_text(encoding='utf-8'))
 if review.get('schema')!=1:raise ValueError('清理核验记录版本不匹配')
 existing=[entry for entry in review['targets'] if Path(entry['path']).exists()]
 if not existing:return {},[]
 proof_path=Path(review['proof'])
 if not proof_path.resolve().is_relative_to((parent/NAMES[2]/'build').resolve()) or digest(proof_path)!=review['proof_sha256']:
  raise ValueError('隔离重建证据缺失、越界或已改变')
 proof=json.loads(proof_path.read_text(encoding='utf-8'))
 if proof.get('status')!='passed' or proof.get('source_fingerprint')!=review['source_fingerprint']:
  raise ValueError('没有本轮成功的隔离重建证明')
 if sorted(proof.get('excluded_targets',[]))!=sorted(entry['path'] for entry in review['targets']):
  raise ValueError('隔离排除范围与清理清单不一致')
 log=Path(proof['build_log'])
 if not log.resolve().is_relative_to((parent/NAMES[2]/'build').resolve()) or digest(log)!=proof['build_log_sha256']:
  raise ValueError('隔离构建原始日志缺失或已改变')
 if not review.get('experience_merged') or not review.get('reason'):
  raise ValueError('缺少用途核验与经验归并记录')
 targets=[Path(entry['path']).resolve() for entry in review['targets']]
 if source_fingerprint(files,targets)!=review['source_fingerprint']:
  raise ValueError('核验后源码、配置或文档已改变，请重新核验')
 approved={}
 for entry in existing:
  base=Path(entry['base']).resolve();target=Path(entry['path']).resolve()
  if base not in [parent/name for name in NAMES]:raise ValueError('核验项目根不匹配')
  assert_target(base,target)
  rel=target.relative_to(base)
  if rel.parts[0]!='build' or len(rel.parts)<2:raise ValueError('专项清理仅接受项目 build 下的历史资料')
  current={base/'build'/name for name in ('Excel_Iterate.dll','WordUpdater.dll','Interop.Microsoft.Office.Core.dll','Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
           'Excel_Iterate.dll.manifest','Excel_Iterate.vsto','WordUpdater.dll.manifest','WordUpdater.vsto','InstallerAssemblyInfo.g.cs','PayloadProbe.exe','site-embedded.snapshot.html')}
  current|={base/'build'/(p.stem+'.exe') for p in (base/'src/tools').glob('*.cs')}
  config=base/'project.json'
  if config.is_file():current|={base/'build'/(t['name']+'.exe') for t in json.loads(config.read_text(encoding='utf-8-sig')).get('tools',[])}
  members=files_without_links(target)
  if {str(file.relative_to(target)) if target.is_dir() else file.name:digest(file) for file in members}!=entry['identities']:
   raise ValueError('核验目标内容已变化：'+str(target))
  for file in members:
   parts=file.relative_to(base).parts
   if file in current or file.name in CURRENT_LOGS or file.name.startswith(('.EI_Operation_', '.EI_Link_')) or any('recovery' in part.casefold() for part in parts):
    raise ValueError('目标含当前产物、证据或未确认恢复资料：'+str(file))
   if refs.get(str(file.resolve()).casefold()):raise ValueError('目标仍有现行引用：'+str(file))
   if file.suffix.lower() in SECRET and not public_certificate_bundle(file):raise ValueError('目标含未确认签名/密钥资料：'+str(file))
  approved[target]=(base,'用户明确废弃；用途/经验核验及排除候选从零构建通过')
 return approved,targets


def scan(parent):
 files,records=collect(parent);refs=references(files);document_paths=document_input_paths(files)
 reviewed,excluded=reviewed_candidates(parent,files,refs)
 tools={}
 for name in NAMES:
  base=Path(parent)/name
  tools[name]={p.stem+'.exe' for p in (base/'src/tools').glob('*.cs')}|{'PayloadProbe.exe'}
  config=base/'project.json'
  if config.is_file() and not linked(config):tools[name]|={t['name']+'.exe' for t in json.loads(config.read_text(encoding='utf-8-sig')).get('tools',[])}
 for path,base in files:
  rel=path.relative_to(base);uses=sorted(refs.get(str(path.resolve()).casefold(),[]));category,reason=UNKNOWN,'用途或生命周期未确认，保留';old=lifecycle(path,base)
  if any(path.is_relative_to(target) for target in reviewed):category,reason=HISTORY,'本轮用途核验、经验归并及隔离从零重建已通过'
  elif input_file(path,base,document_paths):category,reason=INPUT,'现行源码、配置、文档、依赖或签名资料'
  elif any('recovery' in part.casefold() for part in rel.parts) or path.name.startswith(('.EI_Operation_', '.EI_Link_')):category,reason=EVIDENCE,'可能仍用于恢复；未核验还原完成及确认，不参与历史清理'
  elif uses:category,reason=(EVIDENCE if path.suffix.lower() in ('.log','.json','.txt','.png') else PRODUCT),'现行源码或文档存在引用'
  elif rel.parts[0] in ('AA-Program','AB-Graphic'):category,reason=PRODUCT,'正式交付分区中的当前文件'
  elif path.name in CURRENT_LOGS or re.fullmatch(r'(?:Excel_Iterate|WordUpdater)-(?:rebuild|validation|payload|ui-validation)\.log',path.name):category,reason=EVIDENCE,'现行验证或交付记录'
  elif old:category,reason=HISTORY,old[1]
  elif rel.parts[0]=='build' and (path.suffix.lower() in ('.dll','.vsto','.manifest') or path.name in tools[base.name] or path.name=='InstallerAssemblyInfo.g.cs'):category,reason=PRODUCT,'当前编译载荷或现行工具'
  records.append({'path':str(path),'base':str(base),'category':category,'reason':reason,'references':uses,'bytes':path.stat().st_size,'sha256':digest(path)})
 candidates=dict(reviewed);bypath={r['path']:r for r in records}
 for path,base in files:
  old=lifecycle(path,base)
  if old:candidates[old[0]]=(base,old[1])
 remove=[]
 for path,(base,reason) in sorted(candidates.items()):
  try:
   assert_target(base,path);members=files_without_links(path)
   if any(bypath[str(p)]['category']!=HISTORY for p in members):
    for p in members:
     if bypath[str(p)]['category']==HISTORY:bypath[str(p)].update(category=UNKNOWN,reason='批次包含受保护资料或现行引用，整批保留')
    continue
   remove.append({'path':str(path),'base':str(base),'category':HISTORY,'reason':reason,'files':len(members),'bytes':sum(p.stat().st_size for p in members),'identities':{str(p.relative_to(path)) if path.is_dir() else p.name:digest(p) for p in members},'status':'pending'})
  except (OSError,ValueError):continue
 retain=[r for r in records if r['category']!=HISTORY]
 for record in records:record['status']='pending' if record['category']==HISTORY else 'retained'
 fingerprint=source_fingerprint(files,excluded)
 return records,remove,retain,fingerprint

def inventory(parent):
 _,remove,retain,_=scan(parent);return remove,retain

def apply(entries,parent=None,expected_sources=None):
 parent=Path(parent or ROOT.parent).resolve();_,current,_,fingerprint=scan(parent)
 if expected_sources is not None and fingerprint!=expected_sources:raise ValueError('源码、配置或文档改变，请重新dry-run并审阅')
 allowed={e['path']:e for e in current}
 def verify(entry):
  if Path(entry['base']).resolve() not in {parent/n for n in NAMES}:raise ValueError('清单项目根不匹配：'+entry['base'])
  p=Path(entry['path']);assert_target(entry['base'],p);members=files_without_links(p);now={str(f.relative_to(p)) if p.is_dir() else f.name:digest(f) for f in members}
  if now!=entry['identities']:raise ValueError('清单后内容改变，停止清理：'+str(p))
  if str(p) not in allowed or allowed[str(p)]['identities']!=now:raise ValueError('分类、引用或批次状态已变化，保留：'+str(p))
 for entry in entries:verify(entry)
 for entry in entries:
  try:
   verify(entry)
   target=Path(entry['path'])
   if target.is_dir():shutil.rmtree(target)
   else:target.unlink()
   entry['status']='deleted'
  except (OSError,ValueError) as error:entry['status']='failed';entry['error']=str(error);raise
