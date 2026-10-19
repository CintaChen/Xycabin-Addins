"""Generate a five-category inventory; apply only a reviewed dry-run manifest."""
import argparse
import json
import sys
from pathlib import Path
sys.dont_write_bytecode=True
from file_inventory import ROOT, INPUT, PRODUCT, EVIDENCE, HISTORY, UNKNOWN, scan, inventory, apply, assert_target

# rule=R-23
def main():
 parser=argparse.ArgumentParser(description=__doc__)
 mode=parser.add_mutually_exclusive_group(required=True)
 mode.add_argument('--dry-run',action='store_true');mode.add_argument('--apply',action='store_true')
 args=parser.parse_args();path=ROOT/'build/cleanup-report.json'
 if args.dry_run:
  records,remove,retain,fingerprint=scan(ROOT.parent)
  report={'schema':3,'mode':'dry-run','completed':False,'inventory':records,'remove':remove,'retain':retain,'source_fingerprint':fingerprint,'files':sum(e['files'] for e in remove),'bytes':sum(e['bytes'] for e in remove)}
 else:
  report=json.loads(path.read_text(encoding='utf-8-sig'))
  if report.get('schema')!=3 or report.get('mode')!='dry-run':raise ValueError('先运行--dry-run并审阅本轮清单，再执行--apply')
  report['mode']='apply'
  try:apply(report['remove'],ROOT.parent,report['source_fingerprint']);report['completed']=True
  except (OSError,ValueError) as error:report['error']=str(error);raise
  finally:
   for record in report['inventory']:
    for entry in report['remove']:
     if Path(record['path']).is_relative_to(Path(entry['path'])):record['status']=entry['status'];break
   path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
 path.parent.mkdir(exist_ok=True);path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
 print(json.dumps({c:sum(r['category']==c for r in report['inventory']) for c in (INPUT,PRODUCT,EVIDENCE,HISTORY,UNKNOWN)},ensure_ascii=False))
 print('合计：%d候选文件，%d字节；只在用户要求清理时执行--apply'%(report['files'],report['bytes']))

if __name__=='__main__':
 try:main()
 except (OSError,ValueError) as error:print('FAIL: '+str(error));sys.exit(1)
