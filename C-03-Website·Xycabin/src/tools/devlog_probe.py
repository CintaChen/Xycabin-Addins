"""Website development-log hierarchy, handoff markers and experience ownership."""
from pathlib import Path
import re
import sys
sys.dont_write_bytecode = True
from markdown_gate import hierarchy, verify_guide, verify_entry, verify_local_anchors, self_test

ROOT = Path(__file__).resolve().parents[2]


def verify():
    path = ROOT / 'AP3-说明/开发日志.md'
    text = path.read_text(encoding='utf-8-sig')
    expected = ['一. Agent及开发者接手概览','二. 项目档案','三. XyCabin项目维护纪律','四. XyCabin专属开发经验']
    errors, titles = hierarchy(text, str(path))
    if titles != expected:
        raise ValueError('网站开发日志必须保留四个编号章节')
    self_test()
    verify_local_anchors(path)
    if errors: raise ValueError('\n'.join(errors))
    verify_guide(path.with_name('使用说明.md'))
    verify_entry(path.with_name('AGENTS.md'))
    for name in ('agent-overview','delivery-audit'):
        a,b = '<!-- '+name+':start -->','<!-- '+name+':end -->'
        if text.count(a)!=1 or text.count(b)!=1 or text.index(a)>=text.index(b):
            raise ValueError('文档标记必须唯一. 有序：' + name)
    if len(re.findall(r'^最后更新：\d{4}-\d{2}-\d{2}',text,re.M)) != 1:
        raise ValueError('顶部必须有唯一更新日期')
    if any(s in text for s in ('Excel窗口四档','Word1139','公式事务','两DLL内嵌网页、三个输出预览')):
        raise ValueError('宿主业务或过期交付描述应归入对应加载项日志')
    print('PASS: Website开发日志四部分、三份文档五级编号及本地标题锚点、概览/审计标记及经验归属',flush=True)


if __name__=='__main__':
    try:
        verify()
    except (OSError,ValueError,AssertionError) as error:
        print('FAIL: '+str(error),flush=True)
        sys.exit(1)
