"""Build a self-contained, offline website from C3-XycabinSite_Configure.py."""
from pathlib import Path
import copy
import base64
import html
import importlib.util
import json
import math
import hashlib
import re
import sys
from urllib.parse import urlsplit
from configurators import CONFIGURE, BY_PRODUCT
sys.dont_write_bytecode = True


def load(path):
    spec = importlib.util.spec_from_file_location("_xy_site_config", path)
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module


def payload(config, root):
    data = {key.lower(): copy.deepcopy(config[key]) for key in ("SITE", "TEXT", "PRODUCTS", "WORKFLOW", "FUTURE", "SOUND", "BACKGROUND", "DOWNLOAD_GUIDE", "FEEDBACK", "DOWNLOADS")}
    if not data["products"] or not data["workflow"] or not data["future"]:
        raise ValueError("网站的项目、开发流程和未来计划不能为空")
    sound = data["sound"]
    if type(sound["enabled"]) is not bool or sound["waveform"] not in ("sine", "triangle", "square", "sawtooth"):
        raise ValueError("SOUND.enabled 须为布尔值，waveform 须为浏览器支持的波形")
    for key, low, high in (("frequency_hz", 20, 20000), ("duration_ms", 10, 500), ("gain", 0, 1), ("max_gain", 0.01, 1)):
        value = sound[key]
        if type(value) not in (int, float) or not math.isfinite(value) or not low <= value <= high:
            raise ValueError("SOUND." + key + " 超出有效范围")
    if sound["gain"] > sound["max_gain"]:
        raise ValueError("SOUND.gain 不可大于 max_gain")
    background = data["background"]
    value = background["volume"]
    if type(value) not in (int, float) or not math.isfinite(value) or not 0 <= value <= 1:
        raise ValueError("BACKGROUND.volume 须为0～1的有限数值")
    value = background["max_gain"]
    if type(value) not in (int, float) or not math.isfinite(value) or not 0 < value <= 1:
        raise ValueError("BACKGROUND.max_gain 须为大于0且不超过1的有限数值")
    if type(background["autoplay"]) is not bool:
        raise ValueError("BACKGROUND.autoplay 须为布尔值")
    background.setdefault("play_on_first_interaction", True)
    if type(background["play_on_first_interaction"]) is not bool:
        raise ValueError("BACKGROUND.play_on_first_interaction 须为布尔值")
    scenes = background["scenes"]
    if not isinstance(scenes, list) or not scenes:
        raise ValueError("BACKGROUND.scenes 至少须保留一种音景")
    ids = set()
    for scene in scenes:
        if scene["id"] not in ("afternoon", "tranquility") or scene["id"] in ids:
            raise ValueError("背景音景id须为已支持且不重复的算法名称")
        ids.add(scene["id"])
        if any(not isinstance(scene[k], str) or not scene[k].strip() for k in ("name", "description")):
            raise ValueError("音景名称与描述须为非空字符串")
        if scene["id"] == "tranquility":
            relative = scene.get("audio_file")
            if not isinstance(relative, str) or not relative or Path(relative).is_absolute():
                raise ValueError("audio_file 必须是网站项目内的相对MP3路径")
            asset = (root / relative).resolve()
            if not asset.is_relative_to(root.resolve()) or asset.suffix.lower() != ".mp3" or not asset.is_file():
                raise ValueError("网站音频资产缺失或越出项目：" + relative)
            scene["audio_base64"] = base64.b64encode(asset.read_bytes()).decode("ascii")
    if background["default_scene"] not in ids:
        raise ValueError("默认背景音必须在scenes中")
    guide = data["download_guide"]
    if not isinstance(guide, dict) or not isinstance(guide.get("steps"), list) or not guide["steps"]:
        raise ValueError("DOWNLOAD_GUIDE.steps 须为非空步骤列表")
    if any(not isinstance(step, str) or not step.strip() for step in guide["steps"]):
        raise ValueError("下载指引步骤须为非空字符串")
    for key in ("image_file", "image_alt"):
        if not isinstance(guide.get(key), str) or not guide[key].strip():
            raise ValueError("DOWNLOAD_GUIDE." + key + " 须为非空字符串")
    relative = guide["image_file"]
    if Path(relative).is_absolute():
        raise ValueError("image_file 必须是网站项目内的相对PNG路径")
    image = (root / relative).resolve()
    if not image.is_relative_to(root.resolve()) or image.suffix.lower() != ".png" or not image.is_file():
        raise ValueError("下载指引图片缺失或越出项目：" + relative)
    guide["image_base64"] = base64.b64encode(image.read_bytes()).decode("ascii")
    for key in ("download_guide_title", "download_guide_lead", "download_guide_continue", "download_guide_cancel"):
        value = data["text"].get(key)
        if not isinstance(value, str) or not value.strip():
            raise ValueError("TEXT." + key + " 须为非空字符串（下载指引弹窗文案）")
    feedback = data["feedback"]
    if any(not isinstance(feedback[key], str) for key in ("title", "intro", "subject", "hint")):
        raise ValueError("问题反馈文案须为字符串")
    if not isinstance(feedback["emails"], list) or not feedback["emails"]:
        raise ValueError("问题反馈至少需要一个邮箱")
    for email in feedback["emails"]:
        if not isinstance(email, str) or not re.fullmatch(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", email):
            raise ValueError("无效反馈邮箱：" + str(email))
    urls = feedback.get("webmail_urls", {})
    if not isinstance(urls, dict):
        raise ValueError("FEEDBACK.webmail_urls 须为邮箱域名到 HTTPS 官网的映射")
    for domain, url in urls.items():
        if not isinstance(domain, str) or domain != domain.lower() or not re.fullmatch(r"[a-z0-9.-]+\.[a-z]{2,}", domain):
            raise ValueError("官网映射须使用小写邮箱域名")
        if not isinstance(url, str) or re.search(r"[\s\x00-\x1f]", url):
            raise ValueError("官网地址不可含空白或控制字符")
        parsed = urlsplit(url)
        if parsed.scheme != "https" or not parsed.hostname or parsed.username is not None or parsed.password is not None:
            raise ValueError("邮箱官网只接受无账号密码的 HTTPS 地址")
    downloads=data["downloads"]
    if not isinstance(downloads,dict) or set(downloads)!={"word","excel"}:
        raise ValueError("DOWNLOADS 必须分别配置 word/excel")
    for identity, channels in downloads.items():
        if not isinstance(channels,dict) or set(channels)!={"baidu","quark"}:
            raise ValueError("下载方式必须为 baidu/quark")
        for value in channels.values():
            if not isinstance(value,str): raise ValueError("下载链接须为字符串")
            if not value: continue
            parsed=urlsplit(value)
            if re.search(r"[\s\x00-\x1f]",value) or parsed.scheme!="https" or not parsed.hostname or parsed.username is not None or parsed.password is not None:
                raise ValueError("网盘只接受无账号密码的 HTTPS 地址")
    identities = set()
    project_paths = {"excel": root.parent / "C-01-ExcelIterate", "word": root.parent / "C-02-WordUpdater"}
    for product in data["products"]:
        identity = product["id"]
        if not re.fullmatch(r"[a-z][a-z0-9_-]*", identity) or identity in identities:
            raise ValueError("项目 id 必须唯一且为小写字母、数字或下划线/连字符")
        identities.add(identity)
        program = project_paths.get(identity, root / "__absent__") / "AA-Program" / BY_PRODUCT.get(identity, "__absent__.py")
        if program.is_file():
            version = load(program).get_version()
            if product["versions"][0]["version"] == product["version"]:
                product["versions"][0]["version"] = version
            product["version"] = version
        for version in [product["version"]] + [v["version"] for v in product["versions"]]:
            if not re.fullmatch(r"\d+\.\d+\.\d+\.\d+", version):
                raise ValueError("网站版本须为四段数字")
        if not product["guides"] or not product["features"]:
            raise ValueError("每个项目至少需要一条功能与使用指南")
        for guide in product["guides"]:
            for key in ("title", "summary", "note"):
                if not isinstance(guide[key], str):
                    raise ValueError("指南文案必须是字符串")
            if not guide["steps"] or any(not isinstance(s, str) for s in guide["steps"]):
                raise ValueError("指南步骤必须是非空字符串列表")
    return data


def render(config, root):
    data = payload(config, root)
    template = (root / "src/site_template.html").read_text(encoding="utf-8")
    # JSON text is never interpreted as HTML, including user-authored </script>.
    serialized = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    theme = config["THEME"]
    if theme["default_mode"] not in ("dark", "light"):
        raise ValueError("THEME.default_mode 须为 dark 或 light")
    colors = {
        "accent": theme["accent"], "accent-dark": theme["accent_dark"],
        "action-fill": theme.get("action_fill", "#e8a45f"),
        "copy-fill": theme.get("copy_fill", "#9cc5f4"),
        "version-fill": theme.get("version_fill", "#192e4c"),
        "version-selected-fill": theme.get("version_selected_fill", "#243e60"),
        "version-text": theme.get("version_text", "#aacdff"),
    }
    for color in colors.values():
        if not isinstance(color, str) or not re.fullmatch(r"#[a-fA-F0-9]{6}", color):
            raise ValueError("THEME 的颜色须为 #RRGGBB")
    replacements = {
        "__SITE_DATA__": serialized,
        "__SITE_DOWNLOADS__": (root / "src/site_downloads.js").read_text(encoding="utf-8"),
        "__SITE_AUDIO__": (root / "src/site_audio.js").read_text(encoding="utf-8"),
        "__SITE_DEFAULT_THEME__": theme["default_mode"],
        "__SITE_TITLE__": html.escape(data["site"]["title"]),
        "__SITE_DESCRIPTION__": html.escape(data["site"]["intro"], quote=True),
        "__SITE_THEME__": ";".join("--" + name + ":" + color for name, color in colors.items()),
    }
    # Replace tokens in one pass, so custom text that resembles a token stays literal.
    return re.sub("|".join(map(re.escape, replacements)), lambda m: replacements[m.group()], template)


ASSET_PATTERN = r'(<script id="installer-assets" type="application/json">)(.*?)(</script>)'

def installer_assets(text):
    matches=list(re.finditer(ASSET_PATTERN,text,re.S))
    if len(matches)!=1: raise ValueError("网页须有唯一安装包数据区")
    data=json.loads(matches[0].group(2))
    if not isinstance(data,dict) or set(data)!={"packages"} or not isinstance(data["packages"],dict): raise ValueError("安装包数据格式不正确")
    return data

def embedded_snapshot(text):
    installer_assets(text)
    return re.sub(ASSET_PATTERN,lambda m:m.group(1)+'{"packages":{}}'+m.group(3),text,flags=re.S)

def attach_installers(text, packages):
    if installer_assets(text)["packages"]: raise ValueError("不能将带安装包的网页重新用作内嵌快照")
    data=json.loads(re.search(r'<script id="site-data" type="application/json">(.*?)</script>',text,re.S).group(1))
    products={p["id"]:p for p in data["products"]}
    assets={}
    for identity,path in packages.items():
        filename={"word":"WordUpdater_Setup.exe","excel":"Excel_Iterate_Setup.exe"}.get(identity)
        if filename is None: raise ValueError("未知安装包项目")
        content=Path(path).read_bytes()
        if not content.startswith(b'MZ'): raise ValueError("安装包不是 Windows EXE")
        assets[identity]={"filename":filename,"version":products[identity]["version"],"bytes":len(content),"sha256":hashlib.sha256(content).hexdigest(),"base64":base64.b64encode(content).decode('ascii')}
    serialized=json.dumps({"packages":assets},ensure_ascii=False,separators=(',',':')).replace('<','\\u003c')
    return re.sub(ASSET_PATTERN,lambda m:m.group(1)+serialized+m.group(3),text,flags=re.S)

def verify_installers(text, packages):
    """Compare the final page to this delivery's verified installer files."""
    assets=installer_assets(text)['packages']
    if set(assets)!=set(packages): raise ValueError('网页安装包项目与本轮交付不一致')
    for identity,(path,expected_version) in packages.items():
        asset=assets[identity];content=Path(path).read_bytes()
        expected_name={'word':'WordUpdater_Setup.exe','excel':'Excel_Iterate_Setup.exe'}[identity]
        if (asset.get('filename')!=expected_name or asset.get('version')!=expected_version
                or asset.get('bytes')!=len(content) or asset.get('sha256')!=hashlib.sha256(content).hexdigest()
                or base64.b64decode(asset.get('base64',''),validate=True)!=content):
            raise ValueError('网页安装包与正式交付不一致：'+identity)

def preview_installers(root, text):
    # Only packages fingerprinted by a successful central delivery qualify.
    record=root/'build/current-delivery.json'
    try: evidence=json.loads(record.read_text(encoding='utf-8'))
    except (OSError,ValueError): return {}
    data=json.loads(re.search(r'<script id="site-data" type="application/json">(.*?)</script>',text,re.S).group(1))
    versions={p['id']:p['version'] for p in data['products']};result={}
    for identity,folder,name in [('word','C-02-WordUpdater','WordUpdater'),('excel','C-01-ExcelIterate','Excel_Iterate')]:
        path=root.parent/folder/'AA-Program'/(name+'_Setup.exe');item=evidence.get(str(path.resolve()),{})
        try:
            content=path.read_bytes()
            if item.get('version')==versions[identity] and item.get('bytes')==len(content) and item.get('sha256')==hashlib.sha256(content).hexdigest(): result[identity]=path
        except OSError: pass
    return result

def generate(config, root, target=None, standalone=True):
    root = Path(root).resolve()
    text=render(config,root)
    if standalone: text=attach_installers(text,preview_installers(root,text))
    content=text.encode("utf-8")
    target = Path(target).resolve() if target is not None else root / "AA-Program/XyCabin.html"
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.is_file() or target.read_bytes() != content:
        target.write_bytes(content)
    print("网站预览：" + str(target))
    return target


def build(root, target=None):
    root = Path(root).resolve()
    config = vars(load(root / "AA-Program/C3-XycabinSite_Configure.py"))
    return generate(config, root, target or root / "build/site-embedded.snapshot.html", standalone=False)


def validate_snapshot(path, project=None, require_embedded=False):
    """Read-only validation of the immutable HTML supplied to an add-in build."""
    path = Path(path).resolve()
    text = path.read_text(encoding="utf-8")
    if require_embedded and installer_assets(text)["packages"]:
        raise ValueError("加载项只能编译不含安装包数据的内嵌快照，请使用统一重建入口")
    match = re.search(r'<script id="site-data" type="application/json">(.*?)</script>', text, re.S)
    if not match:
        raise ValueError("网站快照缺少数据区：" + str(path))
    data = json.loads(match.group(1))
    ids = {p["id"]: p for p in data["products"]}
    if not {"word", "excel"}.issubset(ids):
        raise ValueError("网站快照必须包含两个加载项")
    if project is not None:
        project = Path(project).resolve()
        host = json.loads((project / "project.json").read_text(encoding="utf-8-sig"))["host"]
        expected = load(project / "AA-Program" / CONFIGURE[project.name]).get_version()
        if ids["word" if host == "Word" else "excel"]["version"] != expected:
            raise ValueError("网站快照与程序版本不一致：" + str(project))
    return path
