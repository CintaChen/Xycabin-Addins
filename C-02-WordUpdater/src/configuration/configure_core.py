import base64
import io
from pathlib import Path
import os
import re
import struct
import sys
import unicodedata

_HERE = os.path.dirname(os.path.abspath(__file__))


def _find_project_root(start):
    """从本文件所在目录往上找项目根：能同时看到 build.py 和 src\\addin 的那一层。

    这样脚本放在项目根还是「AA-Program」文件夹里都能跑，不会把文件写到错误的地方。
    """
    d = start
    for _ in range(5):
        if (os.path.isfile(os.path.join(d, "build.py"))
                and os.path.isdir(os.path.join(d, "src", "addin"))):
            return d
        parent = os.path.dirname(d)
        if parent == d:
            break
        d = parent
    return start


ROOT = _find_project_root(_HERE)
OUT_CS = os.path.join(ROOT, "src", "addin", "Ribbon.g.cs")
OUT_INSTALLER_CS = os.path.join(ROOT, "src", "installer", "InstallerText.g.cs")
OUT_PLATFORMS_CS = os.path.join(ROOT, "src", "addin", "ApiPlatforms.g.cs")
CONNECT_CS = os.path.join(ROOT, "src", "addin", "Connect.cs")


# =============================================================================
#                          生成逻辑（一般不用改）
# =============================================================================

ICON_TARGET_SIZE = {"large": 32, "normal": 16}
_VALID_ACTION = ("feature", "url")
_ID_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")

# 安装器底部那一行的固定前缀（不可配置：大写 V + 全角冒号）
_INSTALLER_VERSION_PREFIX = "Version："

# 结果弹窗文案模板：占位符按【回调】区分（回调决定用哪个引擎的结果对象填空）
_CALLBACK_PLACEHOLDERS = {
    "OnUpdateLinks": ("found", "success", "failed", "changed", "cleaned"),
    "OnLinkToExcel": ("rows", "cols", "cells", "refreshed",
                      "failures", "source", "cleaned"),
}


def die(msg):
    sys.stderr.write("[配置错误] " + msg + "\n")
    sys.exit(2)


def warn(msg):
    sys.stdout.write("[提醒] " + msg + "\n")


def image_size(path):
    """读 PNG / JPEG / BMP / ICO 的像素尺寸；读不出返回 None。"""
    try:
        with open(path, "rb") as f:
            head = f.read(4096)
    except OSError:
        return None
    if head[:8] == b"\x89PNG\r\n\x1a\n" and len(head) >= 24:
        w, h = struct.unpack(">II", head[16:24])
        return int(w), int(h)
    if head[:2] == b"\xff\xd8":  # JPEG：扫 SOFn 段
        i = 2
        while i + 9 < len(head):
            if head[i] != 0xFF:
                i += 1
                continue
            m = head[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                     0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack(">HH", head[i + 5:i + 9])
                return int(w), int(h)
            if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            ln = struct.unpack(">H", head[i + 2:i + 4])[0]
            i += 2 + ln
        return None
    if head[:2] == b"BM" and len(head) >= 26:
        w, h = struct.unpack("<ii", head[18:26])
        return int(abs(w)), int(abs(h))
    if head[:4] == b"\x00\x00\x01\x00":  # ICO
        w, h = head[6], head[7]
        w = 256 if w == 0 else w
        h = 256 if h == 0 else h
        return int(w), int(h)
    return None


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def tip_attr(s):
    """提示文案 -> XML 属性值：\\r\\n 转成 Office 认的硬换行实体。"""
    s = (s or "").replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    return xml_escape("\r\n".join(s.split("\n"))).replace("\r\n", "&#13;&#10;")


def cs_verbatim(s):
    """转 C# @"..." 逐字字符串（双引号翻倍）。"""
    return '@"' + s.replace('"', '""') + '"'


def cs_string_literal(s):
    """转 C# 普通字符串字面量：反斜杠/引号转义，真实换行写成 \\r\\n 转义序列。"""
    out = (s or "").replace("\\", "\\\\").replace('"', '\\"')
    out = out.replace("\r\n", "\\r\\n").replace("\r", "\\r").replace("\n", "\\n")
    return '"' + out + '"'


def _char_w(ch):
    """显示宽度：中文/全角按 2，其他按 1。"""
    return 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1


def _disp_len(s):
    return sum(_char_w(ch) for ch in s)


def _max_line_width(s):
    """一段（\\r\\n 分行）提示里最长一行的显示宽度 —— 悬停方框的宽度由它决定。"""
    lines = (s or "").replace("\r\n", "\n").split("\n")
    return max([_disp_len(x) for x in lines] or [0])


def wrap_tip_text(text, width):
    """按显示宽度给提示自动断行；width 为 None/<=0 时原样返回。

    已有的 \\r\\n 视为强制换行保留；超长行优先在空格处断，没有空格就硬断
    （中文没有空格，硬断正是想要的效果）。
    """
    if not text or not width or width <= 0:
        return text
    out = []
    for para in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if _disp_len(para) <= width:
            out.append(para)
            continue
        rest = para
        while _disp_len(rest) > width:
            w = 0
            cut = -1        # 宽度内最后一个空格的下标（在这里断最自然）
            hard = 0        # 放不下的位置（硬断点）
            for i, ch in enumerate(rest):
                if w + _char_w(ch) > width:
                    hard = max(i, 1)
                    break
                w += _char_w(ch)
                if ch == " ":
                    cut = i
                    hard = i + 1
            if cut > 0:
                out.append(rest[:cut])
                rest = rest[cut + 1:]
            else:
                out.append(rest[:hard])
                rest = rest[hard:]
        out.append(rest)
    return "\r\n".join(out)


def cs_base64_literal(b64, indent):
    """把长 base64 拆成多段拼接，避免超长行。"""
    CH = 96
    chunks = [b64[i:i + CH] for i in range(0, len(b64), CH)] or [""]
    return "\n".join(
        [indent + '"' + chunks[0] + '"'] +
        [indent + '+ "' + c + '"' for c in chunks[1:]])


def arithmetic_defaults():
    import math
    values = globals().get("ARITHMETIC_DEFAULTS", {"add": 100, "subtract": 100, "multiply": 100, "divide": 10000})
    keys = ("add", "subtract", "multiply", "divide")
    if not isinstance(values, dict) or set(values) != set(keys):
        die("ARITHMETIC_DEFAULTS 必须包含 add/subtract/multiply/divide 四项")
    result = {}
    for key in keys:
        value = values[key]
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            die("ARITHMETIC_DEFAULTS 参数必须是数值：" + key)
        try: number = float(value)
        except (OverflowError, ValueError): die("ARITHMETIC_DEFAULTS 数值超出范围：" + key)
        if not math.isfinite(number) or (key == "divide" and number == 0):
            die("ARITHMETIC_DEFAULTS 参数必须有限且除数不能为零：" + key)
        result[key] = format(number, '.0f') if number.is_integer() else repr(number)
    return result

def _build_single_model():
    arithmetic_defaults()
    if TAB_ID_MSO:
        tab_kind, tab_name = "mso", TAB_ID_MSO
        if not _ID_RE.match(TAB_ID_MSO):
            die("TAB_ID_MSO 不是合法标识：" + repr(TAB_ID_MSO))
    else:
        tab_kind, tab_name = "custom", TAB_ID
        if not _ID_RE.match(TAB_ID or ""):
            die("TAB_ID 不是合法标识（只能字母数字下划线，且不能以数字开头）："
                + repr(TAB_ID))
        if not (TAB_LABEL or "").strip():
            die("自定义选项卡必须给 TAB_LABEL 起个显示名")
        if TAB_INSERT_BEFORE_MSO and not _ID_RE.match(TAB_INSERT_BEFORE_MSO):
            die("TAB_INSERT_BEFORE_MSO 不是合法标识：" + repr(TAB_INSERT_BEFORE_MSO))
    if not _HEX_RE.match(ICON_BACKGROUND or ""):
        die("ICON_BACKGROUND 需形如 #RRGGBB，当前是：" + repr(ICON_BACKGROUND))

    icon_dir = ICON_DIR if os.path.isabs(ICON_DIR) else os.path.join(ROOT, ICON_DIR)
    connect_src = ""
    try:
        with open(CONNECT_CS, "r", encoding="utf-8-sig") as f:
            connect_src = f.read()
        for part in sorted(Path(CONNECT_CS).parent.glob("Connect.*.cs")):
            connect_src += "\n" + part.read_text(encoding="utf-8-sig")
    except OSError:
        warn("读不到 " + CONNECT_CS + "，跳过回调存在性检查")

    seen = {}
    group_index = {}
    group_order = []
    model = []
    entries = []
    def visit(items, parent=None, group=None, hidden=False, depth=0):
        if depth > 8:
            die("菜单嵌套超过 8 层")
        if not isinstance(items, (list, tuple)) or not items:
            die("BUTTONS / children 必须是非空列表")
        for original in items:
            if not isinstance(original, dict):
                die("菜单配置条目必须是字典")
            item = dict(original)
            if "default_parameter" in item:
                key = item["default_parameter"]
                values = arithmetic_defaults()
                if key not in values: die("未知 default_parameter：" + str(key))
                item["label"] = "默认" + {"add": "加", "subtract": "减", "multiply": "乘", "divide": "除"}[key] + values[key]
            kind = item.get("type", "button")
            if kind not in ("button", "menu", "toggleButton"):
                die("type 只能是 button/menu/toggleButton")
            if parent is not None and "size" in item:
                die("嵌套菜单项不允许 size；图标统一为 16 个逻辑像素")
            if parent is not None and "group" in item:
                die("嵌套菜单项继承根分组，不允许 group")
            column = item.get("column")
            if column is not None:
                if not isinstance(column, str) or not _ID_RE.match(column):
                    die("column 必须是合法的容器 id")
                if parent is not None or item.get("size", "normal") != "normal":
                    die("column 只允许根层 normal 按钮或菜单")
            item["group"] = item.get("group", group)
            item["size"] = item.get("size", "normal")
            item["hidden"] = hidden or bool(item.get("hidden"))
            item["type"] = kind
            item["parent"] = parent
            if kind == "menu":
                if item.get("action") or item.get("callback") or item.get("url"):
                    die("菜单只展开，不能配置 action/callback/url")
                item["action"] = None
            elif "children" in item:
                die("只有 type=menu 可以有 children")
            if kind == "toggleButton":
                if parent is not None or item.get("action") != "feature" or not item.get("get_pressed"):
                    die("toggleButton 必须是根层 feature，并指定 get_pressed")
            entries.append(item)
            if kind == "menu":
                visit(item.get("children"), item.get("id"), item["group"], item["hidden"], depth + 1)
    visit(BUTTONS)
    for i, b in enumerate(entries):
        tag = "控件[%d]" % i
        for key in ("id", "group", "label", "size", "action"):
            if key not in b:
                die(tag + " 缺少必填字段：" + key)
        if not isinstance(b.get("enabled", True), bool): die(tag + " enabled 必须是布尔值")
        bid = b["id"]
        if not _ID_RE.match(bid):
            die(tag + " id 只能含字母数字下划线：" + repr(bid))
        if bid in seen:
            die("按钮 id 重复：" + bid)
        seen[bid] = True
        if b["size"] not in ICON_TARGET_SIZE:
            die(tag + " size 只能是 large/normal：" + repr(b["size"]))
        if b["type"] != "menu" and b["action"] not in _VALID_ACTION:
            die(tag + " action 只能是 feature/url：" + repr(b["action"]))
        if b["action"] == "feature":
            if not b.get("callback"):
                die(tag + " action=feature 必须填 callback")
            if connect_src and not re.search(
                    r"public\s+void\s+" + re.escape(b["callback"]) + r"\s*\(",
                    connect_src):
                die(tag + " 的回调 " + b["callback"] + " 在 Connect.cs 里找不到，"
                    "请先在 Connect.cs 实现该回调，并生成对应包装方法")
        if b["type"] == "toggleButton" and connect_src:
            require_callback(connect_src, b["callback"], "void", True)
            require_callback(connect_src, b["get_pressed"], "bool")
        elif b["action"] == "url":
            u = (b.get("url") or "").strip()
            if not u:
                die(tag + " action=url 必须填 url")
            if not re.match(r"^[A-Za-z][A-Za-z0-9+.\-]*://|^mailto:", u):
                warn(tag + " 的 url 看起来不像标准地址：" + u)

        # 专属回调契约：XML 里 getImage="GetImage_<id>"、url 按钮
        # onAction="OnOpenUrl_<id>"，WPS 按名字分发 —— 这些一行包装方法
        # 必须真实存在于 Connect.cs（返回值是 IPictureDisp / void）
        if connect_src and not re.search(
                r"GetImage_" + re.escape(bid) + r"\s*\(", connect_src):
            die(tag + " 需要在 Connect.cs 加图标回调包装方法："
                "public IPictureDisp GetImage_" + bid + "(object c)"
                " { return GetButtonImageCore(\"" + bid + "\"); }")
        if b["action"] == "url" and connect_src and not re.search(
                r"OnOpenUrl_" + re.escape(bid) + r"\s*\(", connect_src):
            die(tag + " 需要在 Connect.cs 加网址回调包装方法："
                "public void OnOpenUrl_" + bid + "(object c)"
                " { OpenUrlCore(\"" + bid + "\"); }")

        group_id, glabel = b["group"]
        if not _ID_RE.match(group_id or ""):
            die(tag + " 分组 id 只能含字母数字下划线：" + repr(group_id))
        if group_id not in group_index:
            group_index[group_id] = len(group_order)
            group_order.append((group_id, glabel))

        px = ICON_TARGET_SIZE[b["size"]]
        custom = None
        fname = b.get("icon")
        if isinstance(fname, str):
            fname = fname.strip()
            # 配置里常直接写字符串 "None" 表示"不用自定义图"，等价于留空
            if fname.lower() == "none" or fname == "":
                fname = None
        if fname:
            path = os.path.join(icon_dir, fname)
            if not os.path.isfile(path):
                # 允许只写主文件名（如 "更新Xy"），自动补常见扩展名
                for ext in ("png", "jpg", "jpeg", "bmp", "ico"):
                    cand = os.path.join(icon_dir, fname + "." + ext)
                    if os.path.isfile(cand):
                        path = cand
                        fname = fname + "." + ext
                        break
            if not os.path.isfile(path):
                die(tag + " 指定了图标 " + fname + "，但 " + path + " 不存在")
            raw = open(path, "rb").read()
            dim = image_size(path)
            if dim is None:
                warn(tag + " 无法读取图片尺寸（非 PNG/JPEG/BMP/ICO？）：" + fname)
            elif dim[0] != px or dim[1] != px:
                warn(tag + " 图标 %s 实际 %dx%d，建议 %dx%d，运行时会高质量缩放"
                     % (fname, dim[0], dim[1], px, px))
            if len(raw) > 256 * 1024:
                warn(tag + " 图标 %s 超过 256KB（%d 字节），图标会随 DLL 常驻内存"
                     % (fname, len(raw)))
            custom = {
                "ext": os.path.splitext(fname)[1].lstrip(".").lower(),
                "b64": base64.b64encode(raw).decode("ascii"),
            }
        elif b.get("builtin_icon") not in (None, "update", "paste", "mail", "view", "table", "formula", "filter", "link", "directory", "agent", "web", "document", "settings"):
            die(tag + " builtin_icon 只能是 None/update/paste/mail/view/table/formula/filter/link/directory/agent/web/document/settings")

        # 悬停提示按宽度自动断行：全局 TIP_WRAP_WIDTH，可被按钮级 tip_width 覆盖
        tip_w = b.get("tip_width")
        if tip_w is None:
            tip_w = TIP_WRAP_WIDTH
        if tip_w is not None:
            try:
                tip_w = int(tip_w)
            except (TypeError, ValueError):
                die(tag + " tip_width 必须是整数（半角字符数）："
                    + repr(b.get("tip_width")))
            if tip_w < 12:
                die(tag + " tip_width 太小（至少 12 个半角宽）：" + repr(tip_w))
        screentip = wrap_tip_text(b.get("screentip") or "", tip_w)
        supertip = wrap_tip_text(b.get("supertip") or "", tip_w)

        model.append({
            "id": bid, "gid": group_id, "glabel": glabel,
            "type": b["type"], "parent": b["parent"], "children": [], "column": b.get("column"),
            "label": b["label"], "size": b["size"], "px": px,
            "screentip": screentip,
            "supertip": supertip,
            "action": b["action"], "callback": b.get("callback"), "get_pressed": b.get("get_pressed"),
            "url": (b.get("url") or "").strip(),
            "builtin": b.get("builtin_icon"),
            "hidden": bool(b.get("hidden")),
            "enabled": b.get("enabled",True),
            "icon": custom,
        })

    by_id = {m["id"]: m for m in model}
    for m in model:
        if m["parent"] is not None:
            by_id[m["parent"]]["children"].append(m)
    for m in reversed(model):
        if m["type"] == "menu" and all(child["hidden"] for child in m["children"]):
            m["hidden"] = True
    if set(seen).intersection(group_index) or (tab_kind == "custom" and (tab_name in seen or tab_name in group_index)):
        die("选项卡、分组和控件 id 必须全局唯一")
    groups = []
    columns = {}
    for group_id, glabel in group_order:
        items = [m for m in model if m["gid"] == group_id and m["parent"] is None]
        current, closed = None, set()
        for m in items:
            column = m["column"]
            if column != current:
                if current is not None:
                    closed.add(current)
                if column in closed:
                    die("同一 column 的按钮必须连续排列：" + column)
                current = column
            if column is not None:
                if column in seen or column in group_index or (tab_kind == "custom" and column == tab_name):
                    die("column 容器 id 与选项卡、分组或控件冲突：" + column)
                if column in columns and columns[column] != group_id:
                    die("column 不能跨分组使用：" + column)
                columns[column] = group_id
        groups.append((group_id, glabel, items))
    visible = [m for m in model if not m["hidden"]]
    if not visible:
        die("所有按钮都被 hidden 了，功能区将没有任何按钮")

    # ---- 结果弹窗文案模板（RESULT_TEMPLATES）校验与归并 ----
    # 键=按钮 id；生成的数据表按【回调名】存（同一个回调可以被多个按钮
    # 共用，但给它们配的模板必须一致，否则没法归并）。
    result_texts = {}
    if RESULT_TEMPLATES:
        callback_of = {}
        for b in model:
            if b.get("action") == "feature":
                callback_of[b["id"]] = b.get("callback")
        for key in RESULT_TEMPLATES:
            if not isinstance(key, str) or key not in seen:
                die("RESULT_TEMPLATES 的键必须是 BUTTONS 里某个按钮的 id，当前是："
                    + repr(key))
            cb = callback_of.get(key)
            if cb is None:
                die("RESULT_TEMPLATES 的键 " + key
                    + ' 对应的按钮不是 action="feature"，没有结果弹窗可自定义')
            if cb not in _CALLBACK_PLACEHOLDERS:
                die("RESULT_TEMPLATES 的键 " + key + " 对应回调 " + str(cb)
                    + " 没有带统计数字的完成弹窗，暂不支持自定义；当前可配的回调："
                    + "、".join(sorted(_CALLBACK_PLACEHOLDERS)))
            tpl = RESULT_TEMPLATES[key]
            if not isinstance(tpl, str) or not tpl.strip():
                die("RESULT_TEMPLATES[" + key + "] 的值必须是非空文案模板字符串")
            used = set(re.findall(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", tpl))
            allowed = set(_CALLBACK_PLACEHOLDERS[cb])
            bad = sorted(used - allowed)
            if bad:
                die("RESULT_TEMPLATES[" + key + "] 用了不认识的占位符："
                    + "、".join("{" + x + "}" for x in bad)
                    + "；回调 " + cb + " 可用的占位符："
                    + "、".join("{" + x + "}" for x in _CALLBACK_PLACEHOLDERS[cb]))
            if cb in result_texts and result_texts[cb] != tpl:
                die("回调 " + cb + " 被多个按钮共用，但 RESULT_TEMPLATES 给它们配的"
                    "模板不一致；共用同一回调的按钮请配相同模板")
            result_texts[cb] = tpl
    return groups, model, (tab_kind, tab_name), result_texts


def check_installer_version(val):
    """校验 INSTALLER_VERSION：必须四段式数字、每段 0~65534；通过返回去空格后的值。

    不补零、不推导 —— 配置里写什么就是什么，写错当场 die()（退出码 2）。
    这里是版本规则的唯一实现：本文件生成 .g.cs 用它，build.py /
    gen_manifest.py 也是 import 本文件后复用它（版本逻辑不另开文件）。
    """
    if not isinstance(val, str) or not val.strip():
        die("INSTALLER_VERSION 不能为空，必须是四段式数字，例如 \"0.9.5.0\""
            "（当前是：" + repr(val) + "）")
    s = val.strip()
    parts = s.split(".")
    if len(parts) != 4:
        die("INSTALLER_VERSION 必须是四段式数字（用英文句点分隔），例如 \"0.9.5.0\"；"
            "当前 " + repr(val) + " 是 " + str(len(parts)) + " 段")
    for p in parts:
        if not p or not all(c in "0123456789" for c in p):
            die("INSTALLER_VERSION 的每一段都必须是数字（当前这段是 " + repr(p)
                + "，完整值 " + repr(val) + "）")
        if int(p) > 65534:
            die("INSTALLER_VERSION 每段不能超过 65534（.NET 程序集版本上限）；"
                "当前这段是 " + p + "（完整值 " + repr(val) + "）")
    return s


def api_platforms():
    """校验 API_PLATFORMS（未配置返回 None）：[(显示名, 网址), ...]。

    网址留空 "" 表示占位项（界面里灰色、不可点击）；配置了就必须非空列表，
    名字必须非空，网址要么为空要么像标准地址 —— 写错当场 die()。
    """
    items = globals().get("API_PLATFORMS")
    if items is None:
        return None
    if not isinstance(items, (list, tuple)) or not items:
        die("API_PLATFORMS 必须是非空列表，每项为 (显示名, 网址)")
    names, urls = [], []
    for index, item in enumerate(items):
        tag = "API_PLATFORMS[%d]" % index
        if not isinstance(item, (list, tuple)) or len(item) != 2:
            die(tag + " 必须是 (显示名, 网址) 两项：" + repr(item))
        name, url = item
        if not isinstance(name, str) or not name.strip():
            die(tag + " 显示名必须是非空字符串：" + repr(name))
        if not isinstance(url, str):
            die(tag + " 网址必须是字符串（占位项请写空串 \"\"）：" + repr(url))
        url = url.strip()
        if url and not re.match(r"^[A-Za-z][A-Za-z0-9+.\-]*://", url):
            warn(tag + " 的网址看起来不像标准地址：" + url)
        names.append(name.strip())
        urls.append(url)
    return names, urls


API_PLATFORMS_HEADER = """// =====================================================================
// 本文件由 AA-Program 配置器自动生成，请勿手改。
//
//   改「跳转到API平台」弹窗的选项与网址 -> 编辑 AA-Program 配置器顶部的
//   API_PLATFORMS 配置区，然后重跑：python AA-Program 下的配置器
//
//   每项 = (显示名, 网址)；网址为空字符串表示占位项（灰色、不可点击）。
//
// 生成时间：%(stamp)s
// =====================================================================
namespace %(namespace)s
{
    internal static class ApiPlatforms
    {
        // Names 与 Urls 一一对应；网址为空 = 占位项。
        internal static readonly string[] Names = new string[] {
%(names)s
        };
        internal static readonly string[] Urls = new string[] {
%(urls)s
        };
    }
}
"""


def build_installer_text():
    """安装程序界面配置区校验；返回 (功能说明, 底部版权行, 窗口标题, 主标题, 副标题)。

    版权行 = （INSTALLER_VERSION_LINE_PREFIX + " "） + "Version："
           + INSTALLER_VERSION + （"，" + INSTALLER_COPYRIGHT）
    前缀可留空（未定义或空串时省略），用 globals().get 读取以免未注入方报错。
    """
    for name, val in (("INSTALLER_FEATURE_TEXT", INSTALLER_FEATURE_TEXT),
                      ("INSTALLER_VERSION", INSTALLER_VERSION),
                      ("INSTALLER_COPYRIGHT", INSTALLER_COPYRIGHT),
                      ("INSTALLER_WINDOW_TITLE", INSTALLER_WINDOW_TITLE),
                      ("INSTALLER_TITLE", INSTALLER_TITLE),
                      ("INSTALLER_SUBTITLE", INSTALLER_SUBTITLE)):
        if not isinstance(val, str):
            die(name + " 必须是字符串（留空请写 \"\"），当前是：" + repr(val))
    ver = check_installer_version(INSTALLER_VERSION)
    line = _INSTALLER_VERSION_PREFIX + ver
    prefix = globals().get("INSTALLER_VERSION_LINE_PREFIX", "")
    if prefix:
        line = prefix + " " + line
    if INSTALLER_COPYRIGHT:
        line += "，" + INSTALLER_COPYRIGHT
    return (INSTALLER_FEATURE_TEXT, line,
            INSTALLER_WINDOW_TITLE, INSTALLER_TITLE, INSTALLER_SUBTITLE)


# =============================================================================
#                     版本号同步（版本准绳的"下游"）
#          build.py / gen_manifest.py 会 import 本文件后调用这两个函数
# =============================================================================
#
# 版本号只认上面的 INSTALLER_VERSION。这里做的事是把它"下发"到
# src\addin\AssemblyInfo.cs（程序集版本）与 VSTO 清单（gen_manifest.py
# 读 get_version()）。本函数不补零、不推导 —— 配置写什么就是什么；
# 版本变更只按用户明确指定的新版本修改，本配置器不自动递增。
#
ASSEMBLY_INFO_CS = os.path.join(ROOT, "src", "addin", "AssemblyInfo.cs")


def _read_text_file(path):
    """读文本文件，自动处理 BOM 与 gbk 兜底；返回 (文本, 是否带 BOM)。"""
    with io.open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return raw.decode(enc), raw.startswith(b"\xef\xbb\xbf")
        except UnicodeDecodeError:
            continue
    die("读不出这个文件的编码：" + path)


def _line_eol(line):
    """取这一行原本的行尾，回写时原样保留（AssemblyInfo.cs 是 LF）。"""
    if line.endswith("\r\n"):
        return "\r\n"
    if line.endswith("\n"):
        return "\n"
    return ""


# rule=R-17
def get_version():
    """返回四段式版本号，如 "0.9.5.0"（VSTO 清单与程序集都用这个）。

    唯一来源 = INSTALLER_VERSION；顺带做一次格式校验（不合法当场 die）。
    """
    return check_installer_version(INSTALLER_VERSION)


def sync_assembly_info(path=ASSEMBLY_INFO_CS, ver=None):
    """把 AssemblyVersion / AssemblyFileVersion 回写成当前版本（不用正则）。

    逐行扫描，按「[assembly: 前缀 XXX(" ... ")]」定位，命中数必须各为 1，
    否则说明文件结构变了 —— 宁可报错也不瞎改。
    返回 (changed, ver)：changed=False 表示本来就是对的、没动文件。
    """
    if ver is None:
        ver = get_version()
    text, bom = _read_text_file(path)
    lines = text.splitlines(True)          # 保留原行尾，LF/CRLF 原样回去
    out = []
    hits = {"AssemblyVersion": 0, "AssemblyFileVersion": 0}
    changed = False
    for line in lines:
        s = line.strip()
        if s.startswith("[assembly:") and s.endswith('")]'):
            inner = s[len("[assembly:"):-3].strip()   # 形如：AssemblyVersion("0.9.5.0"
            for attr in ("AssemblyVersion", "AssemblyFileVersion"):
                head = attr + '("'
                if inner.startswith(head):
                    hits[attr] += 1
                    # inner 形如 AssemblyVersion("0.9.5.0"（结尾引号已随 [:-3] 去掉），
                    # 所以 head 之后整段就是旧值；比对一致才不动文件。
                    if inner[len(head):] != ver:
                        changed = True
                        line = "[assembly: " + attr + '("' + ver + '")]' + _line_eol(line)
                    break
        out.append(line)
    for attr in ("AssemblyVersion", "AssemblyFileVersion"):
        if hits[attr] != 1:
            die("AssemblyInfo.cs 里 " + attr + " 命中 " + str(hits[attr])
                + " 次（应为 1 次），不敢瞎改：" + path)
    if not changed:
        return False, ver
    data = "".join(out).encode("utf-8")
    if bom:
        data = b"\xef\xbb\xbf" + data
    with io.open(path, "wb") as f:
        f.write(data)
    return True, ver


def require_callback(source, name, result, pressed=False):
    if not isinstance(name, str) or not _ID_RE.match(name):
        die("回调名称不是合法标识：" + repr(name))
    args = r"object\s+\w+" + (r"\s*,\s*bool\s+\w+" if pressed else "")
    if not re.search(r"public\s+" + result + r"\s+" + re.escape(name) + r"\s*\(\s*" + args + r"\s*\)", source):
        die("Connect 分部缺少匹配签名的回调：" + name)


def build_model():
    global BUTTONS, TAB_ID_MSO, TAB_ID, TAB_LABEL, TAB_INSERT_BEFORE_MSO, RESULT_TEMPLATES
    original = (BUTTONS, TAB_ID_MSO, TAB_ID, TAB_LABEL, TAB_INSERT_BEFORE_MSO, RESULT_TEMPLATES)
    groups, model, tab, result_texts = _build_single_model()
    for m in model: m["tab_id"] = tab[1]
    tabs = globals().get("EXTRA_TABS", [])
    if not isinstance(tabs, (list, tuple)): die("EXTRA_TABS 必须是列表")
    source = ""
    if Path(CONNECT_CS).is_file():
        source = Path(CONNECT_CS).read_text(encoding="utf-8-sig")
        source += "\n".join(f.read_text(encoding="utf-8-sig") for f in sorted(Path(CONNECT_CS).parent.glob("Connect.*.cs")))
    onload = globals().get("RIBBON_ON_LOAD")
    if onload and source: require_callback(source, onload, "void")
    used = set()
    def reserve(value):
        if value in used: die("所有选项卡、分组、容器和控件 id 必须全局唯一：" + value)
        used.add(value)
    def reserve_tab(tab_id, gs, ms):
        reserve(tab_id)
        for gid, label, items in gs: reserve(gid)
        for m in ms: reserve(m["id"])
        for col in {m["column"] for m in ms if m["column"]}: reserve(col)
    reserve_tab(tab[1], groups, model)
    try:
        for extra in tabs:
            if not isinstance(extra, dict): die("EXTRA_TABS 条目必须是字典")
            TAB_ID_MSO, TAB_ID, TAB_LABEL = None, extra.get("id"), extra.get("label")
            TAB_INSERT_BEFORE_MSO = extra.get("insert_before_mso")
            BUTTONS, RESULT_TEMPLATES = extra.get("buttons"), {}
            visible = extra.get("get_visible")
            if visible and source: require_callback(source, visible, "bool")
            if visible and not _ID_RE.match(visible): die("get_visible 非法")
            gs, ms, xt, _ = _build_single_model()
            reserve_tab(xt[1], gs, ms)
            for m in ms: m["tab_id"] = xt[1]
            groups.extend(gs); model.extend(ms)
    finally:
        BUTTONS, TAB_ID_MSO, TAB_ID, TAB_LABEL, TAB_INSERT_BEFORE_MSO, RESULT_TEMPLATES = original
    return groups, model, tab, result_texts


def build_xml(groups, tab):
    onload = globals().get("RIBBON_ON_LOAD")
    lines = ['<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"' +
             (' onLoad="%s"' % onload if onload else '') + '>', '  <ribbon>', '    <tabs>']
    tabs = [{"id": tab[1], "label": TAB_LABEL, "mso": tab[0] == "mso", "insert_before_mso": TAB_INSERT_BEFORE_MSO}] + list(globals().get("EXTRA_TABS", []))
    def emit(m, indent):
        if m["hidden"]: return
        attrs = ['id="%s"' % m["id"], 'label="%s"' % xml_escape(m["label"])]
        if m["parent"] is None: attrs.append('size="%s"' % m["size"])
        attrs.append('getImage="GetImage_%s"' % m["id"])
        if not m.get('enabled',True): attrs.append('enabled="false"')
        for key in ("screentip", "supertip"):
            if m[key]: attrs.append('%s="%s"' % (key, tip_attr(m[key])))
        if m["type"] == "menu":
            attrs.append('itemSize="normal"')
            lines.append(indent + '<menu ' + ' '.join(attrs) + '>')
            for child in m["children"]: emit(child, indent + '  ')
            lines.append(indent + '</menu>')
        else:
            cb = m["callback"] if m["action"] == "feature" else "OnOpenUrl_" + m["id"]
            attrs.append('onAction="%s"' % cb)
            if m["type"] == "toggleButton": attrs.append('getPressed="%s"' % m["get_pressed"])
            lines.append(indent + '<' + m["type"] + ' ' + ' '.join(attrs) + ' />')
    for current_tab in tabs:
        attrs = ('idMso="%s"' % current_tab["id"] if current_tab.get("mso") else
                 'id="%s" label="%s"' % (current_tab["id"], xml_escape(current_tab["label"])))
        for config_key, xml_key in (("get_visible", "getVisible"), ("insert_before_mso", "insertBeforeMso")):
            if current_tab.get(config_key): attrs += ' %s="%s"' % (xml_key, xml_escape(current_tab[config_key]))
        lines.append('      <tab ' + attrs + '>')
        for gid, glabel, items in groups:
            vis = [m for m in items if not m["hidden"] and m["tab_id"] == current_tab["id"]]
            if not vis: continue
            lines.append('        <group id="%s" label="%s">' % (gid, xml_escape(glabel)))
            column = None
            for m in vis:
                if column != m["column"]:
                    if column: lines.append('          </box>')
                    column = m["column"]
                    if column: lines.append('          <box id="%s" boxStyle="vertical">' % column)
                emit(m, '            ' if column else '          ')
            if column: lines.append('          </box>')
            lines.append('        </group>')
        lines.append('      </tab>')
    lines.extend(['    </tabs>', '  </ribbon>', '</customUI>'])
    # The C# text writer converts LF to CRLF once; tooltip entities are separate.
    return "\n".join(lines)


ICON_README = """图标放置与规格说明（本文件由 AA-Program 配置器自动生成）
================================================================

【放哪里】
  把图片文件直接丢进本目录（AB-Graphic），
  然后在 AA-Program 配置器的配置区里，把对应按钮的 icon 写成文件名即可，例如：

      "icon": "update.png",

  只写文件名，不要写路径。

【格式】
  推荐 PNG（支持透明）；JPEG / BMP / ICO 也能用。ICO 会被按主图标取图。

【尺寸】
  根控件 size="large"  -> 32 x 32 逻辑像素
  根控件 size="normal" -> 16 x 16 逻辑像素
  嵌套菜单项不配置 size -> 16 x 16 逻辑像素（itemSize="normal"）
  给了 2 倍图（如 64x64）也可以，程序会自动高质量缩放，但 1:1 最清晰。
  非正方形的图会被拉伸成正方形。

【透明】
  WPS 功能区按钮只接收 HBITMAP，不支持真正的透明通道。
  带透明像素的图会先把透明部分合成到 %s 底色上（可在配置区 ICON_BACKGROUND 改）。
  想要绝对干净的效果，请直接出不带透明通道的图。

【为什么改完图还要重跑脚本 + 重新编译】
  加载项运行时 Assembly.Location 指向 CLR 的 dl3 影子缓存而不是安装目录，
  “运行时去目录里读图片”在 WPS 里必然失败。
  所以图片是在编译前被转成 base64 常量写进 src/addin/Ribbon.g.cs 的，随 DLL 一起走。

【流程】
  1. 改 AA-Program 配置器的配置区（标题 / 提示 / 图标 / 网址）
  2. python AA-Program 下的配置器
  3. python build.py
  4. python resign.py            （生成清单 + SHA-256 双签名，自动完成）
  5. python build_installer.py
  6. 完全退出 WPS 后运行 AA-Program目录中的本项目 Setup.exe 重装，重启 WPS 生效。

【不要】
  不要手改 src/addin/Ribbon.g.cs —— 每次跑脚本都会被整文件覆盖。
""" % ICON_BACKGROUND


HEADER = """// =====================================================================
// 本文件由 AA-Program 配置器自动生成，请勿手改。
//
//   改标题 / 提示文案 / 图标 / 跳转网址 -> 编辑 AA-Program 配置器顶部的
//    「按钮配置区」，然后重跑：python AA-Program 下的配置器
//
//   负责向 Connect.cs 提供四样东西：
//     RibbonXml        功能区 XML（选项卡 / 分组 / 按钮 / 提示文案）
//     GetImage(id)     自定义图标 -> IPictureDisp；没有就用内置矢量图标
//     GetUrl(id)       action="url" 按钮的目标地址
//     GetResultText(cb) 结果弹窗的自定义文案模板（配置区 RESULT_TEMPLATES）
//
//   注意：XML 里每个按钮的 getImage / url 按钮的 onAction 都是【专属回调名】
//   （GetImage_<按钮id> / OnOpenUrl_<按钮id>），函数名即按钮身份 ——
//   网络盘文档的回调代理读不出 control.Id，靠名字才能认对按钮。
//   新增按钮时记得在 Connect.cs 加对应的一行包装方法（本脚本会校验）。
//
// 生成时间：%(stamp)s
// =====================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace %(namespace)s
{
    internal static class RibbonConfig
    {
        // 内置图标代号（Connect.cs 分发到功能区专用绘图与共享 OLE 尺寸转换）
        public const string BuiltinUpdate = "update";
        public const string BuiltinPaste = "paste";
        public const string BuiltinMail = "mail";
        public const string BuiltinView = "view";
        public const string BuiltinTable = "table";
        public const string BuiltinFormula = "formula";
        public const string BuiltinFilter = "filter";
        public const string BuiltinLink = "link";
        public const string BuiltinDirectory = "directory";
        public const string BuiltinAgent = "agent";
        public const string BuiltinWeb = "web";
        public const string BuiltinDocument = "document";
        public const string BuiltinSettings = "settings";
        public static string ArithmeticDefault(string key)
        {
%(arithmetic)s
            throw new ArgumentException("Unknown arithmetic default", "key");
        }

        public static string RibbonXml { get { return Xml; } }

        private const string Xml = %(xml)s;

        // 透明区域合成底色（HBITMAP 不支持透明通道，只能先垫一层底）
        private static readonly Color PlateColor = ColorTranslator.FromHtml("%(plate)s");

        private static readonly Dictionary<string, int> Sizes = MakeSizes();
        private static readonly Dictionary<string, string> Icons = MakeIcons();
        private static readonly Dictionary<string, string> BuiltinIcons = MakeBuiltinIcons();
        private static readonly Dictionary<string, string> Urls = MakeUrls();
        private static readonly Dictionary<string, string> ResultTexts = MakeResultTexts();
        private static readonly Dictionary<string, IPictureDisp> Cache =
            new Dictionary<string, IPictureDisp>();

        /// <summary>自定义图标；该按钮没配图则返回 null（调用方会改用内置图标）。</summary>
        public static IPictureDisp GetImage(string id)
        {
            try
            {
                if (string.IsNullOrEmpty(id)) return null;
                string b64;
                if (!Icons.TryGetValue(id, out b64)) return null;
                lock (Cache)
                {
                    IPictureDisp hit;
                    if (Cache.TryGetValue(id, out hit) && hit != null) return hit;
                }
                int px = GetImageSize(id);
                IPictureDisp pic = ToPicture(Convert.FromBase64String(b64), px);
                lock (Cache) Cache[id] = pic;
                return pic;
            }
            catch (Exception ex)
            {
                // 静默返回 null 的时代 WPS 只"白给你看"，什么线索都没有；
                // 现在记一笔日志，并把缓存清掉让下次回调重建
                try { Connect.Log("GetImage(" + (id ?? "") + ") 异常：" + ex.Message); }
                catch { }
                lock (Cache) Cache.Remove(id);
                return null;
            }
        }

        // Shared by custom and built-in image callbacks; public callbacks stay unchanged.
        internal static int GetImageSize(string id)
        {
            int size;
            return id != null && Sizes.TryGetValue(id, out size) ? size : 32;
        }

        /// <summary>该按钮回落用的内置图标代号；没有则返回 null。</summary>
        public static string BuiltinIconKind(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            string kind;
            return BuiltinIcons.TryGetValue(id, out kind) ? kind : null;
        }

        /// <summary>action="url" 按钮的目标地址；没有则返回 null。</summary>
        public static string GetUrl(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            string url;
            return Urls.TryGetValue(id, out url) ? url : null;
        }

        /// <summary>
        /// 结果弹窗的自定义文案模板（键=回调名，如 "OnUpdateLinks"）。
        /// 没有配置则返回 null，调用方（Connect.cs 的 FormatResult / FormatLinkResult）
        /// 会改用内置默认文案。
        /// </summary>
        public static string GetResultText(string callback)
        {
            if (string.IsNullOrEmpty(callback)) return null;
            string tpl;
            return ResultTexts.TryGetValue(callback, out tpl) ? tpl : null;
        }

        // ------------------------------ 数据表 ------------------------------

        private static Dictionary<string, string> MakeIcons()
        {
            var d = new Dictionary<string, string>();
%(icons)s
            return d;
        }

        private static Dictionary<string, int> MakeSizes()
        {
            var d = new Dictionary<string, int>();
%(sizes)s
            return d;
        }

        private static Dictionary<string, string> MakeBuiltinIcons()
        {
            var d = new Dictionary<string, string>();
%(builtins)s
            return d;
        }

        private static Dictionary<string, string> MakeUrls()
        {
            var d = new Dictionary<string, string>();
%(urls)s
            return d;
        }

        private static Dictionary<string, string> MakeResultTexts()
        {
            var d = new Dictionary<string, string>();
%(results)s
            return d;
        }

        // ------------------------- base64 -> IPictureDisp -------------------------

        private static IPictureDisp ToPicture(byte[] raw, int px)
        {
            using (MemoryStream ms = new MemoryStream(raw))
            using (Bitmap src = new Bitmap(ms))
            {
                // Preserve configured plate color and custom-image priority. OLE cannot
                // reliably show alpha in WPS, so compose an opaque 24bpp physical bitmap.
                return RibbonImages.Create(px, delegate(int width, int height)
                {
                    Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
                    try
                    {
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.Clear(PlateColor);
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.DrawImage(src, 0, 0, width, height);
                        }
                        return bitmap;
                    }
                    catch { bitmap.Dispose(); throw; }
                }, Connect.Log);
            }
        }
    }
}
"""


INSTALLER_TEXT_HEADER = """// =====================================================================
// 本文件由 AA-Program 配置器自动生成，请勿手改。
//
//   改安装器顶部色带的功能说明 / 底部版权行 -> 编辑 配置器脚本
//   的「安装程序界面配置区」，然后重跑配置脚本并重新打包安装器。
//
// 生成时间：%(stamp)s
// =====================================================================

namespace %(installer_namespace)s
{
    /// <summary>安装器界面文字（Installer.cs 在布局时读取）。</summary>
    internal static class InstallerText
    {
        /// <summary>窗口标题栏文字。</summary>
        internal const string WindowTitle = %(window_title)s;

        /// <summary>顶部色带主标题（大字）。</summary>
        internal const string Title = %(title)s;

        /// <summary>顶部色带副标题（小字，主标题下方）。</summary>
        internal const string Subtitle = %(subtitle)s;

        /// <summary>顶部色带里的功能说明；空串 = 不显示该 label。</summary>
        internal const string FeatureText = %(feature)s;

        /// <summary>窗口底部居中的版权行（Version：版本，文案）；空串 = 不显示。</summary>
        internal const string Copyright = %(copyright)s;
    }
}
"""


def main():
    groups, model, tab, result_texts = build_model()
    feature_text, copyright_line, window_title, title_text, subtitle_text = build_installer_text()

    icon_dir = ICON_DIR if os.path.isabs(ICON_DIR) else os.path.join(ROOT, ICON_DIR)
    os.makedirs(icon_dir, exist_ok=True)
    with open(os.path.join(icon_dir, "README.txt"), "w", encoding="utf-8") as f:
        f.write(ICON_README)

    icon_lines, size_lines, builtin_lines, url_lines, result_lines = [], [], [], [], []
    IND = "            "
    for m in model:
        if m["icon"]:
            icon_lines.append('%sd["%s"] =\n%s;' % (
                IND, m["id"], cs_base64_literal(m["icon"]["b64"], IND + "    ")))
        size_lines.append('%sd["%s"] = %d;' % (IND, m["id"], m["px"]))
        if m["builtin"]:
            builtin_lines.append(
                '%sd["%s"] = Builtin%s;'
                % (IND, m["id"],
                   {"update": "Update", "paste": "Paste", "mail": "Mail", "view": "View",
                    "table": "Table", "formula": "Formula", "filter": "Filter", "link": "Link", "directory": "Directory", "agent": "Agent", "web": "Web", "document": "Document", "settings": "Settings"}[m["builtin"]]))
        if m["action"] == "url":
            url_lines.append('%sd["%s"] = %s;' % (IND, m["id"], cs_verbatim(m["url"])))
    # 结果弹窗模板按回调名入表（键值已在 build_model 里校验过）
    for cb in sorted(result_texts):
        result_lines.append('%sd["%s"] = %s;'
                            % (IND, cb, cs_string_literal(result_texts[cb])))

    if not icon_lines:
        icon_lines.append(IND + "// 未配置任何自定义图标（全部按钮使用内置矢量图标）")
    if not builtin_lines:
        builtin_lines.append(IND + "// 未指定内置图标回落")
    if not url_lines:
        url_lines.append(IND + "// 未配置跳转网址")
    if not result_lines:
        result_lines.append(IND + "// 未自定义结果弹窗文案（用 Connect.cs 里的内置默认文案）")

    import datetime
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    body = HEADER % {
        "stamp": stamp,
        "arithmetic": "\n".join('            if (key == "%s") return "%s";' % (key,value) for key,value in arithmetic_defaults().items()),
        "xml": cs_verbatim(build_xml(groups, tab)),
        "plate": ICON_BACKGROUND if ICON_BACKGROUND.startswith("#") else "#" + ICON_BACKGROUND,
        "icons": "\n".join(icon_lines),
        "sizes": "\n".join(size_lines),
        "builtins": "\n".join(builtin_lines),
        "urls": "\n".join(url_lines),
        "results": "\n".join(result_lines),
    }

    os.makedirs(os.path.dirname(OUT_CS), exist_ok=True)
    with open(OUT_CS, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(body)

    # Dedicated callback names remain reliable in WPS proxy dispatch.
    callbacks = ['// Generated; canonical command identities are never renamed.', 'namespace ' + __import__('json').loads(Path(os.path.join(ROOT, 'project.json')).read_text(encoding='utf-8-sig'))['namespace'], '{', '    public partial class Connect', '    {']
    for command in model:
        callbacks.append('        public string GetLabel_%s(object control) { return RibbonCustomization.Label("%s"); }' % (command['id'], command['id']))
        if command['type'] == 'menu' and command['parent'] is None:
            callbacks.append('        public string GetContent_%s(object control) { return RibbonCustomization.Content("%s"); }' % (command['id'], command['id']))
    callbacks.extend(['    }', '}', ''])
    with open(os.path.join(os.path.dirname(OUT_CS), 'CustomizationCallbacks.g.cs'), 'w', encoding='utf-8', newline='\r\n') as stream:
        stream.write('\n'.join(callbacks))

    # 安装器界面文字 -> src\installer\InstallerText.g.cs
    installer_body = INSTALLER_TEXT_HEADER % {
        "stamp": stamp,
        "feature": cs_string_literal(feature_text),
        "copyright": cs_string_literal(copyright_line),
        "window_title": cs_string_literal(window_title),
        "title": cs_string_literal(title_text),
        "subtitle": cs_string_literal(subtitle_text),
    }
    os.makedirs(os.path.dirname(OUT_INSTALLER_CS), exist_ok=True)
    with open(OUT_INSTALLER_CS, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(installer_body)

    # 「跳转到API平台」选项 -> src\addin\ApiPlatforms.g.cs（未配置则不生成）
    platforms = api_platforms()
    if platforms is not None:
        platform_names, platform_urls = platforms
        namespace = __import__('json').loads(
            Path(os.path.join(ROOT, 'project.json')).read_text(encoding='utf-8-sig'))['namespace']
        platforms_body = API_PLATFORMS_HEADER % {
            "stamp": stamp,
            "namespace": namespace,
            "names": "\n".join('            ' + cs_string_literal(n) + ',' for n in platform_names),
            "urls": "\n".join('            ' + cs_string_literal(u) + ',' for u in platform_urls),
        }
        os.makedirs(os.path.dirname(OUT_PLATFORMS_CS), exist_ok=True)
        with open(OUT_PLATFORMS_CS, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(platforms_body)
        print("OK -> " + os.path.relpath(OUT_PLATFORMS_CS, ROOT)
              + "   （%d 个API平台选项：%s）" % (len(platform_names), "、".join(platform_names)))

    vis = [m for m in model if not m["hidden"]]
    tab_kind, tab_name = tab
    tab_desc = ("内置选项卡 " + tab_name if tab_kind == "mso"
                else "自定义选项卡「%s」" % TAB_LABEL
                     + ("（插在 " + TAB_INSERT_BEFORE_MSO + " 之前）"
                        if TAB_INSERT_BEFORE_MSO else "（追加到最右）"))
    print("OK -> " + os.path.relpath(OUT_CS, ROOT)
          + "   （项目根：" + ROOT + "）")
    print("    %s：%d 个分组，%d 个控件（含嵌套菜单项）"
          % (tab_desc,
             len([g for g in groups if any(not i["hidden"] for i in g[2])]), len(vis)))
    for gid, glabel, items in groups:
        shown = [i for i in items if not i["hidden"]]
        if not shown:
            print("      [%s] %s —— 全部隐藏" % (gid, glabel))
            continue
        print("      [%s] %s" % (gid, glabel))
        def report(m, depth=0):
            if m["hidden"]:
                return
            icon_desc = "自定义图" if m["icon"] else "内置图 " + str(m["builtin"] or "无")
            tail = "展开菜单" if m["type"] == "menu" else (m["url"] if m["action"] == "url" else m["callback"] + "()")
            print("        " + "  " * depth + "%s · %s · %d逻辑像素 · %s · %s" % (m["id"], m["label"], m["px"], icon_desc, tail))
            for child in m["children"]:
                report(child, depth + 1)
        for m in shown:
            report(m)
    if any(m["hidden"] for m in model):
        print("    已隐藏：" + ", ".join(m["id"] for m in model if m["hidden"]))

    # 悬停方框宽度报告：方框宽度由【最长一行】撑出，给出数字方便对比调整
    tip_rows = [(m["id"], _max_line_width(m["screentip"]),
                 _max_line_width(m["supertip"]))
                for m in vis if m["screentip"] or m["supertip"]]
    if tip_rows:
        print("    悬停方框宽度（方框由最长一行撑出，数字=半角字符宽，中文按 2 计）：")
        for bid, w1, w2 in tip_rows:
            print("        %-16s 标题行 %d / 说明行 %d" % (bid, w1, w2))
        if TIP_WRAP_WIDTH is not None:
            print("        （全局 TIP_WRAP_WIDTH=%d 生效中，超宽行已自动断行）"
                  % int(TIP_WRAP_WIDTH))
    if result_texts:
        print("    结果弹窗自定义文案已生效：" + "、".join(sorted(result_texts)))
        print("        （失败明细、后台进程释放说明仍会自动附加在文案后面）")

    print("OK -> " + os.path.relpath(OUT_INSTALLER_CS, ROOT))
    print("    安装器界面：功能说明 %d 字%s / 版本号 %s / 版权文案 %s"
          % (len(feature_text),
             "" if feature_text else "（已隐藏）",
             INSTALLER_VERSION,
             ("「" + INSTALLER_COPYRIGHT + "」") if INSTALLER_COPYRIGHT else "（留空）"))
    print("        （底部整行显示：" + copyright_line + "）")

    # 版本号下发到程序集（build.py 也会做一次，这里是让"跑完配置脚本"就一致）
    changed, ver = sync_assembly_info()
    print("    版本号 %s -> %s%s"
          % (ver, os.path.relpath(ASSEMBLY_INFO_CS, ROOT),
             "（已更新）" if changed else "（已是该版本）"))
    print("\n下一步：python build.py  （若用自定义图标，图片已随本文件嵌入 DLL）")



def bind(settings, config_file):
    """Bind one configuration to this isolated module without generating files."""
    globals().update({k: v for k, v in settings.items() if k.isupper()})
    global ROOT, OUT_CS, OUT_INSTALLER_CS, OUT_PLATFORMS_CS, CONNECT_CS, ASSEMBLY_INFO_CS, HEADER, INSTALLER_TEXT_HEADER
    ROOT = _find_project_root(os.path.dirname(os.path.abspath(config_file)))
    OUT_CS = os.path.join(ROOT, "src", "addin", "Ribbon.g.cs")
    OUT_INSTALLER_CS = os.path.join(ROOT, "src", "installer", "InstallerText.g.cs")
    OUT_PLATFORMS_CS = os.path.join(ROOT, "src", "addin", "ApiPlatforms.g.cs")
    CONNECT_CS = os.path.join(ROOT, "src", "addin", "Connect.cs")
    ASSEMBLY_INFO_CS = os.path.join(ROOT, "src", "addin", "AssemblyInfo.cs")
    import json
    with open(os.path.join(ROOT, "project.json"), encoding="utf-8-sig") as stream:
        project = json.load(stream)
    HEADER = HEADER.replace("%(namespace)s", project["namespace"])
    INSTALLER_TEXT_HEADER = INSTALLER_TEXT_HEADER.replace("%(installer_namespace)s", project["namespace"] + "Setup")
