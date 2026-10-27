# -*- coding: utf-8 -*-
"""Generate the host-specific VSTO manifests from the shared pipeline."""
import base64
import hashlib
import os
import sys
sys.dont_write_bytecode = True
from project_support import PROJECT, CONFIGURE_NAME
from xml.sax.saxutils import escape

ROOT = os.path.dirname(os.path.abspath(__file__))
BUILD = os.path.join(ROOT, "build")

# 版本号不在这里维护 —— 唯一准绳是 AA-Program 配置器的 INSTALLER_VERSION
# （四段式）。改版本只改那里，此处自动跟随。
# 配置器有 __name__ == "__main__" 保护，可以安全地当模块导入。
def _load_configure():
    import importlib.util
    cfg = os.path.join(ROOT, "AA-Program", CONFIGURE_NAME)
    if not os.path.isfile(cfg):
        sys.exit("[manifest] 找不到配置脚本：" + cfg)
    old = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec = importlib.util.spec_from_file_location("_wps_link_configure", cfg)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
    finally:
        sys.dont_write_bytecode = old
    return mod


ADDIN_VERSION = _load_configure().get_version()
ASSEMBLY = PROJECT["assembly"]
MAIN_DLL = ASSEMBLY + ".dll"
PRODUCT = escape(PROJECT["display_name"], {'"': "&quot;"})
# 随包分发（不在 GAC）的依赖：文件名 -> (程序集名, 版本, publicKeyToken 或 None)
# 注意：Interop.Microsoft.Office.Core 是未签名程序集（无 pkt，v2.8.0.0），
# 与本产物 AssemblyRef 完全一致（ReflectionProbe 实测）。
BUNDLED = {
    "Microsoft.Office.Tools.Common.v4.0.Utilities.dll":
        ("Microsoft.Office.Tools.Common.v4.0.Utilities", "10.0.0.0", "B03F5F7F11D50A3A"),
    "Interop.Microsoft.Office.Core.dll":
        ("Interop.Microsoft.Office.Core", "2.8.0.0", None),
}
# 需存在于 GAC 的引用（preRequisite）：程序集名 -> (版本, publicKeyToken)
PREREQ = [
    ("Microsoft.Windows.CommonLanguageRuntime", "4.0.30319.0", None),
    ("Microsoft.CSharp", "4.0.0.0", "B03F5F7F11D50A3A"),
    ("Microsoft.Office.Tools", "10.0.0.0", "B03F5F7F11D50A3A"),
    ("Microsoft.Office.Tools.Common", "10.0.0.0", "B03F5F7F11D50A3A"),
    ("Microsoft.Office.Tools.v4.0.Framework", "10.0.0.0", "B03F5F7F11D50A3A"),
    ("Microsoft.Office.Tools." + PROJECT["host"], "10.0.0.0", "B03F5F7F11D50A3A"),
    ("Microsoft.VisualStudio.Tools.Applications.Runtime", "10.0.0.0", "B03F5F7F11D50A3A"),
    ("System", "4.0.0.0", "B77A5C561934E089"),
    ("System.Core", "4.0.0.0", "B77A5C561934E089"),
    ("System.Drawing", "4.0.0.0", "B03F5F7F11D50A3A"),
    ("System.Windows.Forms", "4.0.0.0", "B77A5C561934E089"),
]


def sha256_b64(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return base64.b64encode(h.digest()).decode("ascii")


def file_size(path):
    return os.path.getsize(path)


def hash_block(path):
    return (
        "      <hash>\r\n"
        "        <dsig:Transforms>\r\n"
        "          <dsig:Transform Algorithm=\"urn:schemas-microsoft-com:HashTransforms.Identity\" />\r\n"
        "        </dsig:Transforms>\r\n"
        "        <dsig:DigestMethod Algorithm=\"http://www.w3.org/2000/09/xmldsig#sha256\" />\r\n"
        "        <dsig:DigestValue>" + sha256_b64(path) + "</dsig:DigestValue>\r\n"
        "      </hash>\r\n"
    )


def gen_app_manifest(dirpath):
    main_path = os.path.join(dirpath, MAIN_DLL)
    parts = []
    parts.append('<?xml version="1.0" encoding="utf-8"?>\r\n')
    parts.append('<asmv1:assembly xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd" manifestVersion="1.0" xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" xmlns="urn:schemas-microsoft-com:asm.v2" xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" xmlns:dsig="http://www.w3.org/2000/09/xmldsig#" xmlns:co.v2="urn:schemas-microsoft-com:clickonce.v2">\r\n')
    # 占位 publicKeyToken：签名时 SignedCmiManifest2.ReplacePublicKeyToken 会替换为证书真实 token
    parts.append('  <asmv1:assemblyIdentity name="' + MAIN_DLL + '" version="' + ADDIN_VERSION + '" publicKeyToken="0000000000000000" language="neutral" processorArchitecture="msil" type="win32" />\r\n')
    parts.append('  <description xmlns="urn:schemas-microsoft-com:asm.v1">' + PRODUCT + '</description>\r\n')
    parts.append('  <application />\r\n')
    parts.append('  <entryPoint>\r\n')
    parts.append('    <co.v1:customHostSpecified />\r\n')
    parts.append('  </entryPoint>\r\n')
    parts.append('  <trustInfo>\r\n')
    parts.append('    <security>\r\n')
    parts.append('      <applicationRequestMinimum>\r\n')
    parts.append('        <PermissionSet Unrestricted="true" ID="Custom" SameSite="site" />\r\n')
    parts.append('        <defaultAssemblyRequest permissionSetReference="Custom" />\r\n')
    parts.append('      </applicationRequestMinimum>\r\n')
    parts.append('      <requestedPrivileges xmlns="urn:schemas-microsoft-com:asm.v3">\r\n')
    parts.append('        <requestedExecutionLevel level="asInvoker" uiAccess="false" />\r\n')
    parts.append('      </requestedPrivileges>\r\n')
    parts.append('    </security>\r\n')
    parts.append('  </trustInfo>\r\n')
    parts.append('  <dependency>\r\n')
    parts.append('    <dependentOS>\r\n')
    parts.append('      <osVersionInfo>\r\n')
    parts.append('        <os majorVersion="5" minorVersion="1" buildNumber="2600" servicePackMajor="0" />\r\n')
    parts.append('      </osVersionInfo>\r\n')
    parts.append('    </dependentOS>\r\n')
    parts.append('  </dependency>\r\n')
    # preRequisite（GAC 依赖）
    for name, ver, pkt in PREREQ:
        parts.append('  <dependency>\r\n')
        parts.append('    <dependentAssembly dependencyType="preRequisite" allowDelayedBinding="true">\r\n')
        id_attr = '      <assemblyIdentity name="' + name + '" version="' + ver + '"'
        if pkt:
            id_attr += ' publicKeyToken="' + pkt + '" language="neutral"'
        parts.append(id_attr + ' />\r\n')
        parts.append('    </dependentAssembly>\r\n')
        parts.append('  </dependency>\r\n')
    # 主 DLL 本身（install + hash）
    # 注意：主 DLL 未强名称签名（GetPkt 实测 NO-STRONG-NAME，
    # 与 CPAHelper 主 DLL 一致），其 dependentAssembly 标识必须**不带**
    # publicKeyToken 属性——占位零 token 也是真实 token 值，会与未签名
    # 程序集的实际标识（null）不匹配，导致安装期标识校验失败。
    # ReplacePublicKeyToken 只替换清单顶层 assemblyIdentity 的占位符，
    # 不会触及此处。
    parts.append('  <dependency>\r\n')
    parts.append('    <dependentAssembly dependencyType="install" allowDelayedBinding="true" codebase="' + MAIN_DLL + '" size="' + str(file_size(main_path)) + '">\r\n')
    parts.append('      <assemblyIdentity name="' + ASSEMBLY + '" version="' + ADDIN_VERSION + '" language="neutral" processorArchitecture="msil" />\r\n')
    parts.append(hash_block(main_path))
    parts.append('    </dependentAssembly>\r\n')
    parts.append('  </dependency>\r\n')
    # 随包依赖（install + hash）
    for fname, (aname, aver, apkt) in BUNDLED.items():
        fpath = os.path.join(dirpath, fname)
        if not os.path.exists(fpath):
            print("missing bundled file: " + fpath)
            sys.exit(3)
        parts.append('  <dependency>\r\n')
        parts.append('    <dependentAssembly dependencyType="install" allowDelayedBinding="true" codebase="' + fname + '" size="' + str(file_size(fpath)) + '">\r\n')
        id_attr = '      <assemblyIdentity name="' + aname + '" version="' + aver + '"'
        if apkt:
            id_attr += ' publicKeyToken="' + apkt + '"'
        parts.append(id_attr + ' language="neutral" processorArchitecture="msil" />\r\n')
        parts.append(hash_block(fpath))
        parts.append('    </dependentAssembly>\r\n')
        parts.append('  </dependency>\r\n')
    # VSTO 加载项元数据段（vstav3/vstov4）—— vstor 的 FindApplicationManifest
    # 依赖此段识别"这是一个 VSTO 应用程序清单"，缺失时报"未找到应用程序清单"。
    # keyName 必须与注册表 HKCU\...\<宿主>\Addins\<keyName> 子键名一致。
    parts.append('  <vstav3:addIn xmlns:vstav3="urn:schemas-microsoft-com:vsta.v3">\r\n')
    parts.append('    <vstav3:entryPointsCollection>\r\n')
    parts.append('      <vstav3:entryPoints>\r\n')
    parts.append('        <vstav3:entryPoint class="' + PROJECT['namespace'] + '.ThisAddIn">\r\n')
    parts.append('          <assemblyIdentity name="' + ASSEMBLY + '" version="' + ADDIN_VERSION + '" language="neutral" processorArchitecture="msil" />\r\n')
    parts.append('        </vstav3:entryPoint>\r\n')
    parts.append('      </vstav3:entryPoints>\r\n')
    parts.append('    </vstav3:entryPointsCollection>\r\n')
    parts.append('    <vstav3:update enabled="true">\r\n')
    parts.append('      <vstav3:expiration maximumAge="7" unit="days" />\r\n')
    parts.append('    </vstav3:update>\r\n')
    parts.append('    <vstav3:application>\r\n')
    parts.append('      <vstov4:customizations xmlns:vstov4="urn:schemas-microsoft-com:vsto.v4">\r\n')
    parts.append('        <vstov4:customization>\r\n')
    parts.append('          <vstov4:appAddIn application="' + PROJECT['host'] + '" loadBehavior="3" keyName="' + PROJECT['addin_key'] + '">\r\n')
    parts.append('            <vstov4:friendlyName>' + PRODUCT + '</vstov4:friendlyName>\r\n')
    parts.append('            <vstov4:description>' + escape(PROJECT['description']) + '</vstov4:description>\r\n')
    parts.append('          </vstov4:appAddIn>\r\n')
    parts.append('        </vstov4:customization>\r\n')
    parts.append('      </vstov4:customizations>\r\n')
    parts.append('    </vstav3:application>\r\n')
    parts.append('  </vstav3:addIn>\r\n')
    parts.append('</asmv1:assembly>\r\n')
    out = os.path.join(dirpath, MAIN_DLL + ".manifest")
    with open(out, "wb") as f:
        f.write(b"\xef\xbb\xbf")
        f.write("".join(parts).encode("utf-8"))
    print("OK -> " + out)
    return out


def read_identity(app_manifest_path):
    """读取应用清单顶层 assemblyIdentity 的全部属性（签名后含真实 token）。"""
    import re
    with open(app_manifest_path, "r", encoding="utf-8-sig") as f:
        text = f.read(4096)
    m = re.search(r'<asmv1:assemblyIdentity\s+([^/>]+)/>', text)
    if not m:
        m = re.search(r'<assemblyIdentity\s+([^/>]+)/>', text)
    if not m:
        return {}
    return dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))


def gen_vsto(dirpath, app_manifest_path):
    manifest_rel = os.path.basename(app_manifest_path)
    # .vsto 的 dependentAssembly 标识必须与应用清单顶层标识完全一致
    # （包括签名后替换上的真实 publicKeyToken），否则 vstor 报
    # “部署中的引用与应用程序清单中定义的标识不匹配”。
    ident = read_identity(app_manifest_path)
    dep_line = ('      <assemblyIdentity name="' + ident.get("name", MAIN_DLL)
                + '" version="' + ident.get("version", ADDIN_VERSION) + '"')
    if ident.get("publicKeyToken"):
        dep_line += ' publicKeyToken="' + ident["publicKeyToken"] + '"'
    if ident.get("language"):
        dep_line += ' language="' + ident["language"] + '"'
    if ident.get("processorArchitecture"):
        dep_line += ' processorArchitecture="' + ident["processorArchitecture"] + '"'
    if ident.get("type"):
        dep_line += ' type="' + ident["type"] + '"'
    dep_line += ' />\r\n'
    parts = []
    parts.append('<?xml version="1.0" encoding="utf-8"?>\r\n')
    parts.append('<asmv1:assembly xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd" manifestVersion="1.0" xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" xmlns="urn:schemas-microsoft-com:asm.v2" xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" xmlns:xrml="urn:mpeg:mpeg21:2003:01-REL-R-NS" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" xmlns:dsig="http://www.w3.org/2000/09/xmldsig#" xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1" xmlns:co.v2="urn:schemas-microsoft-com:clickonce.v2">\r\n')
    parts.append('  <assemblyIdentity name="' + ASSEMBLY + '.vsto" version="' + ADDIN_VERSION + '" publicKeyToken="0000000000000000" language="neutral" processorArchitecture="msil" xmlns="urn:schemas-microsoft-com:asm.v1" />\r\n')
    parts.append('  <description asmv2:publisher="' + PRODUCT + '" asmv2:product="' + PRODUCT + '" xmlns="urn:schemas-microsoft-com:asm.v1" />\r\n')
    parts.append('  <deployment install="false" />\r\n')
    parts.append('  <compatibleFrameworks xmlns="urn:schemas-microsoft-com:clickonce.v2">\r\n')
    parts.append('    <framework targetVersion="4.8" profile="Full" supportedRuntime="4.0.30319" />\r\n')
    parts.append('  </compatibleFrameworks>\r\n')
    parts.append('  <dependency>\r\n')
    parts.append('    <dependentAssembly dependencyType="install" codebase="' + manifest_rel + '" size="' + str(file_size(app_manifest_path)) + '">\r\n')
    parts.append(dep_line)
    parts.append(hash_block(app_manifest_path))
    parts.append('    </dependentAssembly>\r\n')
    parts.append('  </dependency>\r\n')
    parts.append('</asmv1:assembly>\r\n')
    out = os.path.join(dirpath, ASSEMBLY + ".vsto")
    with open(out, "wb") as f:
        f.write(b"\xef\xbb\xbf")
        f.write("".join(parts).encode("utf-8"))
    print("OK -> " + out)


def main():
    # 确认部署文件齐备
    for f in [MAIN_DLL] + list(BUNDLED.keys()):
        p = os.path.join(BUILD, f)
        if not os.path.exists(p):
            print("missing: " + p + "  (bundled deps should be copied to build/ first)")
            sys.exit(3)
    if "--vsto-only" in sys.argv:
        # 只重生成 .vsto（应用清单已签名，重算其哈希）
        m = os.path.join(BUILD, MAIN_DLL + ".manifest")
        if not os.path.exists(m):
            print("missing: " + m)
            sys.exit(3)
        gen_vsto(BUILD, m)
        return
    m = gen_app_manifest(BUILD)
    gen_vsto(BUILD, m)


if __name__ == "__main__":
    main()
