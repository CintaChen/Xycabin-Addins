# WPS Office 加载项与 XyCabin 站点

本项目包含三个相互关联的子项目：两个面向 WPS Office 的 VSTO 加载项，以及一个用于展示与分发这两款加载项的离线站点生成器。

三者共享同一套文档体系与交付门禁，因此以单仓库（monorepo）形式维护。

## 项目组成

| 目录 | 类型 | 宿主 | 功能概述 |
|---|---|---|---|
| `C-00-Shared` | 共用源码与构建实现 | 编译到两个加载项 | 单一公共实现、窗口与网格组件、配置生成、PDF 工具、安装器及验证工具 |
| `C-01-ExcelIterate` | VSTO 加载项 | WPS 表格 / Excel | 公式引用与从属追踪、一键公式封装与数值计算、按列控制显示、文件批量整理与重命名、OCR 文本提取 |
| `C-02-WordUpdater` | VSTO 加载项 | WPS 文字 / Word | 将 Excel 区域粘贴为表格链接、对选中范围的链接批量更新 |
| `C-03-Website·Xycabin` | 站点生成器 | 浏览器 | 生成单文件离线站点，展示两款加载项的介绍、使用指南、版本与下载 |

两个加载项通过 `project.json` 中的 `@shared/` 引用直接编译公共源码，产品脚本只保留传入产品根目录的薄入口，不增加运行时共享 DLL。`verify_shared.py` 校验公共实现唯一、引用有效、入口纯转发及必要依赖一致；`verify_devlog.py` 校验统一文档与门禁。开发入口见 [B0-Entry.md](C-04-Devlog/B0-Entry.md)。

按界面功能定位代码请打开 [B4-Code Index.md](<C-04-Devlog/B4-Code Index.md>)：两个产品各自的 `AC-Original Code` 按选项卡、分组、菜单和末级功能归档，索引提供中文功能树、英文目录、实际回调和共用依赖。新增或迁移功能后运行 `python C-00-Shared/feature_index.py --write` 更新索引，省略 `--write` 执行只读校验。

## 技术栈

- **加载项**：C# 5（由 .NET Framework 自带的 `csc.exe` 编译），VSTO vstov4 清单，无 Visual Studio 工程文件；独立 PDF 工具按锁定版本和 SHA-256 恢复依赖。
- **构建与验证**：Python 3，仅使用标准库（无 `requirements.txt`）。
- **站点**：由 Python 模板渲染为单个 HTML 文件，内嵌 CSS、JavaScript 与图片资源，可离线打开。

## 目录结构

```
.
├── C-00-Shared/          公共源码、资源与构建实现（两个加载项直接引用）
├── C-01-ExcelIterate/
│   ├── AA-Program/        配置器、启动器与安装包
│   ├── AB-Graphic/        界面截图、功能区图标与示意 SVG
│   ├── AC-Original Code/  按菜单层级归类的功能入口、算法和弹窗
│   ├── src/               加载框架、生成文件、配置、安装器与测试工具
│   ├── build/             构建产物、缓存及验证证据（未入库）
│   └── lib/               微软版权互操作程序集（未入库）
├── C-02-WordUpdater/      结构同上
├── C-03-Website·Xycabin/
│   ├── AA-Program/        站点配置器、启动器与生成的网页
│   └── src/               模板、控制器、资产
├── C-04-Devlog/           唯一文档维护位置（B0、A-UserGuide、B1、B2、B3、B4）
├── verify_devlog.py       文档结构、概览、审计与规则注册表校验
├── verify_shared.py       公共实现、源引用和产品薄入口校验
└── AGENTS.md              仅指向 B0-Entry.md
```

## 构建

### 环境要求

1. Windows 10 / 11
2. .NET Framework 4.x（提供 `csc.exe`，系统自带）
3. Python 3.9 或更高版本
4. WPS Office —— 仅在运行与验收时需要，编译不需要

### 补回互操作程序集

仓库不包含 `lib/` 下的微软版权程序集（`Microsoft.Office.Tools.*.dll`、`Microsoft.VisualStudio.Tools.Applications.Runtime.dll` 等），因为其再分发受微软许可条款约束。首次构建前，请从本机的 VSTO 运行时或 Visual Studio Tools for Office PIA 目录将这些程序集复制到对应项目的 `lib/` 目录。

缺少任一会导致编译报错，错误信息会指出具体文件名。

### 编译

请保留 `C-00-Shared`、相应加载项、`C-03-Website·Xycabin` 和 `C-04-Devlog` 的同级目录关系。公共源码迁移后，单独下载一个加载项目录不足以构建。

```bash
# 仅编译 DLL 与依赖，不签名、不封装安装包
python C-01-ExcelIterate/build.py
python C-02-WordUpdater/build.py
```

### 完整重建

```bash
python C-01-ExcelIterate/rebuild.py --background --mode full
python C-02-WordUpdater/rebuild.py --background --mode full
python "C-03-Website·Xycabin/rebuild.py" --background --mode full
```

或双击各项目 `AA-Program/` 下的 `一键重建并打包.bat`。

完整重建依次执行：写入配置 → 编译 → 生成并签名清单 → 封装安装包 → 校验。`--mode incremental` 会复用 `build/incremental-cache/` 中的编译缓存。

### 关于代码签名

完整重建的最后阶段会使用 Authenticode 证书为 DLL、应用清单与部署清单签名。签名身份由各项目 `AA-Program` 下的配置器决定：Excel-Iterate 为 `C1-Exceliterate_Configure.py`，WordUpdater 为 `C2-Wordupdater_Configure.py`，其中的 `SIGNING_CERT_THUMBPRINT` 是唯一来源。

**自行构建时，该指纹指向的证书在你的机器上并不存在，签名步骤会失败。** 有两种处理方式：

1. 在对应项目的配置器中将 `SIGNING_CERT_THUMBPRINT` 改为你自有证书的指纹，并确保证书位于当前用户的个人证书存储区且带私钥；
2. 跳过完整重建，改用 `build.py` 编译出 DLL，再手动安装 VSTO 加载项用于开发调试。

原项目的发布证书为自签名证书，未加入任何公共信任根。因此安装官方构建的安装包时，Windows 会提示「未知的发行者」，这属于预期行为。

## 未纳入版本管理的内容

| 路径 | 原因 |
|---|---|
| `build/`、`stage/` | 构建快照与中间产物，体积可达数百 MB，可由源码重建 |
| `.workbuddy/`、`.ai-memory/` | AI 辅助开发的本地工作记录，含本机路径与过程性内容 |
| `**/lib/*.dll` | 微软版权程序集，再分发受许可限制 |
| `AA-Program/*_Setup.exe` | 发布产物，通过 Releases 分发 |
| `C-03-Website·Xycabin/AA-Program/XyCabin.html` | 由模板渲染生成的产物 |
| `*.cer`、`*.pfx`、`*.snk` | 证书与私钥材料 |

## 源码提交、推送与 Release 附件

源码、公共目录、测试和统一文档通过 Git 提交；两个安装包与独立 HTML 按上述忽略规则保留在本地，通过 GitHub Releases 分发。提交代码及 `git push` 不会上传或替换 Release 附件，两者必须分开执行。

远端为 `https://github.com/Xymircai/Xycabin-Addins.git`，长期分支 `main`；`restructure/20261019` 为历史分支，不再合入。

### 提交前核对

```bash
python verify_devlog.py                 # 六份文档、版本表与门禁注册表
python verify_shared.py                 # 公共实现唯一、源引用与功能映射
python C-00-Shared/feature_index.py     # 只读校验代码索引
git status --short                      # 确认改动范围
```

三条校验全部 PASS 再提交。迁移源码、增删菜单或改动回调后，先运行 `python C-00-Shared/feature_index.py --write` 重建 `C-04-Devlog/B4-Code Index.md`，否则 `verify_shared.py` 会以「B4 代码索引与实际功能映射不一致」失败。

完整重建期间不要改文档、共享源码与内嵌图片；如有改动，等本轮构建结束后重跑，保证构建前后输入指纹一致。

### 分批提交

一次提交只表达一件事，便于回溯与单独回退：

```bash
git add C-01-ExcelIterate/AA-Program/C1-Exceliterate_Configure.py
git commit -m "feat(excel): 审计工具箱子项更名为「表头匹配」"

git add C-04-Devlog README.md
git commit -m "docs(release): 补充提交与附件更新流程"
```

提交信息约定 `类型(范围): 摘要`：类型取 `feat`、`fix`、`refactor`、`docs`、`chore`、`test`；范围取 `excel`、`word`、`site`、`shared`，跨产品时省略。正文写清动机、影响面与已执行的验证。

- 不要用 `git add -A` 或 `git add .` 一次提交全部改动，容易把无关修改混进同一提交；
- 提交前用 `git diff --cached` 复核暂存内容，确认没有带入证书、私钥或本机绝对路径；
- `build/`、`stage/`、`AA-Program/*_Setup.exe`、`XyCabin.html`、`lib/*.dll` 与证书已被 `.gitignore` 排除，`git status` 看不到它们属正常。

### 推送

```bash
git fetch origin
git rebase origin/main        # 远端有新提交时先变基，保持线性历史
git push origin main
git log --oneline -3          # 与 GitHub 网页上的提交号比对
```

推送前先 `git fetch`；被拒时用 `git pull --rebase` 处理，不要用 `git push --force` 覆盖他人提交。推送成功后 GitHub 代码页即显示最新提交。

### 更新 Release 附件

版本号唯一来源是各项目配置器的 `INSTALLER_VERSION`，只有明确要求升版时才修改。版本不变而内容更新时，用同名附件覆盖现有 Release：

```bash
# 1) 确认版本号与三份产物均已按当前源码重建
grep -n "INSTALLER_VERSION =" C-01-ExcelIterate/AA-Program/C1-Exceliterate_Configure.py
ls -l C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe \
      C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe \
      "C-03-Website·Xycabin/AA-Program/XyCabin.html"

# 2) 覆盖上传；同名附件必须加 --clobber，否则报 already exists
gh release upload v0.9.9.3 \
  C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe \
  C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe \
  "C-03-Website·Xycabin/AA-Program/XyCabin.html" --clobber

# 3) 重新下载，核对远端字节与本地一致
gh release download v0.9.9.3 --dir build/release-verify --clobber
sha256sum build/release-verify/*.exe build/release-verify/*.html
sha256sum C-01-ExcelIterate/AA-Program/Excel_Iterate_Setup.exe \
          C-02-WordUpdater/AA-Program/WordUpdater_Setup.exe \
          "C-03-Website·Xycabin/AA-Program/XyCabin.html"
```

1. 三份文件必须同批重建、同批上传。独立 HTML 内嵌两个安装包，其正文又与两 DLL 的内嵌网页一致；只替换其中一件会造成网页与安装包错配。
2. 上传前确认工作区没有未提交的源码改动，避免附件比仓库代码新。
3. HTML 内嵌的两个 EXE 必须与待上传安装包逐字节一致；移除安装包数据区后的页面须与两 DLL 的内嵌网页一致。
4. 更新附件不会移动已有标签，GitHub 自动生成的源码压缩包仍取自标签所指提交。版本不变时「同版本不同字节」的追溯依据是本轮提交号与附件 SHA-256，记录在 [B1 交付证据](C-04-Devlog/B1-Project-Files.md)。
5. 需要正式换版时，先在配置器升 `INSTALLER_VERSION`，重建、提交、推送，再新建标签与 Release，不要复用旧标签。
6. 上传超时可直接重跑 `gh release upload`，`--clobber` 会覆盖同名附件；中断可能留下未完成的附件，用 `gh release view v0.9.9.3` 核对后再补传。

## 免责声明

本项目为个人开发作品，**与金山办公（WPS Office 的开发方）无隶属或背书关系**，亦非 WPS Office 官方组件。

加载项通过 VSTO 接口与宿主进程交互，会读取并修改当前文档内容。请在测试文档中先行验证，并自行承担使用风险。作者不对因使用本软件导致的数据丢失、文档损坏或业务中断承担责任。

## 许可证

本项目采用 [Mozilla Public License 2.0](LICENSE)（MPL-2.0）授权。

MPL-2.0 是**文件级** copyleft 许可证，其边界与常见许可证不同，使用前请留意：

- 允许商业使用、修改与再分发，也允许收费分发。
- 你修改过的**源文件**，在分发时必须继续以 MPL-2.0 公开。
- 但 MPL-2.0 不要求你新增的文件或与之组合的其他代码开源，可以与其他许可证（含专有）代码组合成一个更大的作品。
- 含明确的专利授权条款；不授予商标使用权。
- 本项目未声明「Incompatible With Secondary Licenses」，因此按 MPL-2.0 默认规则，可与 GPL-2.0 或更高版本进行二次组合。

源文件的许可声明以仓库根目录的 `LICENSE` 文件为准，未在每个文件头部逐一附加 Exhibit A 通知。若你希望在自己维护的分支中补全文件头，可参考 `LICENSE` 末尾的 Exhibit A 模板。

## 问题反馈

请通过本仓库的 Issues 提交缺陷报告或功能建议。提交时请说明 WPS Office 版本、Windows 版本，以及可复现的操作步骤。

PDF 工具适配 0.9.9.3：使用与已知问题 PDF-001 见 [统一使用说明](C-04-Devlog/A-UserGuide.md)，实现边界见 [共享开发经验](C-04-Devlog/B2-DevExperience.md)。PDF-001 待复现、未修复。
