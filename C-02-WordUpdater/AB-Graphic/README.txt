图标放置与规格说明（本文件由 AA-Program 配置器自动生成）
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
  带透明像素的图会先把透明部分合成到 #F5F6F7 底色上（可在配置区 ICON_BACKGROUND 改）。
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
