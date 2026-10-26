# 公文格式助手

面向 Mac 和 Windows 的离线 Word 公文排版程序。桌面版不需要浏览器、服务器或小程序账号。原始文件留在本机，用户自行选择输出位置。

## 直接使用

- 当前已生成 Apple Silicon Mac 测试包：`dist/公文格式助手-macOS-arm64.zip`。解压后双击 `.app`。
- Windows `.exe` 需在 Windows 环境构建。仓库已有跨平台构建脚本和 GitHub Actions 工作流；连接 GitHub 仓库并运行 `Build desktop downloads` 后，可下载 Windows 和 Mac 的 ZIP 构建产物。
- 也可在有 Python 和 Tkinter 的电脑上运行 `python desktop.py`。

下载包中的 `使用说明.md` 面向普通用户。公开分发前，Mac 版仍需开发者签名和公证，Windows 版建议代码签名。

## 构建桌面安装包

在目标操作系统上安装 Python 3.12 和 PyInstaller 6.22.3，然后运行：

```bash
python -m pip install pyinstaller==6.22.3
python build_desktop.py
```

Mac 生成 `.app` 及 ZIP；Windows 生成单文件 `.exe` 及 ZIP。PyInstaller 不支持从 Mac 直接交叉编译 Windows 程序。

## 排版范围

- 按 [GB/T 9704—2012](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=F3CC9BEF482524C895FDA7A08BB4A70E) 设置 A4、37 mm 上白边、28 mm 左白边、156 × 225 mm 版心。
- 为标题、主送机关、正文和常见四级层次序号设置字体、字号、对齐及首行缩进。
- 保留文字、表格、图片及其他文档部件，并列出需要复核的项目。
- 模板规则在 `templates/*.json`；高校论文格式可通过新模板和分类器扩展。

目前仅处理 `.docx`。不会补写缺失的机关名称、发文字号、日期或印章，也不会自动生成单双页页码。标题和段落类型由规则识别，正式发文前需复核红头、印章、页码、附件、表格和跨页布局。

## 此前的浏览器及小程序原型

`app.py` 和 `static/` 是浏览器调试版，须先运行 `python app.py` 才能打开 <http://127.0.0.1:8765>。`miniprogram/` 是小程序原型，发布仍需要 AppID、HTTPS 域名和微信平台配置；详见 `上线准备.md`。桌面版无需这些步骤。
