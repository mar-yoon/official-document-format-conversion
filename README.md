# 公文格式助手

面向 Mac 和 Windows 的离线 Word 公文排版程序。桌面版不需要浏览器、服务器或小程序账号。选中文档后，程序会自动在原文件旁生成“原文件名_标准格式.docx”，保留原件；也可另存到其他位置。

## 直接使用

- GitHub Actions 的 `Build desktop downloads` 工作流会构建 Windows、Apple Silicon Mac 和 Intel Mac 的 ZIP 安装包。
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
- “国企公文（培训材料版）”依据《公文格式培训》第 2—14 页：正文仿宋_GB2312 三号、固定行距 28 磅、两端对齐；标题方正小标宋简体二号；一级至四级标题分别使用黑体、楷体_GB2312、仿宋_GB2312；英文数字采用 Times New Roman，表格文字采用宋体四号。图片独占段落按居中单倍行距处理，已有页码字段的字体调为宋体四号。A4 纸张沿用公文模板设置。

目前仅处理 `.docx`。不会补写缺失的机关名称、发文字号、日期或印章，也不会自动生成单双页页码。标题和段落类型由规则识别，正式发文前需复核红头、印章、页码、附件、表格和跨页布局。国企模板不会自动插入附件分节符或将图片表格转成可编辑表格。

## 此前的浏览器及小程序原型

`app.py` 和 `static/` 是浏览器调试版，须先运行 `python app.py` 才能打开 <http://127.0.0.1:8765>。`miniprogram/` 是小程序原型，发布仍需要 AppID、HTTPS 域名和微信平台配置；详见 `上线准备.md`。桌面版无需这些步骤。
