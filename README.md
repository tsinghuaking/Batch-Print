简体中文 | [English](README_en.md)

<div align="center">
  <img src="docs/logo.png" alt="Batch Print Tool">
</div>

<div align="center">

![License](https://img.shields.io/badge/License-MIT-green.svg)
![Python](https://img.shields.io/badge/Python-3.14-blue.svg)
![OS](https://img.shields.io/badge/OS-Windows-blue.svg)
[![Release](https://img.shields.io/github/v/release/tsinghuaking/Batch-Print?label=Release)](https://github.com/tsinghuaking/Batch-Print/releases)
[![Download](https://img.shields.io/github/downloads/tsinghuaking/Batch-Print/total?label=Downloads)](https://github.com/tsinghuaking/Batch-Print/releases)

</div>

## 项目简介

**批量打印工具（Batch Print Tool）** 是一款运行在 Windows 上的本地小工具，用来**批量打印 Word / Excel / PDF**，
并支持逐文件设置 **双面 / 单面、打印份数、彩色 / 黑白、打印页码范围**。文件只在本机处理，全程不上传任何网络。

> 为什么需要它：Windows 自带右键「打印」无法逐项设置细节，而这个工具把它们集中到一个网页里，选好一次点「开始打印」即可。

主要特性：

- 📄 **支持 PDF / Word(.doc/.docx) / Excel(.xls/.xlsx)** —— Word、Excel 会先由本机 LibreOffice 转成 PDF 再打印
- 🖨️ **逐文件精细设置**：双面（长边翻转 / 短边翻转）、份数、彩色 / 黑白、页码范围（如 `1-3`、`1,3,5-7`）
- ⚫ **默认黑白**：新上传的文件默认按黑白打印，避免误打彩色浪费墨水
- 💾 **记住默认打印机**：界面点「设为默认」后，下次打开自动选中，无需每次下拉选择
- ✅ **打印状态回执**：每个文件发送成功后，对应行状态变绿色「已完成」；失败显示红色「失败」，下方日志有详情
- 🚀 **单文件绿色 exe**：基于 Python 标准库（无第三方 Web 框架），SumatraPDF 打印引擎直接**内嵌进 exe**，**双击即启动并自动打开浏览器**；也可把 `bin/` 放到 exe 同目录改用外置引擎以减小体积

## 预览

<div align="center">
  <img src="docs/preview.png" alt="界面预览">
</div>

## 快速开始（普通用户，推荐）

1. 到 [Releases](https://github.com/tsinghuaking/Batch-Print/releases) 下载 **`BatchPrint.exe`**（单文件，自带 SumatraPDF 引擎，约 38MB）。
2. **双击** `BatchPrint.exe`：
   - 会弹出一个黑色控制台窗口（不要关闭，关掉就停止服务）；
   - 自动打开默认浏览器到 `http://127.0.0.1:5001`。
3. 用完关掉那个黑色窗口即可。

> 如果 `5001` 端口已被占用（之前已开过一个），程序会直接打开浏览器复用已有服务，不会重复启动。

### 想让 exe 更小 / 单独更新打印引擎？

使用「外置版」：把 `bin/` 文件夹（含 `SumatraPDF.exe` 及 `libmupdf.dll`、`PdfFilter.dll`、`PdfPreview.dll`）
与 `BatchPrint.exe` 放在同一目录。此时 exe 会优先使用外置引擎，自身可小到约 10MB。
仓库里的 `bin/` 目录已随源码提供，直接取用即可。

## 使用说明

1. **选打印机**：顶部下拉框选择真实打印机；点「刷新」重新读取；选好常用打印机后点「设为默认」。
2. **上传文件**：点中间虚线框或把文件拖入，可一次多选 PDF / Word / Excel。
3. **设参数**：
   - 全部一样：在「全局设置」行选好后点「应用到全部」；
   - 单个不同：在表格里每行单独改；
   - 页码写法：`1-3`（第 1~3 页）、`1,3,5-7`（指定页），留空 = 全部页。
4. **开始打印**：点「开始打印」，每个文件在表格里实时显示「已完成 / 失败」，下方日志有详情。

## 从源码运行（开发者）

```bash
pip install -r requirements.txt
python app.py
# 浏览器访问 http://127.0.0.1:5001
```

## 打包成 exe

```bash
pip install pyinstaller
python build_exe.py
```

产物在 `dist/批量打印工具.exe`。脚本已用 `--add-data` 把 `bin/`（SumatraPDF 引擎）与 `templates/` 全部打进 exe，
发布时只需这一个文件（自包含、双击即用）。

> 注意：GitHub 网页拖拽上传附件有 25MB 限制；发布时用命令行
> `gh release create vX.Y "dist/批量打印工具.exe"` 上传（走 uploads API，支持到 2GB）。

## 目录结构

```
批量打印工具/
├── app.py              # 网页服务（Python 标准库 http.server，无第三方依赖）
├── print_core.py       # 打印核心：枚举打印机 / Office 转 PDF / 调 SumatraPDF
├── templates/
│   └── index.html      # 网页界面
├── bin/                # 打印引擎 SumatraPDF（随仓库提供，也可外置）
├── requirements.txt
├── build_exe.py        # 打包脚本
├── LICENSE
└── .gitignore
```

## 常见问题

**1. 打印没出纸 / 提示失败？**
- 确认选的是**真实打印机**。「Microsoft Print to PDF / WPS PDF / Adobe PDF」是虚拟打印机，会弹「另存为」窗口，不适合静默批量打印。
- 去 Windows「设置 → 打印机」看该打印机是否处于「已暂停」或「脱机」状态。

**2. 电脑里那台 "WPS PDF" 名字带乱码？**
那是 WPS 安装时写入系统的坏名字（非本工具 bug），且它是虚拟打印机，建议别用来真打。

**3. 为什么要先装 LibreOffice 才能打 Word / Excel？**
Word / Excel 没有系统级命令行打印接口，本工具借 LibreOffice 把它们转成 PDF 再交给 SumatraPDF 打印。
首次转换某个文件可能要等 1~2 分钟（LibreOffice 首次启动慢），之后变快。若只打 PDF 则无需安装。

**4. 上传的文件存在哪？**
临时存放在 `%USERPROFILE%/.batch_print/uploads/`，打印完不会自动删除，可手动清理。

## 第三方组件许可

- 打印引擎 **SumatraPDF**（内嵌于发布版 exe，也可外置）采用 GPLv3，源码见
  <https://github.com/sumatrapdfreader/sumatrapdf>
- 本项目仅用 Python 标准库（无 Flask 等第三方 Web 框架），自身以 MIT 许可发布（见 [LICENSE](LICENSE)）

## 许可证

本项目以 [MIT License](LICENSE) 发布。
