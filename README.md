# 批量打印工具

一个本地小工具，用来**批量打印 Word / Excel / PDF**，并且能逐个文件设置：
**双面 / 单面、打印份数、彩色 / 黑白、打印页码范围**。文件只在本机处理，不上传任何网络。

## 功能特性

- 📄 支持 PDF / Word(.doc/.docx) / Excel(.xls/.xlsx)
- 🖨️ 逐个文件设置：双面(长边/短边)、份数、彩色/黑白、页码范围
- ⚫ **默认黑白**（新上传文件默认按黑白打印，避免误打彩色浪费）
- 💾 **记住默认打印机**：在界面点“设为默认”后，下次打开自动选中，不用每次下拉选
- ✅ **打印状态回执**：每个文件打印发送成功后，对应行状态变成绿色“已完成”；失败显示红色“失败”
- 🚀 **自包含 exe**：仅用 Python 标准库、无第三方依赖，SumatraPDF 引擎**直接内嵌进 exe**，**双击即启动并自动打开浏览器**；也可把 `bin/` 放到 exe 同目录改用外置引擎（减小单文件体积）

## 快速开始（普通用户，推荐）

1. 到 [Releases](../../releases) 下载 `批量打印工具.exe`（已内嵌 SumatraPDF 引擎，单文件约 70MB）。
2. **双击** `批量打印工具.exe`：
   - 会弹出一个黑色控制台窗口（不要关，关了就停服务）；
   - 自动打开默认浏览器到 `http://127.0.0.1:5001`。
3. 用完关掉那个黑色窗口即可。

> **想要更小体积 / 单独更新引擎？** 用「外置版」：把 `SumatraPDF.exe` 及 `libmupdf.dll`、`PdfFilter.dll`、`PdfPreview.dll`
> 放进 `bin/` 文件夹，与 `批量打印工具.exe` 放在同一目录。此时 exe 会优先用外置引擎，exe 本身可小到约 10MB。
> 直接下载仓库 `bin/` 目录或自行放置均可。

> **上传到 GitHub Releases 注意**：单个文件超过 25MB 时，**不要用网页拖拽上传**（会报
> "Yowza, that's a big file" 25MB 限制）。请用命令行 `gh release create … "dist/批量打印工具.exe"` 上传，
> CLI 走 GitHub uploads API，支持到 2GB，不受网页 25MB 限制。

> 如果 5001 端口被占用（之前已开过一个），程序会直接打开浏览器复用，不会重复启动。

## 从源码运行（开发者）

```bash
pip install -r requirements.txt
python app.py
# 浏览器访问 http://127.0.0.1:5001
```

## 怎么用

1. **选打印机**：顶部下拉框选真实打印机；点“刷新”重新读取；选好常用打印机后可点“设为默认”。
2. **上传文件**：点中间虚线框或拖拽，可一次多选 PDF / Word / Excel。
3. **设参数**：
   - 全部一样：用“全局设置”那一行选好，点“应用到全部”。
   - 单个不同：在表格里每行单独改。
   - 页码写法：`1-3`（第1~3页）、`1,3,5-7`（指定页），留空=全部页。
4. **开始打印**：点“开始打印”，每个文件在表格里实时显示“已完成 / 失败”，下方日志有详情。

## 打包成 exe

```bash
pip install pyinstaller
python build_exe.py
```

产物在 `dist/批量打印工具.exe`。脚本已用 `--add-data` 把 `bin/`（SumatraPDF 引擎）
与 `templates/` 全部打进 exe，发布时只需这一个 exe 文件（自包含、双击即用）。

## 目录结构

```
批量打印工具/
├── app.py              # 网页服务（Python 标准库 http.server，无第三方依赖）
├── print_core.py       # 打印核心：枚举打印机 / Office 转 PDF / 调 SumatraPDF
├── templates/
│   └── index.html      # 网页界面
├── bin/                # 打印引擎 SumatraPDF（随 release 提供，或自行放入）
├── requirements.txt
├── build_exe.py        # 打包脚本
├── LICENSE
└── .gitignore
```

## 注意事项

- **Word / Excel 会先由 LibreOffice 转成 PDF 再打印**（需本机已装 LibreOffice）。
  首次转换某个文件可能要等 1~2 分钟（LibreOffice 首次启动慢），之后变快。
- 选**真实打印机**。“Microsoft Print to PDF / WPS PDF / Adobe PDF”是虚拟打印机，
  会弹“另存为”窗口，不适合静默批量打印。
- 你电脑里那台 "WPS PDF" 在系统里名字带乱码前缀（WPS 安装时写入的，非本工具 bug），
  它是虚拟打印机，建议别用来真打。
- 上传文件临时存放在 `~/.batch_print/uploads/`，打印完不会自动删，可手动清理。

## 第三方组件许可

- 打印引擎 **SumatraPDF**（内嵌于发布版 exe，也可外置）采用 GPLv3，源码见
  https://github.com/sumatrapdfreader/sumatrapdf
- 本项目仅用 Python 标准库（无 Flask 等第三方 Web 框架），本身以 MIT 许可发布（见 LICENSE）
