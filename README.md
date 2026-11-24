# 费用报销单自动生成工具

[![Release](https://img.shields.io/github/v/release/liutangsheng/reimburse)](https://github.com/liutangsheng/reimburse/releases/latest)
[![Build](https://github.com/liutangsheng/reimburse/actions/workflows/build.yml/badge.svg)](https://github.com/liutangsheng/reimburse/actions/workflows/build.yml)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)](https://github.com/liutangsheng/reimburse/releases/latest)
[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB)](https://www.python.org/)

## 一、免安装打包版使用方式（推荐）

无需安装 Python 或 Tesseract，直接从 [Releases 页面](https://github.com/liutangsheng/reimburse/releases/latest) 下载对应系统的压缩包解压即用：

1. **Windows**：解压 `费用报销单生成工具_Windows绿色版.zip` 到任意文件夹，双击 **`费用报销单生成工具.exe`**。
2. **macOS**：解压 `费用报销单生成工具_macOS版.zip`，双击 **`费用报销单生成工具.app`**（首次打开若提示未签名，右键 →「打开」即可）。
3. 浏览器会自动打开操作页面，点击「选择文件夹」选中报销材料目录，点击「开始生成报销单 PDF」即可。
4. 使用完毕后，直接关闭终端窗口（或退出应用）即可结束。

> 如需自行重新打包 Windows 绿色版，请确保 `vendor/tesseract/` 下包含 `tesseract.exe`、`*.dll` 及 `tessdata/eng.traineddata`、`tessdata/chi_sim.traineddata`，然后运行 `python build_windows.py`。

---

## 二、报销目录结构要求

```text
示例公司+张三+2026.11.20/        # 建议按「公司+姓名+YYYY.MM.DD」命名
├── signatrue.jpg                # 报销人签名图片（支持 signatrue.* 或 signature.*）
├── 餐费/                        # 子目录名即报销单「用途」，可自由增减类目
│   ├── 订单/                    # 订单截图（jpg/jpeg/png/webp/bmp）或订单 PDF
│   └── 发票/                    # 电子发票（支持 PDF 或图片）
└── 交通费/                      # 可继续添加 住宿费/、办公用品/ 等其他报销类目
    ├── 订单/                    # 如滴滴出行行程报销单.pdf 或订单截图
    └── 发票/                    # 电子发票（PDF 或图片）
```

> 同时兼容旧版根目录 `orderImage/`、`发票/`、`traffic/` 结构；生成的 PDF 页面顺序固定为：**报销单 → 各类目订单 → 各类目发票**。

---

## 三、源码环境准备

1. 安装本机 `tesseract`（用于识别订单截图及图片格式发票）：
   - **macOS**：`brew install tesseract`
   - **Windows**：安装 [Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki)（默认路径 `C:\Program Files\Tesseract-OCR\tesseract.exe` 会被自动识别）
2. 安装 Python 依赖：
   ```bash
   pip install -r requirements.txt
   ```

---

## 四、源码使用方式

### 1. 网页界面（推荐）
直接运行 `web_ui.py`（或无参数运行 `app.py`）：
```bash
python web_ui.py
```
浏览器会自动打开操作页面，按提示选择报销材料文件夹即可生成 PDF。

### 2. 命令行直接生成
```bash
# 处理指定报销目录
python main.py "示例公司+张三+2026.11.20"

# 自定义报销部门与输出文件名（加上 --docx 可同步导出 Word）
python main.py "示例公司+张三+2026.11.20" --dept 示例部门 -o "报销单.pdf" --docx
```

### 参数说明
| 参数 | 说明 | 默认值 |
| --- | --- | --- |
| `dir` | 报销材料目录路径 | `示例公司+张三+2026.11.20` |
| `-o, --output` | 输出 PDF 路径 | `<目录名>_生成.pdf` |
| `--dept` | 报销部门 | `Bida` |
| `--docx` | 同步输出可编辑 Word（`.docx`） | 关闭 |
