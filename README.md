# 费用报销单自动生成工具

## 一、报销目录结构要求

```text
东百+刘唐胜+2026.11.20/          # 建议按「公司+姓名+YYYY.MM.DD」命名
├── signatrue.jpg                # 报销人签名图片（支持 signatrue.* 或 signature.*）
├── orderImage/                  # 餐饮订单截图（jpg/jpeg/png/webp/bmp）
│   └── 微信图片_xxx.jpg
├── 发票/                        # 餐饮电子发票（支持 PDF 或图片 jpg/jpeg/png/webp/bmp）
│   └── dzfp_xxx.pdf
└── traffic/                     # 交通报销目录（可选，自动区分行程单与交通发票，发票支持 PDF 或图片）
    ├── 滴滴出行行程报销单.pdf
    └── 滴滴电子发票.pdf
```

---

## 二、环境准备

1. 安装本机 `tesseract`（用于识别订单截图及图片格式发票）：
   - **macOS**：`brew install tesseract`
   - **Windows**：安装 [Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki)（默认路径 `C:\Program Files\Tesseract-OCR\tesseract.exe` 会被自动识别）
2. 安装 Python 依赖：
   ```bash
   pip install -r requirements.txt
   ```

---

## 三、源码使用方式

### 1. 网页界面（推荐）
直接运行 `web_ui.py`（或无参数运行 `app.py`）：
```bash
python web_ui.py
```
浏览器会自动打开操作页面，按提示选择报销材料文件夹即可生成 PDF。

### 2. 命令行直接生成
```bash
# 处理指定报销目录
python main.py "东百+刘唐胜+2026.11.20"

# 自定义报销部门与输出文件名
python main.py "东百+刘唐胜+2026.11.20" --dept Bida -o "报销单.pdf"
```

### 参数说明
| 参数 | 说明 | 默认值 |
| --- | --- | --- |
| `dir` | 报销材料目录路径 | `东百+刘唐胜+2026.11.20` |
| `-o, --output` | 输出 PDF 路径 | `<目录名>_生成.pdf` |
| `--dept` | 报销部门 | `Bida` |
