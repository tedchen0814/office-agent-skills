---
name: file-toolkit
description: >-
  用 Python 處理 Word／Excel／PowerPoint／PDF 檔案的做法手冊：批次套印 Word 範本、Excel 報表計算與依欄位拆檔、PDF 合併／抽頁／加浮水印／轉圖片、舊 PDF／Office 檔轉 Markdown、連結轉 QR Code。使用者要批次處理文件、合併或拆分 PDF、把檔案轉成文字、產生 QR Code，且環境沒有更專用的 docx／xlsx／pdf／pptx skill 時使用。只是要做簡報時改用 html-slide-builder、soil-html-deck 或 soil-teaching-deck。
---

# 檔案處理工具手冊（file-toolkit）

用現成的 Python 套件處理日常文件。先確認需求，再照下表選套件寫一支小腳本執行；**不要手動逐檔操作**，批次處理才是重點。

## 先確認

1. 輸入檔在哪、要輸出到哪（輸出另存新檔，**不要覆寫原檔**）。
2. 批次的規則是什麼（依哪個欄位、檔名怎麼取）。
3. 檔案內容若含個資或公司內部資訊，只在本機處理，不上傳到任何外部服務。

## 套件對照

| 工作 | 套件 | 備註 |
|---|---|---|
| Word 讀寫、套印範本 | `python-docx` | 只支援 `.docx`，舊 `.doc` 要先另存 |
| Excel 讀寫、公式、格式 | `openpyxl` | `.xlsx`；要算統計可臨時裝 `pandas` |
| PowerPoint 讀寫 | `python-pptx` | 新簡報優先用 html-slide-builder 等 skill |
| PDF 合併／拆分／抽頁／加浮水印 | `pypdf`＋`reportlab` | 浮水印用 reportlab 畫一頁再疊上去 |
| PDF 轉圖片、擷取頁面、取出文字 | `PyMuPDF`（`import pymupdf`） | 不要用舊名 `fitz`，會有警告 |
| 任意文件轉 Markdown（餵 AI 前處理） | `markitdown` | 支援 PDF／Word／PowerPoint／Excel |
| 連結轉 QR Code | `qrcode[pil]` | 輸出 PNG，可貼進簡報或文件 |
| 畫圖表 | `matplotlib` | **預設沒裝**，需要時再 `pip install matplotlib` |

套件缺少時直接 `pip install` 即可。Windows 若 `python` 是 Microsoft Store 空殼，改用 `py -3`。

## 做法範例

### 1. 批次套印 Word 範本（名單 Excel → 每人一份）

範本裡用 `{{name}}`、`{{score}}` 這類標記。標記若被 Word 拆成多個 run，要先合併成一個 run 再取代（下面的寫法已處理）。

```python
import openpyxl
from docx import Document

rows = list(openpyxl.load_workbook("名單.xlsx").active.iter_rows(values_only=True))
head, body = rows[0], rows[1:]

for row in body:
    data = dict(zip(head, row))
    doc = Document("範本.docx")
    for p in doc.paragraphs:                      # 表格內文字要另外走 doc.tables
        text = p.text
        for k, v in data.items():
            text = text.replace("{{%s}}" % k, str(v))
        if text != p.text and p.runs:
            for r in p.runs[1:]:
                r._element.getparent().remove(r._element)
            p.runs[0].text = text
    doc.save("輸出/%s.docx" % data[head[0]])
```

要轉 PDF：Windows 有 Word 時可用 `docx2pdf`（臨時安裝）；沒有 Word 時用 LibreOffice `soffice --headless --convert-to pdf`。

### 2. Excel：計算、標色、依欄位拆成多個檔案

```python
import openpyxl
from openpyxl.styles import PatternFill

wb = openpyxl.load_workbook("總表.xlsx")
ws = wb.active
red = PatternFill("solid", fgColor="FFC7CE")
for row in ws.iter_rows(min_row=2):
    if row[2].value is not None and row[2].value < 60:   # 第 3 欄 < 60 標紅
        row[2].fill = red
wb.save("總表_標色.xlsx")

# 依「第 1 欄」拆檔
rows = list(ws.iter_rows(values_only=True))
head, body = rows[0], rows[1:]
for key in sorted({r[0] for r in body}):
    nb = openpyxl.Workbook(); nw = nb.active; nw.append(head)
    for r in body:
        if r[0] == key: nw.append(r)
    nb.save("拆檔/%s.xlsx" % key)
```

公式欄位（`=SUM(...)`）用 `openpyxl` 寫入後，要在 Excel 開啟才會算出值；需要立刻讀到數值時，在 Python 裡先自己算好再寫入。

### 3. PDF：合併、抽頁

```python
from pypdf import PdfReader, PdfWriter

w = PdfWriter()
for f in ["a.pdf", "b.pdf", "c.pdf"]:        # 依順序合併
    w.append(f)
w.write("合併.pdf")

r = PdfReader("合併.pdf")
out = PdfWriter()
for i in range(4, 8):                        # 第 5～8 頁（索引從 0 開始）
    out.add_page(r.pages[i])
out.write("抽頁.pdf")
```

### 4. PDF 加浮水印

```python
import io
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

buf = io.BytesIO()
c = canvas.Canvas(buf, pagesize=A4)
c.setFillAlpha(0.15); c.setFont("Helvetica", 48)
c.translate(300, 400); c.rotate(45); c.drawCentredString(0, 0, "INTERNAL USE ONLY")
c.save(); buf.seek(0)
mark = PdfReader(buf).pages[0]

out = PdfWriter()
for page in PdfReader("原檔.pdf").pages:
    page.merge_page(mark); out.add_page(page)
out.write("原檔_浮水印.pdf")
```

**浮水印要用中文字時**，reportlab 預設字型沒有中文，要先 `pdfmetrics.registerFont(TTFont("msjh", r"C:\Windows\Fonts\msjh.ttc", subfontIndex=0))`（Windows 微軟正黑體）再 `setFont("msjh", 36)`；沒註冊會變成方框。

### 5. PDF 轉圖片、取出文字

```python
import pymupdf

doc = pymupdf.open("講義.pdf")
doc[11].get_pixmap(dpi=150).save("第12頁.png")      # 第 12 頁轉圖，可貼進簡報
text = "\n".join(page.get_text() for page in doc)  # 純文字；掃描檔沒有文字層，需要 OCR
```

掃描檔（圖片式 PDF）取不到文字：要 OCR 時臨時裝 `ocrmypdf` 並另外安裝 Tesseract，這需要系統層級安裝，先問使用者同意。

### 6. 文件轉 Markdown（先整理再給 AI 讀）

```python
from markitdown import MarkItDown
print(MarkItDown().convert("報告.pdf").text_content)   # 也支援 .docx .pptx .xlsx
```

### 7. 連結轉 QR Code

```python
import qrcode
qrcode.make("https://example.com/slides").save("qr.png")
```

做成簡報結尾頁的「掃我開啟」時，圖放在 `images/` 內，大小至少 240px，白底黑碼不要加特效，投影時才掃得到。

## 注意事項

- 輸出一律存新檔，並說明存在哪；批次前先用 1～2 筆試跑、請使用者看過結果再全量執行。
- 大量檔案時先印出「將處理 N 個檔案」並確認，再開始寫入。
- 不要把客戶資料、個資或公司內部文件送到外部線上轉檔網站。
- 套件版本差異造成錯誤時，先看錯誤訊息再查該套件文件，不要憑記憶改寫 API。
