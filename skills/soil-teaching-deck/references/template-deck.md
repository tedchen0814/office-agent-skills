# 用公司範本做「大量文字」簡報：排版、補圖、目錄跳轉

適用：使用者給了公司 `.pptx` 範本（母片），要把一份很長的大綱（幾十頁以上）做成可編輯的簡報。
腳本：`scripts/template_deck_lib.py`（python-pptx；先 `pip install python-pptx pillow`）。

## 0. 先做
1. 開範本，列出所有版型：`[(i, l.name) for i, l in enumerate(layouts)]`，並用 PowerPoint 匯出圖看實際長相。
2. 改 `template_deck_lib.py` 的 CONFIG：配色、`L_COVER / L_INNER / L_SEP / L_END` 索引、內容區 `X0, W0, Y0, Y1`（避開標題與頁尾）。
3. 範本內的範例頁最後再刪（`prs.part.drop_rel` + 從 `sldIdLst` 移除）。
4. 範本若有機密等級封面，先問使用者要不要放，不要自行省略或自行加入。

## 1. 文字排版：一行寫得完就不要換行
- `fit1(text, size, width)`：整行放不下時，最多縮 `steps` 級；預設 `steps=0`（不縮），表格與提示框可傳 1–2。
- `page(...)` 在 15–22pt 之間逐級試算，**優先選「換行條數最少」的字級，其次取最大**；放不下才降到 13pt 以下。
- 同一頁條列用同一個字級；表格儲存格字級可略小（9–13pt）。
- 寬度是以字數估算（全形 1 em、半形約 0.55 em），**一定要用 PowerPoint 匯出圖檢查**，不要只信估算。
- 補東亞字型：每個 run 都要設 `a:ea`、`a:cs` 的 typeface，否則部分漢字會走備用字型。

## 2. 空白處補圖
- `page()` 會把「內容用完後下方剩餘高度」記到 `FREE`；剩餘 ≥ 1.3 吋就在右下角放 `auto_img(title)` 的圖（沒有對應圖就略過）。
- 圖用 `draw` skill 批次生成，檔名 `a_<slug(title)>_*.png`，放 `images/<範本名>/auto/`：
  ```python
  from template_deck_lib import slug
  # 每頁一張，提示詞 = 該頁主題的具體畫面 + 統一風格 + 「不要任何文字」
  cmd = [python, draw_py, f"{desc}，{STYLE}，{NEG}", "--size", "1024x1024", "--quality", "low",
         "--name", f"a_{slug(title)}", "--outdir", "images/xxx/auto"]
  ```
- 風格固定成一句（例：扁平線稿、純白背景、範本強調色為主）；用 8 個並行即可，約 NT$0.3／張。
- 先跑一次建置得到 `FREE` 清單，再只為「有空白」的頁生圖，最後重跑建置。圖是裝飾，不承載資訊。

## 3. 目錄頁的區段代碼可點擊跳轉
`table()` 會把每個表格的位置記進 `TABLES`（`slide, x, y, widths, row_heights`）。建置完成、**刪除範例頁之前**：
```python
from pptx.oxml.ns import qn
tgt = {}                       # 區段代碼 -> 該區第一頁（分隔頁或首頁）
for sl in prs.slides:
    t = sl.shapes.title.text if sl.shapes.title is not None else ""
    # 依自己的標題規則填 tgt["B"] = sl ...
for sl, x, y, widths, hs in TABLES:
    if sl.shapes.title.text != "目錄頁標題": continue
    for i, code in enumerate(codes):             # 第 0 列是表頭
        yy = y + sum(hs[:i + 1])
        r = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(yy), Inches(widths[0]), Inches(hs[i + 1]))
        r.fill.solid(); r.fill.fore_color.rgb = rgb(ORG)
        clr = r.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
        clr.append(clr.makeelement(qn("a:alpha"), {"val": "0"}))   # 全透明
        r.line.fill.background()
        r.click_action.target_slide = tgt[code]
```
- 放映時直接點；普通檢視要按住 Ctrl 再點。頁面下方加一行提示說明。
- 寫完用 `shape.click_action.target_slide` 讀回，核對「代碼 → 頁碼」。

## 4. 踩坑
- 範本檔若正被 PowerPoint 開著，`save` 會 `PermissionError`；請使用者先關檔或另存新檔名，不要強關使用者的視窗。
- 用 sed 改含 `\n` 的字串會被當成真換行；改用編輯工具。
- 長簡報（60 頁以上）建議用 `Export` 匯出後拼成 2×3 總覽圖逐批檢查，修一輪就重匯出有改的頁。
- 內容不放內網位址、機器名稱、帳號、私有路徑；公開 repo 名稱與網址要照原稿。
