# work-slide-skills

把「工作簡報」變成 AI agent 一句話就能做的技能包：互動 HTML 簡報、可編輯 PPTX、整頁 AI 圖簡報，加上它們共用的生圖工具，以及 Word／Excel／PDF 批次處理手冊。給 Claude Code（或 Codex）使用，一鍵安裝。

## 內含 6 個 skill

| Skill                | 產出                                      | 什麼時候用                                |
| -------------------- | --------------------------------------- | ------------------------------------ |
| `html-slide-builder` | Reveal.js 網頁簡報，可部署到 GitHub Pages 取得分享連結 | 要現場互動（文字雲、投票）、要 AI 底圖與去背圖標、要一個網址分享出去 |
| `soil-html-deck`     | 單一 `.html` 檔簡報                          | 要 Chart.js 圖表、可點擊表格、影片嵌入，不需要伺服器、離線可開 |
| `soil-teaching-deck` | 文字可編輯的 `.pptx`（AI 插圖）                   | 只能給對方 PowerPoint、需要事後改字；內建簡報結構設計流程   |
| `soil-image-deck`    | 每頁都是一張 AI 生成整圖的 `.pptx`                 | 視覺衝擊強、不需要改字（開場、社群、海報風）               |
| `draw`               | 圖片（gpt-image-2）                         | 上面四個 skill 內部會呼叫它，也能單獨說「畫一張…」        |
| `file-toolkit`       | Word／Excel／PDF／PPT 批次處理手冊                | 批次套印 Word 範本、Excel 依欄位拆檔、PDF 合併抽頁加浮水印與轉圖片、舊文件轉 Markdown、產生 QR Code |

> 要做「會動的原理動畫」或「配音解說影片」，請改用 [rd-motion-deck](https://github.com/tedchen0814/rd-motion-deck)。

## 一鍵安裝

最簡單：把下面這句話貼給你的 AI agent。

```text
請幫我安裝這個 skill：https://github.com/tedchen0814/work-slide-skills
```

agent 會讀 [AGENT_INSTALL.md](AGENT_INSTALL.md) 自己完成。也可以手動：

**Windows（PowerShell）**

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/tedchen0814/work-slide-skills/main/install.ps1)))
```

**macOS / Linux**

```bash
curl -fsSL https://raw.githubusercontent.com/tedchen0814/work-slide-skills/main/install.sh | bash
```

常用選項：`-Agent codex`（`--agent codex`）裝到 Codex；`-CheckOnly`（`--check-only`）只檢查不改動；`-InstallPrereqs` 用 winget 補裝 Git、Python；`-FirebaseConfigFile`（`--firebase-config`）直接帶入你的 Firebase 設定（見下）。

安裝完**重新啟動 agent**，新 skill 才會載入。已存在的同名 skill 會被搬到 `~/.claude/skill-backups/`，不會直接刪除。

## 需要什麼

| 項目 | 用途 | 沒有的話 |
|---|---|---|
| Python 3.9+ | 生圖、打包與文件處理腳本；會自動 pip 安裝 openai、Pillow、python-pptx、python-docx、openpyxl、pypdf、reportlab、PyMuPDF、qrcode、markitdown（畫圖表用的 matplotlib 不預裝，需要時 agent 再裝） | 無法安裝 |
| Git | 下載 repo | 無法安裝 |
| Node.js 22+ | `soil-teaching-deck` 產生 `.pptx`（PptxGenJS） | 該 skill 不能出 pptx，其他照常 |
| GitHub CLI（`gh`，需已登入） | `html-slide-builder` 發布到 GitHub Pages | 簡報做得出來，但不能一鍵發布 |
| **你自己的 OpenAI API key** | 所有 AI 生圖 | 沒有圖，其他流程照常 |
| **你自己的 Firebase 專案** | `html-slide-builder` 的現場文字雲／投票 | 簡報不放互動頁，其他照常 |

### OpenAI key（生圖）

這個 repo **不含任何金鑰**。生圖費用算在你自己的帳號上（預設用 `quality low`，一份簡報約 4～10 張）。

1. 到 platform.openai.com 建立 API key，並在 Settings → Organization → General 完成 Individual 驗證（沒驗證會出現 `403 Organization must be verified`）。
2. 在你的家目錄建立 `.openai.env`，只有一行：

```text
OPENAI_API_KEY=sk-你的金鑰
```

不要把這個檔案放進任何 repo。

### Firebase（現場互動，可選）

文字雲與投票的資料會寫進**你自己的** Firebase 專案，所以安裝時不會附帶任何人的專案。

1. Firebase Console 建專案 → Build → Firestore Database → 建立資料庫（先選測試模式）。
2. 專案設定 → 你的應用程式 → 新增 Web 應用程式，複製 `firebaseConfig` 存成文字檔。
3. 安裝時帶入：`-FirebaseConfigFile .\cfg.js`（macOS/Linux：`--firebase-config ./cfg.js`）；或裝完後把 `firebaseConfig` 貼給 agent，請它填進 `html-slide-builder` 的 `references/firebase-config.md`。
4. **30 天內**把 Firestore 規則換成該檔裡的「Firestore 安全規則」範本，否則測試模式到期後互動會全部失效。

沒設定 Firebase 時，`html-slide-builder` 會自動不放互動頁，不會出錯。

## 安裝的套件與用途

安裝腳本會替你裝下面這些套件，全部是你的 agent 在背後呼叫的工具，你不需要自己學，只要用說的描述需求。

| 套件 | 用途（白話） | 常見的一句話需求 |
|---|---|---|
| `openai` | 呼叫 OpenAI 生圖（`draw` 使用） | 「畫一張封面圖」 |
| `Pillow` | 圖片裁切、縮放、去背、格式轉換 | 「把這些圖片統一縮成 800 像素寬」 |
| `python-pptx` | 讀寫 PowerPoint（`.pptx`）：產生投影片、改字型、插圖 | 「把每張圖做成一頁投影片」「全部字型改成微軟正黑體」 |
| `python-docx` | 讀寫 Word（`.docx`）：套印範本、批次產生文件 | 「用這份名單套進通知單範本，每人一份」 |
| `openpyxl` | 讀寫 Excel（`.xlsx`）：計算、標色、依欄位拆成多個檔案 | 「依部門欄位把總表拆成一個部門一個檔」 |
| `pypdf` | PDF 合併、拆分、抽頁、疊浮水印 | 「合併這 3 份 PDF，只留第 5～8 頁」 |
| `reportlab` | 產生 PDF 內容，主要用來畫浮水印、簽名頁 | 「每頁加上『僅供內部使用』浮水印」 |
| `PyMuPDF` | PDF 轉圖片、取出文字與圖片、擷取頁面 | 「把這份 PDF 第 12 頁轉成圖，貼進簡報」 |
| `qrcode[pil]` | 把網址轉成 QR Code 圖片 | 「把簡報連結做成 QR Code 放在結尾頁」 |
| `markitdown` | 把 PDF／Word／PowerPoint／Excel 轉成乾淨的 Markdown，方便餵給 AI 讀 | 「把這份舊簡報轉成文字，我要重新整理」 |
| `pptxgenjs`（npm，需 Node.js） | 另一套產生 `.pptx` 的工具，`soil-teaching-deck` 用它排版 | 「做成可以改字的 PowerPoint」 |

**刻意沒預裝**：`matplotlib`（畫圖表，體積大）、`ocrmypdf`＋Tesseract（掃描檔轉可選取文字，需要系統層級安裝）、`docx2pdf`（Word 轉 PDF，需要本機有 Word）。真的需要時，agent 會先告訴你再臨時安裝。

> `PyMuPDF` 採 AGPL／商業雙授權：每個人在自己電腦上下載使用沒有問題，細節見 [NOTICE.md](NOTICE.md)。

## 怎麼用

直接用說的，agent 會挑對的 skill：

- 「把這份大綱做成互動 HTML 簡報，要有投票」→ `html-slide-builder`
- 「做一份單檔網頁簡報，放這張 CSV 的圖表」→ `soil-html-deck`
- 「做成可以改字的 PowerPoint」→ `soil-teaching-deck`
- 「做一份整頁都是圖的簡報」→ `soil-image-deck`
- 「把這份名單套進 Word 範本，每人一份」「把這幾個 PDF 合併並加浮水印」「這份舊 PDF 轉成文字」→ `file-toolkit`

`html-slide-builder` 會先給你一張大綱表確認，確認後才開始生圖與部署，部署到 GitHub Pages 前也會先問你。

## 已知限制

- `soil-teaching-deck` 的「幾何圖形子流程」需要另外的 `jh-math-geometry` skill（本包未附），只有數學幾何簡報會用到，一般工作簡報不受影響。
- `soil-*` 系列原本的措辭偏教學情境，內容流程仍可用於一般簡報。
- Windows 上若腳本因「未經數位簽署」被擋（常見於雲端硬碟），把 repo 複製到 `C:\` 本機磁碟再執行，或用上面的 `irm` 一行指令。
- `install.sh` 在 Windows 的 Git Bash 找不到 Windows 版 Python 時會失敗，Windows 請用 `install.ps1`。

## 授權與來源

見 [NOTICE.md](NOTICE.md)。
