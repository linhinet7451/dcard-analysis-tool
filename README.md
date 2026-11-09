# Dcard 文章與留言蒐集器（含 Gemini 趨勢分析）

一個單檔 Tkinter 桌面工具：選擇 Dcard 看板、蒐集文章與留言、預覽、匯出 Excel，
還能把蒐集到的文章送到 Google Gemini 產生逐篇摘要、分類與趨勢報告。

本專案是 PTT 蒐集工具的 Dcard 版本：介面、操作流程、Gemini 分析與 Excel 匯出
結構相同，資料來源改為 Dcard 內部 `_api`（JSON），不再解析 HTML。

## 功能

- **看板複選**：左側可搜尋、篩選、全選／清空、加入自訂看板代碼；內建 16 個看板
  （感情、心情、有趣、女孩、男孩、美妝、穿搭、美食、旅遊、寵物、影劇、
  運動、遊戲、科技、工作、閒聊），也可按「更新看板清單（線上）」從
  `/_api/forums` 取得最新清單，失敗時自動沿用內建清單。
- **蒐集設定**：每個看板 1～200 篇（預設 10 篇）、最新／熱門篩選、
  是否同步抓留言（每篇最多 500 則）、請求間隔 0.8 秒、每看板最多翻 40 頁，
  可隨時停止，進度條與執行紀錄即時顯示。
- **Dcard API 客戶端**：`requests.Session`，UA 為 Chrome、Accept 為
  `application/json`、Referer 為 `https://www.dcard.tw/`，逾時 20 秒，
  5xx／網路錯誤重試 3 次；403／429 會明確提示並停止。
- **瀏覽器 Cookie 模式（選填）**：若直接連線被拒，可在瀏覽器登入 Dcard 後
  把 Cookie 貼到設定區再試。Cookie 只存在記憶體，絕不寫入檔案或 Excel。
- **HAR 匯入（主要備援）**：按「匯入 HAR 檔案…」，在自己的 Chrome 逛 Dcard 後
  匯出的 HAR 檔可直接匯入，離線可用，不需繞過任何防護；詳見「確切可用的方法」。
- **curl_cffi 重試（選用）**：403 時自動試一次 chrome131 TLS 偽裝；已實測在被擋
  的 IP 上仍是 403，主因是 IP，未安裝時主程式仍可正常使用。
- **Playwright 擷取**：`capture_dcard_playwright.py` 在自己電腦上用持續性
  Chrome profile 擷取 `_api` 回應存成 JSON。
- **展示範例資料**：按「載入展示範例資料」可離線載入 6 篇文章、12 則留言
  （感情／美食／旅遊，內容明確標示為展示），用來測試匯出 Excel 與 Gemini 流程，
  不需網路。
- **Gemini AI 趨勢分析**：逐篇摘要、分類（財經投資、科技、感情人際、生活消費、
  娛樂影視、運動、時事政治、社會事件、美食旅遊、工作職場、其他）、情緒
  （正面／負面／中立／爭議）、關鍵字，再產生 Markdown 趨勢報告與程式計算附錄。
- **Excel 匯出**：工作表「文章」「留言」，若有 AI 結果再加「AI趨勢報告」；
  標題列深藍底白字、凍結首列、自動篩選、欄寬已調整。

## 與 PTT 工具的差異

| 項目 | PTT 工具 | Dcard 工具 |
| --- | --- | --- |
| 資料來源 | ptt.cc 公開網頁版 HTML | Dcard `_api` JSON |
| 看板 | Gossiping、Stock 等 | relationship、food、travel 等 alias |
| 文章指標 | 推／噓／箭頭數 | 愛心數、留言數 |
| 留言 | 文章頁內推文 | `/posts/{id}/comments` 分頁（after=樓層） |
| 年齡確認 | over18 Cookie | 無；改為選填瀏覽器 Cookie |
| 額外依賴 | beautifulsoup4 | 無（不解析 HTML） |

## 安裝

需求：Python 3.11 以上。

```bash
pip install -r requirements_dcard.txt
python dcard_desktop.py
```

`tkinter` 為 Python 標準函式庫，不需另外安裝；Linux 若缺 Tk，請用系統套件
安裝（例如 Ubuntu：`sudo apt install python3-tk`）。

## 使用步驟

1. 執行 `python dcard_desktop.py`。
2. 左側勾選要蒐集的看板（可在搜尋框輸入名稱或 alias 篩選；輸入英文代碼後
   按「加入自訂」可加入內建清單沒有的看板）。
3. 設定每個看板文章數、最新／熱門、是否抓留言。
4. （選用）先按「測試 Dcard 連線」確認 `_api` 是否可達。
5. 按「▶ 開始蒐集」，可在「文章預覽」「留言預覽」查看結果，隨時按「■ 停止」。
6. 按「匯出 Excel…」選擇儲存位置，預設檔名 `Dcard_YYYYMMDD_HHMMSS.xlsx`。
7. （選用）在「AI 趨勢分析」分頁輸入 Gemini API Key，按「測試連線／列出模型」，
   再按「✨ 產生趨勢報告」，完成後可按「儲存報告…」（預設
   `Dcard趨勢報告_YYYYMMDD_HHMMSS.md`），報告也會隨 Excel 一起匯出。

## 使用的 Dcard API

基底網址：`https://www.dcard.tw/_api`

| 用途 | 端點 |
| --- | --- |
| 看板清單 | `GET /forums` |
| 看板文章列表 | `GET /forums/{alias}/posts?popular=false&limit=30&before={postId}` |
| 單篇文章全文 | `GET /posts/{id}` |
| 文章留言 | `GET /posts/{id}/comments?after={floor}&limit=30` |

文章欄位：id、title、excerpt／content、createdAt、commentCount、likeCount、
forumAlias／forumName、anonymous、gender、school。
留言欄位：id、floor、content、createdAt、likeCount、anonymous、gender、school。

以上端點為 Dcard 網頁版自身使用的內部 API，並非官方公開文件；Dcard 改版時
可能變動，屆時需更新解析邏輯。

## 確切可用的方法（已實測）

直接問「Dcard 到底能不能抓」：可以，但不是每種方法在每台機器上都通。
以下是在這台機器上實際測過的結論，依推薦順序排列：

| 排名 | 方法 | 實測結果 | 成本 |
| --- | --- | --- | --- |
| 1 | **在家用／手機網路直接用 `_api`** | 最簡單免費。在自己的電腦、家用或手機網路上通常可達；雲端／資料中心 IP 會被 Cloudflare 擋 403 | 免費 |
| 2 | **HAR 匯入（本工具「匯入 HAR 檔案…」）** | **保證可用**。在自己的 Chrome 逛 Dcard，F12 → Network → 勾 Preserve log → 右鍵 Save all as HAR with content，再匯入本工具。不需繞過任何防護，因為資料是你自己的瀏覽器正常載入的 | 免費 |
| 3 | **Playwright 用自己的 Chrome profile 擷取 `_api`** | 在自己電腦上最穩的自動化方式。見 `capture_dcard_playwright.py`：用持續性 profile 開 Chromium，監聽 `/_api/` 回應存成 JSON。不要在雲端 VM 上跑，一定會被擋 | 免費 |
| 4 | **curl_cffi（TLS 偽裝）** | 已實測：chrome131 與 chrome120 在被擋的 IP 上**也還是 403**。TLS 指紋不是主因，IP 信譽才是。本工具 403 時會自動試一次，失敗就明確提示，不會無限重試 | 免費（選用） |
| 5 | **付費爬蟲 API** | Oxylabs Web Scraper API 有 Dcard 專用 source；Firecrawl、ScrapingBee、ScraperAPI 等可處理 Cloudflare；另有 omniscrape.dev 的 Dcard 專用 API。適合不想自己處理防護、願意付費的情況 | 付費 |

注意：`/_api` 與 `/service/api/v2` 在同一台主機、同一層 Cloudflare 後面，
換版本不會繞過封鎖。真正決定能不能通的，是「你用哪個 IP、哪個瀏覽器工作階段」，
而不是換哪支端點。

## HAR 匯入步驟

1. 在一般 Chrome（自己的電腦）開啟 Dcard 並登入。
2. 按 F12 開啟開發人員工具，切到 Network 分頁，勾選 Preserve log。
3. 逛要分析的看板、點開幾篇文章、滑動留言。
4. 在 Network 面板按右鍵 → Save all as HAR with content，存成 `.har` 檔。
5. 回到本工具，按「匯入 HAR 檔案…」選擇該檔案，即可預覽、匯出 Excel、做 Gemini 分析。

## Playwright 擷取步驟

```bash
pip install playwright
playwright install chromium
python capture_dcard_playwright.py --forum food --posts 10
# 需要看畫面或要登入時：
python capture_dcard_playwright.py --forum relationship --posts 20 --headed
```

擷取到的 JSON 會存在 `./har_captures/`，檔名為 `<時間戳>_<編號>.json`。
請在自己的電腦上執行，不要在雲端 VM 上跑。

## Cloudflare 403 疑難排解

從某些伺服器／資料中心 IP 直接請求 `_api`，可能被 Cloudflare 回 403。
本工具遇到 403／429 會先試一次 curl_cffi（若有安裝），失敗就直接停止，
不會無限重試。請依序嘗試：

1. **HAR 匯入（最優先）**：照上一節在自己的 Chrome 匯出 HAR 再匯入本工具。
2. **Playwright 擷取**：在自己電腦上跑 `capture_dcard_playwright.py`。
3. **更換網路**：改用家用／手機網路等非資料中心 IP，再按「測試 Dcard 連線」。
4. **Cookie 模式**：在瀏覽器登入 Dcard，複製請求的 Cookie 貼到設定區的
   「瀏覽器 Cookie」欄位，再按測試連線。Cookie 只存在記憶體，不會存檔。
5. **稍後再試**：等一段時間後再按「測試 Dcard 連線」。
6. **展示模式**：先按「載入展示範例資料」，確認匯出 Excel 與 Gemini 流程正常，
   等網路條件允許再做真實蒐集。

請勿頻繁重試或縮短請求間隔，以免被進一步限流。

## Gemini 設定

- 到 Google AI Studio 取得 Gemini API Key，貼到「AI 趨勢分析」分頁。
- API Key 只存在記憶體（或 `GEMINI_API_KEY` 環境變數），不會寫入檔案或 Excel。
- 按下分析前會先跳出確認視窗：文章內文與留言節錄會傳送到 Google Gemini API。
- 預設模型 `gemini-flash-latest`，可在分頁下拉選擇，或按「測試連線／列出模型」
  載入帳號可用模型。
- 批次大小為每批 8 篇，失敗的批次會略過並記錄，不會中斷整體流程；
  API Key 無效、模型不存在等致命錯誤會直接中止並提示。

## Excel 輸出

- **文章**：看板、看板代碼、文章ID、標題、作者、學校、內文、發布時間、
  愛心數、留言數、文章網址；有 AI 結果時追加 AI摘要、AI分類、AI情緒、AI關鍵字。
- **留言**：看板、文章ID、文章標題、樓層、使用者、學校、留言內容、愛心數、
  時間、文章網址。
- **AI趨勢報告**（有報告時）：整份 Markdown 報告逐行寫入，附程式計算的統計附錄。

## 隱私與聲明

- Gemini API Key 與瀏覽器 Cookie 只存在記憶體，關閉程式即消失，不會寫入
  檔案、Excel 或任何設定檔。
- 本工具僅供個人研究與分析使用。請遵守 Dcard 使用條款，合理使用、控制頻率，
  不要大量抓取、不要用於騷擾、不要轉售或公開他人個資。
- 展示範例資料為內建假資料，僅供測試流程使用，不是真實 Dcard 貼文。

## 專案結構

```
dcard-analysis-tool/
├── dcard_desktop.py            # 主程式（Tkinter 桌面工具，單檔）
├── capture_dcard_playwright.py # Playwright 擷取 _api 回應（自己電腦上執行）
├── requirements_dcard.txt      # requests、openpyxl（curl-cffi、playwright 為選用）
├── test_dcard.py               # 純函式測試：python3 test_dcard.py
├── README.md
└── .gitignore
```

## 測試

```bash
python3 -m py_compile dcard_desktop.py test_dcard.py
python3 test_dcard.py
python3 -c "import dcard_desktop; print('import ok')"
```
