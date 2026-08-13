# 補課通知單產生器

## 功能
上傳「簽到表」與「補課名單」兩份 Excel，自動比對每位學員的缺曠日期與對應課程名稱，
一鍵產出每位學員的補課通知單（Word .docx），打包成 ZIP 下載。

## 檔案說明
- `app.py`：Streamlit 網頁介面（上傳檔案、預覽、下載）
- `core.py`：核心邏輯（讀取 Excel、比對資料、產生 Word 文件）
- `requirements.txt`：套件需求清單

## 本機執行方式
```bash
pip install -r requirements.txt
streamlit run app.py
```
瀏覽器會自動開啟 http://localhost:8501

## 資料格式要求

### 簽到表 Excel
- 檔名需包含班級關鍵字：初級 / 中級 / 高級 / 研經（用來對應補課名單的分頁）
- 第2列：姓名、法名、組別、日期欄位標題
- 第3列：實際上課日期（MM/DD）
- 每列一位學員，最後兩欄需為「缺曠堂數」與「缺曠日期清單」（以「、」分隔）

### 補課名單 Excel
- 每個分頁對應一個班級，分頁名稱需為：初級 / 中級 / 高級 / 研經
- 第一欄：缺課日期（日期格式）
- 第三欄：課程名稱

## 部署到雲端（給機構同仁使用）
推薦使用 **Streamlit Community Cloud**（免費）：
1. 把這個資料夾上傳到一個 GitHub repository
2. 到 https://share.streamlit.io 用 GitHub 帳號登入
3. 選擇該 repository，指定主檔案為 `app.py`，點擊 Deploy
4. 幾分鐘後會得到一個公開網址，機構同仁打開瀏覽器就能使用，不需要安裝任何軟體

如果機構有自己的伺服器或 IT 資源，也可以用 Docker 部署，或改用 Render / Railway 等平台，做法類似。

## 待辦（等待正式 Word 範本後調整）
- [ ] 套用機構正式補課通知單範本（標題、Logo、蓋章欄位等）
- [ ] 視需要加入發文字號、承辦人姓名等欄位
