# BNI 教育協調員工具

三分鐘培訓／會後培訓主題庫 ＋ 引薦案例素材庫。原本是純靜態網頁，這次改成有後端的 Web Service，
新增「案例素材庫」頁籤，資料庫是 **Google 試算表**：可以透過表單即時記錄引薦，也可以直接在
試算表裡手動新增／編輯／刪除案例，網頁會讀到最新內容。

## 專案結構

```
server.js                   Express 伺服器，所有路由跟 API 都在這裡
package.json                套件設定
views/index.html            主題庫首頁（三分鐘培訓／會後培訓／案例素材庫）
views/referral.html         引薦紀錄填寫表單頁
google-apps-script/Code.gs  貼到 Google 試算表 Apps Script 的橋接程式碼
```

## 第一步：設定 Google 試算表

1. 新增一個 Google 試算表（或用你原本打算做案例庫的那份）。
2. 第一列填入以下欄位標題，順序不拘，但文字要完全一致：

   `誰給的引薦` `引薦給誰` `金額` `引薦類型` `內部／外部` `案例簡述` `熱心填寫者` `填寫時間`

3. 建議幫下面這幾欄設下拉選單（選取整欄 → 資料 → 資料驗證 → 下拉式選單），手動輸入時才不會打錯字：
   - 「引薦類型」：選項填「週期性固定引薦」「單次引薦」
   - 「內部／外部」：選項填「內部引薦」「外部引薦」
   - 「誰給的引薦」「熱心填寫者」：選項就是分會的 90 位會員名單（跟網頁表單的下拉選單是同一份，
     這樣不管是手動輸入還是網頁表單送出，名字寫法都會一致，方便之後清資料、比對）。
     「引薦給誰」則不建議限制成下拉選單，因為外部引薦的對象常常不是會員。
4. 打開「擴充功能 → Apps Script」，把 `google-apps-script/Code.gs` 的內容整個貼進去。
5.（建議，非必要）左側齒輪圖示「專案設定」→「指令碼屬性」→新增一組
   鍵：`WRITE_SECRET`　值：自己取一串不容易猜到的字串。
6. 右上角「部署」→「新增部署作業」→ 類型選「網頁應用程式」，執行身分選「我」，
   誰可以存取選「任何人」，部署後複製拿到的網址（結尾是 `/exec`）。
7. 之後如果修改 `Code.gs` 的內容，要用「部署」→「管理部署作業」→ 編輯 → 版本選「新版本」
   → 部署，網址才會套用最新程式碼，單純存檔不會生效。

### 關於「熱心填寫者」欄位

網頁表單新增了「熱心填寫者」下拉選單（跟「誰給的引薦」一樣，只能選 90 位會員之一），
用來記錄這筆資料是誰幫忙填寫的，方便之後追查亂填或重複的資料。這個欄位**不會顯示在
公開的案例素材庫卡片上**，但要注意：整個網站目前是完全公開、沒有登入限制的，`/api/referrals`
這支 API 本身還是讀得到這個欄位（只是網頁畫面沒有顯示出來）。如果之後想要更嚴格地保護
這個欄位，需要另外加登入或密碼保護機制，目前這版本還沒做這一層。

## 第二步：本機測試

```bash
npm install
SHEETS_WEBAPP_URL="你的 Apps Script 網址" SHEETS_WEBAPP_SECRET="你設的 WRITE_SECRET" npm start
```

啟動後打開 http://localhost:3000 即可，案例素材庫會直接讀寫你的 Google 試算表。

## 第三步：部署到 Render（Web Service）

目前的 `bni-education-topics.onrender.com` 是舊的 Static Site，沒辦法直接升級成有後端的服務，
需要另外新增一個 **Web Service**。步驟如下：

1. Render 後台 → **New** → **Web Service**，選這個 GitHub repo（`fortuleina/bni-education-topics`）。
2. **Build Command**：`npm install`
3. **Start Command**：`npm start`（或 `node server.js`）
4. **Environment Variables** 加兩個：
   - `SHEETS_WEBAPP_URL` = 上面 Apps Script 部署拿到的網址
   - `SHEETS_WEBAPP_SECRET` = 跟 Apps Script 指令碼屬性裡 `WRITE_SECRET` 一樣的值
5. 資料存在 Google 試算表，**不需要**額外掛 persistent disk。
6. 先用 Render 自動配的暫時網址（例如 `bni-education-topics-xxxx.onrender.com`）完整測試過一輪：
   - 首頁三個頁籤都能切換
   - 案例素材庫可以正常讀取、篩選、匯出 Excel
   - `/referral` 表單送出後，案例素材庫馬上看得到新資料
   - 直接在 Google 試算表手動加一列，重新整理網頁也看得到
7. 確認都沒問題後，才進行「換回原本網址」的最後一步（把舊 Static Site 刪掉、新 Web Service 改名成
   `bni-education-topics`，拿回原本的 `bni-education-topics.onrender.com` 網址）。這一步請小心操作，
   確認新服務一切正常後再刪舊的，避免空窗期太久。

## 資料備份

案例素材庫頁籤右上角有「匯出 Excel」按鈕，隨時可以下載目前所有引薦紀錄的 `.xlsx` 檔案。
另外資料本身就存在 Google 試算表裡，本來就可以隨時用 Google 試算表自己的「檔案 → 下載 → Microsoft Excel」
另外存一份備份。
