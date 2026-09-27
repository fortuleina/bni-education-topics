const express = require("express");
const path = require("path");
const ExcelJS = require("exceljs");
const archiver = require("archiver");

const app = express();
const PORT = process.env.PORT || 3000;

// Google 試算表用 Apps Script 部署出來的網頁應用程式網址，負責實際讀寫試算表。
// 部署步驟見 README.md。
const SHEETS_WEBAPP_URL = process.env.SHEETS_WEBAPP_URL;
// 選填：跟 Apps Script 裡 Script Properties 設的 WRITE_SECRET 一致，防止網址外流被亂寫入。
const SHEETS_WEBAPP_SECRET = process.env.SHEETS_WEBAPP_SECRET || "";

if (!SHEETS_WEBAPP_URL) {
  console.warn(
    "⚠️  尚未設定 SHEETS_WEBAPP_URL 環境變數，案例素材庫將無法讀寫 Google 試算表。"
  );
}

// 試算表欄位標題（第一列）跟這裡的中文字要完全一致，讀寫都靠這幾個 key 對應。
const COL = {
  referrer: "誰給的引薦",
  recipient: "引薦給誰",
  amount: "金額",
  referralType: "引薦類型",
  source: "內部／外部",
  note: "案例簡述",
  submitter: "熱心填寫者",
  createdAt: "填寫時間",
};

const TYPE_CODE_TO_LABEL = { recurring: "週期性固定引薦", oneoff: "單次引薦" };
const TYPE_LABEL_TO_CODE = { "週期性固定引薦": "recurring", "單次引薦": "oneoff" };
const SOURCE_CODE_TO_LABEL = { internal: "內部引薦", external: "外部引薦" };
const SOURCE_LABEL_TO_CODE = { "內部引薦": "internal", "外部引薦": "external" };

const VALID_TYPES = Object.keys(TYPE_CODE_TO_LABEL);
const VALID_SOURCES = Object.keys(SOURCE_CODE_TO_LABEL);

// 分會會員名單：「誰給的引薦」「熱心填寫者」只能從這份名單選，避免同一個人被打成不同寫法，
// 資料清洗時比對困難。「引薦給誰」允許名單外的名字（外部引薦對象常常不是會員）。
// 依姓氏筆畫排序（跟前端 views/referral.html 的 MEMBERS 順序一致）
// 會員照片下載頁籤：串接朋友做的「BNI 簡報整合系統」API，取得會員大頭照打包成 zip。
// 這支 API 是單一朋友維護、單執行緒、容易冷啟動，所以：
// 1. timeout 抓寬一點（90 秒）
// 2. 批次下載時一定要「逐一」照順序打，絕對不要同時平行打好幾支請求，避免把對方服務打掛
const MEMBER_API_BASE = (process.env.MEMBER_API_BASE || "https://bni-ppt-combine.onrender.com/api/v1").replace(/\/+$/, "");
const MEMBER_API_KEY = process.env.MEMBER_API_KEY || "";

if (!MEMBER_API_KEY) {
  console.warn(
    "⚠️  尚未設定 MEMBER_API_KEY 環境變數，「會員照片下載」頁籤將無法使用。"
  );
}

async function memberApiFetch(pathAndQuery, options = {}) {
  if (!MEMBER_API_KEY) {
    const err = new Error("尚未設定 MEMBER_API_KEY，無法連線到會員照片服務");
    err.code = "NO_API_KEY";
    throw err;
  }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 90000);
  try {
    return await fetch(`${MEMBER_API_BASE}${pathAndQuery}`, {
      ...options,
      headers: { "X-API-Key": MEMBER_API_KEY, ...(options.headers || {}) },
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timer);
  }
}

const MEMBERS = ["丁慶儒","王世品","王孟哲","王懿德","田慶堂","江柏陞","余博軒","吳書林","吳旌緯","呂孟翰","呂炘豪","李坊祥","李恩誠","李曉玫","周子嵐","周順智","周榆璇","周曜贊","林文裕","林宏一","林佩佩","林宗翰","林定豪","林品彤","林思裴","林昱璋","林家蔚","林詠淇","林駿維","侯冠瑋","洪敏智","洪翊嘉","翁煜宸","康清智","張婉靖","張博凱","張綺耘","曹原彰","梁虔駖","莊子樂","許國祥","連育正","郭玉琴","陳沛慈","陳依凡","陳宜信","陳芷綺","陳姵文","陳建安","陳春長","陳朝銘","陳睿民","陳薏如","彭澤偉","曾紹恩","曾惠君","游姿菱 Yu Sleeping Beauty","馮威憲","黃同慶","黃勇嘉","黃建華","黃莞芹","黃瑀珍","黃鈺淳","黃麗娟","黃麟翔","楊和宥","楊嘉慧","葉奕廷","葉蕥霈","雷文鳳","廖婕詠","廖愷伶","廖筱蘭","熊庭逸","劉美葳","劉宸緯","劉雍致","潘建勳","潘祥鈞","蔡伊鎔","蔡昕宏","蔡瑋倫","盧冠臻","賴丰培","戴劭哲","謝宥騏","鍾坤宏","鍾承翰","蘇姮勻"];
const MEMBER_SET = new Set(MEMBERS);

async function readAll() {
  if (!SHEETS_WEBAPP_URL) return [];
  const res = await fetch(SHEETS_WEBAPP_URL, { method: "GET" });
  if (!res.ok) {
    throw new Error(`讀取 Google 試算表失敗（HTTP ${res.status}）`);
  }
  const rows = await res.json();
  if (!Array.isArray(rows)) {
    throw new Error("Google 試算表回傳的格式不正確");
  }

  return rows
    .map((row, i) => {
      const amountRaw = row[COL.amount];
      const amountNum = Number(String(amountRaw).replace(/[^0-9.-]/g, ""));
      const createdAtRaw = row[COL.createdAt];
      let createdAt = null;
      if (createdAtRaw) {
        const d = new Date(createdAtRaw);
        if (!isNaN(d.getTime())) createdAt = d.toISOString();
      }
      return {
        id: "row-" + i,
        referrer: String(row[COL.referrer] || "").trim(),
        recipient: String(row[COL.recipient] || "").trim(),
        amount: Number.isFinite(amountNum) ? amountNum : 0,
        referralType: TYPE_LABEL_TO_CODE[String(row[COL.referralType] || "").trim()] || null,
        source: SOURCE_LABEL_TO_CODE[String(row[COL.source] || "").trim()] || null,
        note: String(row[COL.note] || "").trim(),
        submitter: String(row[COL.submitter] || "").trim(),
        createdAt,
      };
    })
    .filter((r) => r.referrer || r.recipient); // 跳過空白列
}

async function appendRecord(record) {
  if (!SHEETS_WEBAPP_URL) {
    throw new Error("尚未設定 SHEETS_WEBAPP_URL，無法寫入");
  }
  const payload = {
    referrer: record.referrer,
    recipient: record.recipient,
    amount: record.amount,
    referralTypeLabel: TYPE_CODE_TO_LABEL[record.referralType],
    sourceLabel: SOURCE_CODE_TO_LABEL[record.source],
    note: record.note || "",
    submitter: record.submitter,
    secret: SHEETS_WEBAPP_SECRET,
  };

  // Apps Script 的 /exec 網址在執行完 doPost 後，通常會回一個 302，
  // 轉址到 script.googleusercontent.com 底下一組一次性網址去取得執行結果。
  // Node 內建 fetch 對「非 GET 請求收到 302」的自動轉址（會把方法改成 GET
  // 並丟掉 body）在某些情況下會導致轉址失敗，這裡改成自己接手轉址，
  // 行為更穩定、也比較好排查問題。
  let res = await fetch(SHEETS_WEBAPP_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    redirect: "manual",
  });

  if (res.status >= 300 && res.status < 400) {
    const location = res.headers.get("location");
    if (!location) {
      throw new Error(`寫入 Google 試算表失敗（收到轉址回應但沒有 Location，HTTP ${res.status}）`);
    }
    res = await fetch(location, { method: "GET" });
  }

  const rawText = await res.text();
  let data = {};
  try {
    data = JSON.parse(rawText);
  } catch (_) {
    // 回應不是 JSON，通常代表 Google 那邊回了一個錯誤頁面（例如網址錯誤、
    // 部署權限設定不對），下面直接把原始內容印進 log 方便排查。
  }

  if (!res.ok || data.error) {
    console.error(
      "寫入 Google 試算表失敗，原始回應狀態：", res.status,
      "，前 500 字內容：", rawText.slice(0, 500)
    );
    throw new Error(data.error || `寫入 Google 試算表失敗（HTTP ${res.status}）`);
  }
  return data;
}

app.use(express.json());

// 素材下載檔案（目前是三分鐘培訓簡報母片），實際檔案放在 public/downloads/ 底下。
// 存檔案時瀏覽器看到的檔名這裡刻意用英文：測試發現部分瀏覽器（尤其舊版 Chrome／
// 內嵌瀏覽器）在 Content-Disposition 標頭或 HTML download 屬性帶中文檔名時會出錯，
// 直接把檔案存成一個沒有副檔名的「download」檔案，使用者常常打不開、搞不清楚是什麼檔案。
// 改用純英文檔名可以確保各種瀏覽器都能穩定存成正確的 .pptx 檔；網頁上的按鈕文字跟卡片
// 說明還是用中文，使用者看畫面就知道這是什麼檔案，不影響辨識。
// 之後要換新的母片檔案，把新檔案放進 public/downloads/、改這裡的設定即可。
const RESOURCE_DOWNLOADS = {
  "/downloads/three-minute-training-master.pptx": {
    file: "three-minute-training-master.pptx",
    filename: "BNI-three-minute-training-master.pptx",
  },
};
Object.entries(RESOURCE_DOWNLOADS).forEach(([route, { file, filename }]) => {
  app.get(route, (req, res) => {
    res.download(path.join(__dirname, "public", "downloads", file), filename, (err) => {
      if (err && !res.headersSent) {
        console.error(`下載素材檔案失敗（${route}）：`, err);
        res.status(404).send("找不到這個下載檔案");
      }
    });
  });
});

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "views", "index.html"));
});

app.get("/referral", (req, res) => {
  res.sendFile(path.join(__dirname, "views", "referral.html"));
});

app.get("/api/referrals", async (req, res) => {
  try {
    const all = await readAll();
    res.json(all);
  } catch (err) {
    console.error("讀取案例失敗：", err);
    res.status(502).json({ error: "讀取 Google 試算表失敗，請稍後再試" });
  }
});

app.post("/api/referrals", async (req, res) => {
  const { referrer, recipient, amount, referralType, source, note, submitter } = req.body || {};

  // 誰給的引薦／熱心填寫者：可以是名單裡的會員，也可以是網頁上選「其他」自行輸入的名字
  // （例如來賓、非會員協助填寫），所以這裡只檢查有沒有填，不檢查是否在會員名單裡。
  if (!referrer || !String(referrer).trim()) {
    return res.status(400).json({ error: "請填寫「誰給的引薦」" });
  }
  // 引薦給誰：一定要是會員名單裡的人。
  if (!recipient || !MEMBER_SET.has(String(recipient).trim())) {
    return res.status(400).json({ error: "請從名單選擇「引薦給誰」" });
  }
  if (!submitter || !String(submitter).trim()) {
    return res.status(400).json({ error: "請填寫「熱心填寫者」" });
  }
  const amountNum = Number(amount);
  if (!Number.isFinite(amountNum) || amountNum < 0) {
    return res.status(400).json({ error: "金額格式不正確" });
  }
  if (!VALID_TYPES.includes(referralType)) {
    return res.status(400).json({ error: "請選擇引薦類型" });
  }
  if (!VALID_SOURCES.includes(source)) {
    return res.status(400).json({ error: "請選擇內部／外部引薦" });
  }

  const record = {
    referrer: String(referrer).trim(),
    recipient: String(recipient).trim(),
    amount: amountNum,
    referralType,
    source,
    note: note ? String(note).trim() : "",
    submitter: String(submitter).trim(),
  };

  try {
    await appendRecord(record);
    res.status(201).json({ ...record, createdAt: new Date().toISOString() });
  } catch (err) {
    console.error("寫入資料失敗：", err);
    res.status(502).json({ error: err.message || "伺服器寫入資料時發生錯誤，請稍後再試" });
  }
});

app.get("/api/referrals/export.xlsx", async (req, res) => {
  let all;
  try {
    all = await readAll();
  } catch (err) {
    return res.status(502).send("讀取 Google 試算表失敗，請稍後再試");
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("引薦紀錄");

  ws.columns = [
    { header: "誰給的引薦", key: "referrer", width: 16 },
    { header: "引薦給誰", key: "recipient", width: 16 },
    { header: "金額", key: "amount", width: 12 },
    { header: "引薦類型", key: "referralTypeLabel", width: 16 },
    { header: "內部／外部", key: "sourceLabel", width: 12 },
    { header: "案例簡述／分享重點", key: "note", width: 50 },
    { header: "熱心填寫者", key: "submitter", width: 14 },
    { header: "填寫時間", key: "createdAt", width: 22 },
  ];
  ws.getRow(1).font = { bold: true };

  all.forEach((r) => {
    ws.addRow({
      referrer: r.referrer,
      recipient: r.recipient,
      amount: r.amount,
      referralTypeLabel: TYPE_CODE_TO_LABEL[r.referralType] || "",
      sourceLabel: SOURCE_CODE_TO_LABEL[r.source] || "",
      note: r.note || "",
      submitter: r.submitter || "",
      createdAt: r.createdAt || "",
    });
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", 'attachment; filename="referrals.xlsx"');
  await wb.xlsx.write(res);
  res.end();
});

/* ============ 會員照片下載 ============ */

app.get("/api/members", async (req, res) => {
  try {
    const apiRes = await memberApiFetch("/members");
    if (!apiRes.ok) {
      throw new Error(`讀取會員名單失敗（HTTP ${apiRes.status}）`);
    }
    const data = await apiRes.json();
    const members = Array.isArray(data.members) ? data.members : [];
    // 只把畫面需要的欄位傳給前端，API金鑰留在伺服器這邊，不會流出去。
    res.json(
      members
        .filter((m) => m && m.id)
        .map((m) => ({
          id: m.id,
          name: m.name || m.id,
          industry: m.industry || "",
        }))
    );
  } catch (err) {
    console.error("讀取會員照片名單失敗：", err);
    const msg =
      err.code === "NO_API_KEY"
        ? "尚未設定會員照片服務的 API 金鑰，請聯絡管理員設定 MEMBER_API_KEY 環境變數。"
        : "讀取會員名單失敗，請稍後再試（對方服務可能正在冷啟動，可以再試一次）。";
    res.status(502).json({ error: msg });
  }
});

app.post("/api/members/photos.zip", async (req, res) => {
  const idsRaw = Array.isArray(req.body?.ids) ? req.body.ids : [];
  const ids = [...new Set(idsRaw.filter((id) => typeof id === "string" && id.trim()))];

  if (!ids.length) {
    return res.status(400).json({ error: "請至少選擇一位會員" });
  }
  if (!MEMBER_API_KEY) {
    return res.status(503).json({
      error: "尚未設定會員照片服務的 API 金鑰，請聯絡管理員設定 MEMBER_API_KEY 環境變數。",
    });
  }

  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", 'attachment; filename="member-photos.zip"');

  const archive = archiver("zip", { zlib: { level: 9 } });
  let archiveFailed = false;
  archive.on("warning", (err) => console.warn("打包照片 zip 警告：", err));
  archive.on("error", (err) => {
    archiveFailed = true;
    console.error("打包照片 zip 時發生錯誤：", err);
    if (!res.headersSent) res.status(500);
    res.end();
  });
  archive.pipe(res);

  const usedNames = new Set();
  const failed = [];

  // 對方 API 單執行緒、容易冷啟動，這裡刻意用 for...of 逐一 await，
  // 不要用 Promise.all 之類的方式同時平行打多支請求。
  for (const id of ids) {
    if (archiveFailed) break;
    try {
      const photoRes = await memberApiFetch(`/members/${encodeURIComponent(id)}/photo`);
      if (!photoRes.ok) {
        failed.push(id);
        continue;
      }
      const buf = Buffer.from(await photoRes.arrayBuffer());

      // 檔名以 API 實際回傳的為準：伺服器端可能會壓縮圖片，副檔名不一定跟原始上傳的一樣。
      let filename = null;
      const disposition = photoRes.headers.get("content-disposition") || "";
      const match = disposition.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
      if (match) {
        try {
          filename = decodeURIComponent(match[1]);
        } catch (_) {
          filename = match[1];
        }
      }
      if (!filename) {
        const contentType = photoRes.headers.get("content-type") || "";
        const ext = contentType.includes("png") ? "png" : contentType.includes("webp") ? "webp" : "jpg";
        filename = `${id}.${ext}`;
      }

      // 避免不同會員的檔名剛好重複而互相覆蓋。
      let finalName = filename;
      let n = 2;
      while (usedNames.has(finalName)) {
        const dot = filename.lastIndexOf(".");
        finalName = dot === -1 ? `${filename}-${n}` : `${filename.slice(0, dot)}-${n}${filename.slice(dot)}`;
        n++;
      }
      usedNames.add(finalName);

      archive.append(buf, { name: finalName });
    } catch (err) {
      console.error(`下載會員「${id}」照片失敗：`, err);
      failed.push(id);
    }
  }

  if (archiveFailed) return;

  if (failed.length) {
    archive.append(
      `以下會員的照片下載失敗，可能是還沒有上傳照片，或對方服務當下沒有回應：\n\n${failed.join("\n")}\n`,
      { name: "下載失敗名單.txt" }
    );
  }

  await archive.finalize();
});

app.listen(PORT, () => {
  console.log(`BNI education tools server running on port ${PORT}`);
});
