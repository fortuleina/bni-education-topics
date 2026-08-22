const express = require("express");
const path = require("path");
const ExcelJS = require("exceljs");

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

  if (!referrer || !MEMBER_SET.has(String(referrer).trim())) {
    return res.status(400).json({ error: "請從名單選擇「誰給的引薦」" });
  }
  if (!recipient || !String(recipient).trim()) {
    return res.status(400).json({ error: "請填寫「引薦給誰」" });
  }
  if (!submitter || !MEMBER_SET.has(String(submitter).trim())) {
    return res.status(400).json({ error: "請從名單選擇「熱心填寫者」" });
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

app.listen(PORT, () => {
  console.log(`BNI education tools server running on port ${PORT}`);
});
