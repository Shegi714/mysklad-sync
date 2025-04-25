// index.js с буферизацией, автоочисткой и автосозданием листов

import { google } from "googleapis";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });
const spreadsheetId = process.env.SPREADSHEET_ID;

const buffers = {}; // 💾 Хранилище буферов строк по листам

async function getCabinets() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "основа!A2:C",
  });
  return res.data.values || [];
}

async function ensureSheetExists(sheetName) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const exists = meta.data.sheets.some((s) => s.properties.title === sheetName);
  if (!exists) {
    console.log(`📄 Создаю лист: ${sheetName}`);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ addSheet: { properties: { title: sheetName } } }],
      },
    });
  }
}

async function resetSheet(sheetName, headers) {
  await ensureSheetExists(sheetName);
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: `${sheetName}!A1:Z10000`,
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `${sheetName}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: [headers] },
  });
  buffers[sheetName] = [];
}

function bufferRow(sheetName, row) {
  if (!buffers[sheetName]) buffers[sheetName] = [];
  buffers[sheetName].push(row);
}

async function flushBuffers() {
  for (const [sheet, rows] of Object.entries(buffers)) {
    if (rows.length === 0) continue;
    await ensureSheetExists(sheet);
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheet}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });
    buffers[sheet] = []; // очищаем после отправки
  }
}

function buildHeaders(login, password) {
  const auth = Buffer.from(`${login}:${password}`).toString("base64");
  return {
    Authorization: `Basic ${auth}`,
    "Content-Type": "application/json",
    "User-Agent": "mysklad-sync-bot",
  };
}

async function getProduct(href, headers, cache) {
  if (cache[href]) return cache[href];
  const res = await fetch(href, { headers });
  const json = await res.json();
  const result = {
    name: json.name || "—",
    article: String(json.article || "—"),
    code: String(json.code || "—"),
  };
  cache[href] = result;
  return result;
}

async function getStock(login, password, cabinet) {
  const headers = buildHeaders(login, password);
  const res = await fetch("https://api.moysklad.ru/api/remap/1.2/report/stock/all?limit=1000", { headers });
  const json = await res.json();
  for (const row of json.rows || []) {
    const name = row.name || "—";
    const article = String(row.article || "—");
    const code = String(row.code || "—");
    const stock = row.stock || 0;
    bufferRow(`Остатки ${cabinet}`, [name, article, code, stock]);
    bufferRow("Остатки общее", [cabinet, name, article, code, stock]);
  }
}

async function getPurchaseOrders(login, password, cabinet) {
  const headers = buildHeaders(login, password);
  const res = await fetch("https://api.moysklad.ru/api/remap/1.2/entity/purchaseorder?expand=positions,agent,state&limit=100", { headers });
  const json = await res.json();
  const cache = {};
  for (const order of json.rows || []) {
    const state = order.state?.name?.toLowerCase() || "";
    if (state.includes("доставлено")) continue;
    const date = order.moment?.split("T")[0] || "—";
    const agent = order.agent?.name || "—";
    const status = order.state?.name || "—";
    for (const pos of order.positions?.rows || []) {
      const { name, article, code } = await getProduct(pos.assortment?.meta?.href, headers, cache);
      const qty = pos.quantity || 0;
      bufferRow(`ПозицииЗаказов ${cabinet}`, [date, agent, status, name, article, code, qty]);
      bufferRow("ПозицииЗаказов общее", [cabinet, date, agent, status, name, article, code, qty]);
    }
  }
}

async function getShipments(login, password, cabinet) {
  const headers = buildHeaders(login, password);
  const res = await fetch("https://api.moysklad.ru/api/remap/1.2/entity/demand?expand=positions,agent,state&limit=100", { headers });
  const json = await res.json();
  const cache = {};
  for (const ship of json.rows || []) {
    const state = ship.state?.name?.toLowerCase() || "";
    if (state.includes("поступило в продажу")) continue;
    const date = ship.moment?.split("T")[0] || "—";
    const agent = ship.agent?.name || "—";
    const status = ship.state?.name || "—";
    for (const pos of ship.positions?.rows || []) {
      const { name, article, code } = await getProduct(pos.assortment?.meta?.href, headers, cache);
      const qty = pos.quantity || 0;
      bufferRow(`Отгрузки ${cabinet}`, [date, agent, status, name, article, code, qty]);
      bufferRow("Отгрузки общее", [cabinet, date, agent, status, name, article, code, qty]);
    }
  }
}

(async () => {
  const cabinets = await getCabinets();

  await resetSheet("Остатки общее", ["Кабинет", "Наименование", "Артикул", "Код", "Остаток"]);
  await resetSheet("ПозицииЗаказов общее", ["Кабинет", "Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);
  await resetSheet("Отгрузки общее", ["Кабинет", "Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);

  for (const [cabinet, login, password] of cabinets) {
    await resetSheet(`Остатки ${cabinet}`, ["Наименование", "Артикул", "Код", "Остаток"]);
    await resetSheet(`ПозицииЗаказов ${cabinet}`, ["Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);
    await resetSheet(`Отгрузки ${cabinet}`, ["Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);

    console.log(`▶️ Обработка кабинета: ${cabinet}`);
    await getStock(login, password, cabinet);
    await getPurchaseOrders(login, password, cabinet);
    await getShipments(login, password, cabinet);
  }

  await flushBuffers();
  console.log("✅ Готово! Все данные загружены.");
})();
