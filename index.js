import { google } from "googleapis";
import fetch from "node-fetch";
import dotenv from "dotenv";
import fs from "fs";

dotenv.config();

const auth = new google.auth.GoogleAuth({
  keyFile: "credentials.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });
const spreadsheetId = process.env.SPREADSHEET_ID;

// 📄 Чтение таблицы "основа"
async function getCabinets() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "основа!A2:C",
  });

  return res.data.values || [];
}

// 🧹 Очистка и шапка для листа
async function resetSheet(sheetName, headers) {
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
}

// 📝 Добавление строки
async function appendRow(sheetName, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: [row] },
  });
}

// 📦 Получение остатков
async function getStock(login, password, cabinetName) {
  const headers = buildHeaders(login, password);
  const res = await fetch(`https://api.moysklad.ru/api/remap/1.2/report/stock/all?limit=1000`, { headers });
  const json = await res.json();

  for (const row of json.rows || []) {
    const name = row.name || "—";
    const article = row.article || "—";
    const code = row.code || "—";
    const qty = row.quantity || 0;

    await appendRow(`Остатки общее`, [cabinetName, name, article, code, qty]);
    await appendRow(`Остатки ${cabinetName}`, [name, article, code, qty]);
  }
}

// 📋 Позиции заказов
async function getPurchaseOrders(login, password, cabinetName) {
  const headers = buildHeaders(login, password);
  const res = await fetch(`https://api.moysklad.ru/api/remap/1.2/entity/purchaseorder?expand=positions,agent,state&limit=100`, { headers });
  const json = await res.json();
  const cache = {};

  for (const order of json.rows || []) {
    const status = order.state?.name?.toLowerCase() || "";
    if (status.includes("доставлено")) continue;

    const date = order.moment?.split("T")[0] || "—";
    const agent = order.agent?.name || "—";
    const state = order.state?.name || "—";

    for (const pos of order.positions?.rows || []) {
      const href = pos.assortment?.meta?.href;
      const { name, article, code } = await getProduct(href, headers, cache);
      const qty = pos.quantity || 0;

      await appendRow(`ПозицииЗаказов общее`, [cabinetName, date, agent, state, name, article, code, qty]);
      await appendRow(`ПозицииЗаказов ${cabinetName}`, [date, agent, state, name, article, code, qty]);
    }
  }
}

// 🚚 Отгрузки
async function getShipments(login, password, cabinetName) {
  const headers = buildHeaders(login, password);
  const res = await fetch(`https://api.moysklad.ru/api/remap/1.2/entity/demand?expand=positions,agent,state&limit=100`, { headers });
  const json = await res.json();
  const cache = {};

  for (const ship of json.rows || []) {
    const status = ship.state?.name?.toLowerCase() || "";
    if (status.includes("поступило в продажу")) continue;

    const date = ship.moment?.split("T")[0] || "—";
    const agent = ship.agent?.name || "—";
    const state = ship.state?.name || "—";

    for (const pos of ship.positions?.rows || []) {
      const href = pos.assortment?.meta?.href;
      const { name, article, code } = await getProduct(href, headers, cache);
      const qty = pos.quantity || 0;

      await appendRow(`Отгрузки общее`, [cabinetName, date, agent, state, name, article, code, qty]);
      await appendRow(`Отгрузки ${cabinetName}`, [date, agent, state, name, article, code, qty]);
    }
  }
}

// 🧠 Подгрузка товара по href
async function getProduct(href, headers, cache) {
  if (cache[href]) return cache[href];
  const res = await fetch(href, { headers });
  const json = await res.json();
  const result = {
    name: json.name || "—",
    article: json.article || "—",
    code: json.code || "—",
  };
  cache[href] = result;
  return result;
}

// 🔐 Авторизация в МойСклад
function buildHeaders(login, password) {
  const auth = Buffer.from(`${login}:${password}`).toString("base64");
  return {
    "Authorization": `Basic ${auth}`,
    "Content-Type": "application/json",
    "User-Agent": "mysklad-sync-bot",
  };
}

// 🧩 Основной запуск
(async () => {
  const cabinets = await getCabinets();

  // Очистка глобальных листов
  await resetSheet("Остатки общее", ["Кабинет", "Наименование", "Артикул", "Код", "Остаток"]);
  await resetSheet("ПозицииЗаказов общее", ["Кабинет", "Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);
  await resetSheet("Отгрузки общее", ["Кабинет", "Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);

  for (const [cabinetName, login, password] of cabinets) {
    await resetSheet(`Остатки ${cabinetName}`, ["Наименование", "Артикул", "Код", "Остаток"]);
    await resetSheet(`ПозицииЗаказов ${cabinetName}`, ["Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);
    await resetSheet(`Отгрузки ${cabinetName}`, ["Дата", "Контрагент", "Статус", "Товар", "Артикул", "Код", "Количество"]);

    console.log(`📡 Обработка: ${cabinetName}`);
    await getStock(login, password, cabinetName);
    await getPurchaseOrders(login, password, cabinetName);
    await getShipments(login, password, cabinetName);
  }

  console.log("✅ Синхронизация завершена.");
})();
