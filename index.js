const TelegramBot = require("node-telegram-bot-api");
const XLSX = require("xlsx");
const fs = require("fs");

// Telegram Bot Token from env file
const token = process.env.TOKEN;
const bot = new TelegramBot(token, { polling: true });

function getResultFromExcel(code) {
  const workbook = XLSX.readFile("results.xlsx");
  const sheetName = workbook.SheetNames[0];
  const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  const result = sheet.find((row) => row["code"] === parseInt(code));
  if (result) {
    return `نتيجتك هي: ${result["Result"]}`;
  } else {
    return "لم يتم العثور على نتيجتك.";
  }
}

bot.on("message", (msg) => {
  const chatId = msg.chat.id;
  const code = msg.text.trim();

  try {
    const result = getResultFromExcel(code);
    bot.sendMessage(chatId, result);
  } catch (error) {
    bot.sendMessage(chatId, "حدث خطأ أثناء جلب النتيجة.");
    console.error(error);
  }
});
