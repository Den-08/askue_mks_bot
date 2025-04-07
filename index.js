// Описание функций
// функция обработки серийного номера
function normSN(sn) {
  if (isNaN(Number(sn))) return sn
  else return Number(sn)
}

// функция поиска по XLSX
function oprosXLSX(puNumber, workbook) {
  result = [];
  // константы и переменные для Листа №1
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const ws_arr_keyColumn = Object.keys(worksheet).filter(key => key[0] === "B");  // ключевой столбец "B"
  let cell = null;
  let value = null;
  let xlsx_row = null;
  // константы и переменные для Листа №2
  const worksheet2 = workbook.Sheets[workbook.SheetNames[1]];
  const ws_arr_keyColumn2 = Object.keys(worksheet2).filter(key => key[0] === "B");  // ключевой столбец "B"
  let cell2 = null;
  let value2 = null;
  let xlsx_row2 = null;
  // номера столбцов на листах Excel начиная с 0
  const cNum = { odate: 6, ptype: 0, svnum: 2, login: 3, route: 4, uspdn: 5, uspdr: 2 };
  let uspdNumber = null;
  for (let cell_index of ws_arr_keyColumn) {
    cell = cell_index;
    value = worksheet[cell] ? worksheet[cell].v : "";
    if (normSN(value) === normSN(puNumber)) {
      xlsx_row = XLSX.utils.decode_cell(cell).r;
      result.push({
        sn: value,
        odate: worksheet[XLSX.utils.encode_cell({ c: cNum.odate, r: xlsx_row })] ? worksheet[XLSX.utils.encode_cell({ c: cNum.odate, r: xlsx_row })].v : ' ',
        ptype: worksheet[XLSX.utils.encode_cell({ c: cNum.ptype, r: xlsx_row })] ? worksheet[XLSX.utils.encode_cell({ c: cNum.ptype, r: xlsx_row })].v : "",
        svnum: worksheet[XLSX.utils.encode_cell({ c: cNum.svnum, r: xlsx_row })] ? worksheet[XLSX.utils.encode_cell({ c: cNum.svnum, r: xlsx_row })].v : "",
        login: worksheet[XLSX.utils.encode_cell({ c: cNum.login, r: xlsx_row })] ? worksheet[XLSX.utils.encode_cell({ c: cNum.login, r: xlsx_row })].v : "",
        route: worksheet[XLSX.utils.encode_cell({ c: cNum.route, r: xlsx_row })] ? worksheet[XLSX.utils.encode_cell({ c: cNum.route, r: xlsx_row })].v : "",
        uspdn: worksheet[XLSX.utils.encode_cell({ c: cNum.uspdn, r: xlsx_row })] ? worksheet[XLSX.utils.encode_cell({ c: cNum.uspdn, r: xlsx_row })].v : "",
      });
      if (result[result.length - 1].uspdn) {
        uspdNumber = result[result.length - 1].uspdn.split(' ', 1)[0];
        for (let cell_index2 of ws_arr_keyColumn2) {
          cell2 = cell_index2;
          value2 = worksheet2[cell2] ? worksheet2[cell2].v : "";
          if (normSN(value2) === normSN(uspdNumber)) {
            xlsx_row2 = XLSX.utils.decode_cell(cell2).r;
            result[result.length - 1].uspdr = worksheet2[XLSX.utils.encode_cell({ c: cNum.uspdr, r: xlsx_row2 })] ? worksheet2[XLSX.utils.encode_cell({ c: cNum.uspdr, r: xlsx_row2 })].v : ""
          }
        }
      }
    }
  }
  return result  //массив объектов
}

// функция вывода в сообщение в формате с == разделителями
async function sendToMessage(ctx, mesResult_i, uspdr_try) {
  if (mesResult_i.odate === ' ') {
    await ctx.reply(`${mesResult_i.sn} == не опрос
== ${mesResult_i.ptype} == ${mesResult_i.svnum} == ${mesResult_i.login} == ${mesResult_i.route} == ${mesResult_i.uspdn} == ${uspdr_try}`)
  }
  else {
    await ctx.reply(`${mesResult_i.sn} == опрос ${mesResult_i.odate}
== ${mesResult_i.ptype} == ${mesResult_i.svnum} == ${mesResult_i.login} == ${mesResult_i.route} == ${mesResult_i.uspdn} == ${uspdr_try}`)
  }
}

// функция поиска и вывода информации по серийному номеру ПУ
let getInfoBySN = async (ctx, json_key = false) => {
  if (checkMessage(ctx)) {
    let textOrMatch = null;
    if (ctx.match) { textOrMatch = ctx.match } else { textOrMatch = ctx.message.text }
    let mesResult = oprosXLSX(textOrMatch, workbook);
    let uspdr_try = "";
    // начало вывода результата в ответное сообщение
    if (mesResult.length == 0) {
      await ctx.reply(`${textOrMatch} == не найден`)
    }
    else {
      // вывод значений массива в цикле
      for (let mesResult_i of mesResult) {
        if (mesResult_i) {
          uspdr_try = mesResult_i.uspdr ? mesResult_i.uspdr : "";
          json_key ? await ctx.reply(JSON.stringify(mesResult_i, null, 5)) : sendToMessage(ctx, mesResult_i, uspdr_try);
        }
        else await ctx.reply(`${textOrMatch} == не найден`)
      }
    }
  }
}

// функция проверки сообщения
async function checkMessage(ctx) {
  // сообщение администратору
  try {
    await bot.api.sendMessage(admin, `${ctx.message.text} запрос от @${ctx.message.from.username} ${ctx.message.from.id} ${ctx.message.from.first_name}`);
  } catch (error) {
    console.error(error);
  }
  let textOrMatch = null;
  if (ctx.match) { textOrMatch = ctx.match } else { textOrMatch = ctx.message.text }
  try {
    // проверка доступа по группе
    let access = (await bot.api.getChatMember(accessgroup, ctx.message.from.id)).status;
    if (access === 'creator' || access === 'administrator' || access === 'member') {
      return true
    } else {
      await ctx.reply(`Вы не состоите в группе ${(await bot.api.getChat(accessgroup)).title}`)
      return false
    }
  } catch (error) {
    console.error(error)
  }
}

//======разделитель=========//
require('dotenv').config();
const { Bot, Api } = require('grammy');
const XLSX = require('xlsx');
const { EventEmitter } = require('events');
EventEmitter.defaultMaxListeners = 20;

// Инициализация констант
const bot = new Bot(process.env.BOT_API_KEY);
const accessgroup = process.env.ACCESS_GROUP_ID;
const admin = process.env.ADMIN_ID;
const xlsxRoute = process.env.XLSX_ROUTE;

// Кэширование книги xlsx
// let startTime = Date.now();
let workbook = XLSX.readFile(xlsxRoute);
// let endTime = Date.now();

// Команды бота
bot.command('start', async (ctx) => {
  ctx.react("✍️"); // отмечаем сообщение реакцией
  // сообщение администратору о запросе
  try {
    await bot.api.sendMessage(admin, `${ctx.message.text} запрос от @${ctx.message.from.username} ${ctx.message.from.id} ${ctx.message.from.first_name}`);
  } catch (error) {
    console.error(error);
  }
  await ctx.reply('Приветствую, коллеги! Напишите номер ПУ для проверки последней даты опроса')
})

bot.command('sn', async (ctx) => {
  ctx.react("✍️"); // отмечаем сообщение реакцией
  getInfoBySN(ctx)
})

bot.command('json', async (ctx) => {
  ctx.react("✍️"); // отмечаем сообщение реакцией
  let json_key = true;
  getInfoBySN(ctx, json_key)
})

bot.command('test_uspdn', async (ctx) => {
  ctx.react("✍️"); // отмечаем сообщение реакцией
  if (checkMessage(ctx)) {
    let textOrMatch = null;
    if (ctx.match) { textOrMatch = ctx.match } else { textOrMatch = ctx.message.text }
    // начало
    let mesResult = oprosXLSX(textOrMatch, workbook); // получаем массив найденных ПУ
    let jsonResult = null;      
    if (mesResult.length == 0) {
      await ctx.reply(`${textOrMatch} == не найден`)
    }
    else {
      // вывод значений массива в цикле
      for (let mesResult_i of mesResult) {
        if (mesResult_i) {
          jsonResult = JSON.stringify(mesResult_i, null, 5);
          ctx.reply(JSON.parse(jsonResult).uspdn)
        }
        else await ctx.reply(`${textOrMatch} == не найден`)
      }
    }
    // конец 
  };
})

// bot.command('reloadxlsx', async (ctx) => {
//   // сообщение администратору о запросе
//   try {
//     await bot.api.sendMessage(admin, `${ctx.message.text} запрос от @${ctx.message.from.username} ${ctx.message.from.id} ${ctx.message.from.first_name}`);
//   } catch (error) {
//     console.error(error);
//   }
//   if (ctx.message.from.id.toString() == admin.toString()) {
//     let startTime = Date.now(); // начало выполнения
//     let workbook = XLSX.readFile(xlsxRoute);
//     let endTime = Date.now(); // конец выполнения
//     await ctx.reply(`Файл .xlsx перезагружен за ${endTime - startTime}мс`).catch(error => { console.error(error) }); // время выполнения в мс
//   } else {
//     await ctx.reply('Вы не админ').catch(err => { console.error(err) });;
//   }
// })

// Обработка ошибок
bot.catch((error) => {
  console.error(`Глобальная ошибка: ${error.error.message}`)
})

// Запуск бота
// отвечать только на сообщения в личку
bot.chatType("private").on('message', async (ctx) => {
  ctx.react("✍️"); // отмечаем сообщение реакцией
  // let json_key = false;
  getInfoBySN(ctx)
});

bot.start();