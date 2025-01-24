function normSN(sn) {
  if (isNaN(Number(sn))) return sn
  else return Number(sn)
}

function oprosXLSX(puNumber, workbook) {
  result = [];

  let worksheet = workbook.Sheets[workbook.SheetNames[0]];  
  let ws_arr_columnA = Object.keys(worksheet).filter(key => key[0] === "A");  
  let cell = null;
  let value = null;
  let xlsx_row = null;

  let worksheet2 = workbook.Sheets[workbook.SheetNames[1]];
  let ws_arr_columnA2 = Object.keys(worksheet2).filter(key => key[0] === "A");
  let cell2 = null;
  let value2 = null;
  let xlsx_row2 = null;  

  let odate_col = 7 // номер столбца даты опроса
  for (let cell_index in ws_arr_columnA) {  
    cell = ws_arr_columnA[cell_index]
    value = worksheet[cell] ? worksheet[cell].v : "";
    if (normSN(value) === normSN(puNumber)) {
      xlsx_row = XLSX.utils.decode_cell(cell).r;      
        result.push({
          sn: value, 
          odate: worksheet[XLSX.utils.encode_cell({ c: odate_col, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: odate_col, r: xlsx_row})].v : ' ',
          ptype: worksheet[XLSX.utils.encode_cell({ c: 1, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 1, r: xlsx_row})].v : "",
          svnum: worksheet[XLSX.utils.encode_cell({ c: 2, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 2, r: xlsx_row})].v : "",
          login: worksheet[XLSX.utils.encode_cell({ c: 3, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 3, r: xlsx_row})].v : "",
          route: worksheet[XLSX.utils.encode_cell({ c: 5, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 5, r: xlsx_row})].v : "", 
          uspdn: worksheet[XLSX.utils.encode_cell({ c: 6, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 6, r: xlsx_row})].v : "",
        });
        if (result[result.length-1].uspdn) {
          let uspdNumber = result[result.length-1].uspdn.split(' ', 1)[0];
          for (let cell_index2 in ws_arr_columnA2) {
            cell2 = ws_arr_columnA[cell_index2]
            value2 = worksheet2[cell2] ? worksheet2[cell2].v : "";
            if (normSN(value2) === normSN(uspdNumber)) {
              xlsx_row2 = XLSX.utils.decode_cell(cell2).r;
              result[result.length-1].uspdr = worksheet2[XLSX.utils.encode_cell({ c: 5, r: xlsx_row2})] ? worksheet2[XLSX.utils.encode_cell({ c: 5, r: xlsx_row2})].v : ""
            }
          }
        }
    }
  }
  return result  //массив объектов
}

//======разделитель=========//

require('dotenv').config();
const {Bot, Api} = require('grammy');
const bot = new Bot(process.env.BOT_API_KEY);
const accessgroup = process.env.ACCESS_GROUP_ID;
const admin = process.env.ADMIN_ID;

const XLSX = require('xlsx');
//  let startTime = Date.now();
const workbook = XLSX.readFile("./data/ROL.xlsx");
//  let endTime = Date.now();
//  console.log(endTime-startTime); //только опрос 3326ms, c параметрами 8385ms

bot.command('start', async (ctx) => {
  await ctx.reply('Приветствую, коллеги! Напишите номер ПУ для проверки последней даты опроса')
})

// bot.hears(['С днём рождения', 'С днем рождения', 'С днюхой'], async (ctx) => {
//   await ctx.replyWithSticker("CAACAgIAAxkBAAENWLBnYsRvGdASQm7P5k44rcIkq70T8QACOgADr8ZRGutCYzxwMcBJNgQ")
// })

bot.command('sn', async (ctx) => {
  try {
    await bot.api.sendMessage(admin, `Запрос ${ctx.message.text} от ${ctx.message.from.id} ${ctx.message.from.first_name}`)  // сообщение на ADMIN_ID
    let access = (await bot.api.getChatMember(accessgroup, ctx.message.from.id)).status;
    if (access == 'creator' || access == 'administrator' || access == 'member') {
      let mesResult = oprosXLSX(ctx.match, workbook);
      let uspdr_try = "";
      if (mesResult.length == 0) {
        await ctx.reply(`${ctx.match} == не найден`)
      } 
      else {
        for (let i in mesResult) {
          if (mesResult[i]) {
            uspdr_try = mesResult[i].uspdr ? mesResult[i].uspdr : "";
            if (mesResult[i].odate === ' ') {
              await ctx.reply(`${mesResult[i].sn} == не опрос
 == ${mesResult[i].ptype} == ${mesResult[i].svnum} == ${mesResult[i].login} == ${mesResult[i].route} == ${mesResult[i].uspdn} == ${uspdr_try}`)
            }
            else {
              await ctx.reply(`${mesResult[i].sn} == опрос ${mesResult[i].odate}
 == ${mesResult[i].ptype} == ${mesResult[i].svnum} == ${mesResult[i].login} == ${mesResult[i].route} == ${mesResult[i].uspdn} == ${uspdr_try}`)
            }
          }
          else await ctx.reply(`${ctx.match} == не найден`)
        }
      }      
    } else { 
      await ctx.reply(`Вы не состоите в группе ${(await bot.api.getChat(accessgroup)).title}`)      
    }
  } catch (error) {
    console.error(error)
  }
})

bot.on('message').filter(
  async (ctx) => ctx.message.chat.type === "private",
  async (ctx) => {
  try {
    await bot.api.sendMessage(admin, `Запрос ${ctx.message.text} от ${ctx.message.from.id} ${ctx.message.from.first_name}`)  // сообщение на ADMIN_ID
    let access = (await bot.api.getChatMember(accessgroup, ctx.message.from.id)).status;
    if (access == 'creator' || access == 'administrator' || access == 'member') {
      let mesResult = oprosXLSX(ctx.message.text, workbook);
      let uspdr_try = "";
      if (mesResult.length == 0) {
        await ctx.reply(`${ctx.message.text} == не найден`)
      } 
      else {
        for (let i in mesResult) {
          if (mesResult[i]) {
            uspdr_try = mesResult[i].uspdr ? mesResult[i].uspdr : "";
            if (mesResult[i].odate === ' ') {
              await ctx.reply(`${mesResult[i].sn} == не опрос
 == ${mesResult[i].ptype} == ${mesResult[i].svnum} == ${mesResult[i].login} == ${mesResult[i].route} == ${mesResult[i].uspdn} == ${uspdr_try}`)
            }
            else {
              await ctx.reply(`${mesResult[i].sn} == опрос ${mesResult[i].odate}
 == ${mesResult[i].ptype} == ${mesResult[i].svnum} == ${mesResult[i].login} == ${mesResult[i].route} == ${mesResult[i].uspdn} == ${uspdr_try}`)
            }
          }
          else await ctx.reply(`${ctx.message.text} == не найден`)
        }
      }
    } else { 
      await ctx.reply(`Вы не состоите в группе ${(await bot.api.getChat(accessgroup)).title}`)      
    }
  } catch (error) {
    console.error(error)
  }
})

bot.catch((err) => {
  const ctx = err.ctx;
  const e = err.error;
  console.error("Error", e)
})

bot.start();