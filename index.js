function oprosXLSX(puNumber, workbook) {
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];    
  let ws_arr_columnA = Object.keys(worksheet).filter(key => key[0] === "A");  
  result = [];
  let cell = null;
  let value = null;
  let xlsx_row = null;
  let col_odate = 7 // номер столбца даты опроса
  for (let cell_index in ws_arr_columnA) {  
    cell = ws_arr_columnA[cell_index]
    value = worksheet[cell] ? worksheet[cell].v : "";
    if (value * 1 === puNumber * 1) {
      xlsx_row = XLSX.utils.decode_cell(cell).r;
      if (worksheet[XLSX.utils.encode_cell({ c: col_odate, r: xlsx_row})] === undefined) {
        result.push({
          sn: value, 
          odate: ' ',
          ptype: worksheet[XLSX.utils.encode_cell({ c: 1, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 1, r: xlsx_row})].v : "",
          svnum: worksheet[XLSX.utils.encode_cell({ c: 2, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 2, r: xlsx_row})].v : "",
          login: worksheet[XLSX.utils.encode_cell({ c: 3, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 3, r: xlsx_row})].v : "",
          route: worksheet[XLSX.utils.encode_cell({ c: 5, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 5, r: xlsx_row})].v : "", 
          uspdn: worksheet[XLSX.utils.encode_cell({ c: 6, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 6, r: xlsx_row})].v : "",
        })
      }
      else {
        result.push({
          sn: value, 
          odate: worksheet[XLSX.utils.encode_cell({ c: col_odate, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: col_odate, r: xlsx_row})].v : "",
          ptype: worksheet[XLSX.utils.encode_cell({ c: 1, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 1, r: xlsx_row})].v : "",
          svnum: worksheet[XLSX.utils.encode_cell({ c: 2, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 2, r: xlsx_row})].v : "",
          login: worksheet[XLSX.utils.encode_cell({ c: 3, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 3, r: xlsx_row})].v : "",
          route: worksheet[XLSX.utils.encode_cell({ c: 5, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 5, r: xlsx_row})].v : "", 
          uspdn: worksheet[XLSX.utils.encode_cell({ c: 6, r: xlsx_row})] ? worksheet[XLSX.utils.encode_cell({ c: 6, r: xlsx_row})].v : "",
        }) 
      }
    }
  }
  return result  //массив объектов
}

//======разделитель=========//

require('dotenv').config();
const {Bot, Api} = require('grammy');
const bot = new Bot(process.env.BOT_API_KEY)

const XLSX = require('xlsx');
//  let startTime = Date.now();
const workbook = XLSX.readFile("data/ROL.xlsx");
//  let endTime = Date.now();
//  console.log(endTime-startTime); //только опрос 3326ms, c параметрами 8385ms

bot.command('start', async (ctx) => {
  await ctx.reply('Приветствую, коллеги! Напишите номер ПУ для проверки последней даты опроса')
})

// bot.hears(['С днём рождения', 'С днем рождения', 'С днюхой'], async (ctx) => {
//   await ctx.replyWithSticker("CAACAgIAAxkBAAENWLBnYsRvGdASQm7P5k44rcIkq70T8QACOgADr8ZRGutCYzxwMcBJNgQ")
// })

bot.on('message', async (ctx) => {
  try {
    let mesResult = oprosXLSX(ctx.message.text, workbook);
    //console.log(mesResult);
    if (mesResult.length == 0) { 
      console.log(mesResult);      
      await ctx.reply(`${ctx.message.text} == не найден`)
    } 
    else {
      for (let i in mesResult) {
        if (mesResult[i]) {
          if (mesResult[i].odate === ' ') {
            await ctx.reply(`${mesResult[i].sn} == не опрос
 == ${mesResult[i].ptype} == ${mesResult[i].svnum} == ${mesResult[i].login} == ${mesResult[i].route} == ${mesResult[i].uspdn}`)
          }
          else {
            await ctx.reply(`${mesResult[i].sn} == опрос ${mesResult[i].odate}
 == ${mesResult[i].ptype} == ${mesResult[i].svnum} == ${mesResult[i].login} == ${mesResult[i].route} == ${mesResult[i].uspdn}`)
          }
        }
        else await ctx.reply(`${ctx.message.text} == не найден`)
      }
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