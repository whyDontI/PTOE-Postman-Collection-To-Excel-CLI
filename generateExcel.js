const fs = require('fs');
const xl = require('excel4node')

const wb = new xl.Workbook('')
const ws = wb.addWorksheet('Documentation');
let row = 2 
let col = 2

const style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 12,
  }
});

function generateExcel (filename, flags){
  ws.cell(1, 1).string('Folder').toString()
  ws.cell(1, 2).string('Method').toString()
  ws.cell(1, 3).string('Name').toString()
  ws.cell(1, 4).string('Path').toString()
  ws.cell(1, 5, 1, 6, true).string('Headers').toString()
  ws.cell(1, 7, 1, 8, true).string('Query').toString()
  ws.cell(1, 9).string('Body').toString()
  
  let file = fs.readFileSync(filename);
  let collectionObject = JSON.parse(file);
  parseFolder(collectionObject)
  wb.write(`${filename}.xlsx`)
}

function parseItem(item) {
  let queryRow = row
  let bodyRow = row
  let headerRow = row
  // Name
  ws.cell(row, col).string(((item.name)?item.name:'').toString())
  
  // Method
  ws.cell(row, ++col).string(((item.request.method)?item.request.method:'').toString())

  // Path
  ws.cell(row, ++col).string((((item.request.url.path)?item.request.url.path:['']).join('/')).toString())

  // Headers
  if (item.request.header.length) {
    for (let h of item.request.header) {
      ws.cell(headerRow, col+1).string((h.key).toString())
      ws.cell(headerRow, col+2).string((h.value).toString())
      headerRow++
    }
  }

  // Query Params
  if (item.request.url.query != undefined) {
    for (let h of item.request.url.query) {
      ws.cell(queryRow, col+3).string(((h.key)?h.key:'').toString())
      ws.cell(queryRow, col+4).string(((h.value)?h.value:'').toString())
      queryRow++
    }
  }

  // Request Body
  if (item.request.body != undefined && item.request.body.raw != undefined) {
    ws.cell(bodyRow, col+5).string(item.request.body.raw)
  }

  row = Math.max(headerRow, queryRow, bodyRow) + 1
}

function parseFolder (folder) {
  if (folder.hasOwnProperty('request')) {
    parseItem(folder)
    col = 2
  } else {
    if (folder.name){
      ws.cell(row, 1).string((folder.name).toString())
    }
    row++
    for (let item of folder.item) {
      parseFolder(item)
    }
  }
}

module.exports = generateExcel

