const fs = require('fs')
const xl = require('excel4node')

const wb = new xl.Workbook('')
const ws = wb.addWorksheet('Documentation');
let row = 2 
let col = 2

function appendForwardSlash (string) {
  if (string[string.length - 1] == '/') {
    return string
  } else {
    return string + '/'
  }
}

function check(value) {
  return ((value)? value : '').toString()
}

const style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 12,
  }
});

function generateExcel (filename, flags){
  ws.cell(1, 1).string('Folder').toString()
  ws.cell(1, 2).string('Name').toString()
  ws.cell(1, 3).string('Method').toString()
  ws.cell(1, 4).string('Path').toString()
  ws.cell(1, 5, 1, 6, true).string('Headers').toString()
  ws.cell(1, 7, 1, 8, true).string('Query').toString()
  ws.cell(1, 9).string('Body').toString()
  
  let file = fs.readFileSync(filename);
  if (!file.length || !filename) {
    console.log('No File Found')
    return 0
  }
  let collectionObject = JSON.parse(file);
  parseFolder(collectionObject)
  wb.write(`${filename.slice(0, filename.length - 5)}.xlsx`)
  console.log(`Excel file generated with name ${filename.slice(0, filename.length - 5)}.xlsx`)
}

function parseItem(item) {
  let queryRow = row
  let bodyRow = row
  let headerRow = row
  // Name
  ws.cell(row, col).string(check(item.name))
  
  // Method
  ws.cell(row, ++col).string(check(item.request.method))

  // Path
  ws.cell(row, ++col).string(appendForwardSlash((((item.request.url.path)?item.request.url.path:['']).join('/')).toString()))

  // Headers
  if (item.request.header.length) {
    for (let h of item.request.header) {
      ws.cell(headerRow, col+1).string(check(h.key))
      ws.cell(headerRow, col+2).string(check(h.value))
      headerRow++
    }
  }

  // Query Params
  if (item.request.url.query != undefined) {
    for (let h of item.request.url.query) {
      ws.cell(queryRow, col+3).string(check(h.key))
      ws.cell(queryRow, col+4).string(check(h.value))
      queryRow++
    }
  }

  // Request Body
  if (item.request.body != undefined && item.request.body.raw != undefined) {
    ws.cell(bodyRow, col+5).string(check(item.request.body.raw))
  }

  row = Math.max(headerRow, queryRow, bodyRow) + 1
}

function parseFolder (folder) {
  if (folder.hasOwnProperty('request')) {
    parseItem(folder)
    col = 2
  } else {
    if (folder.name){
      ws.cell(row, 1).string(check(folder.name))
    }
    row++
    for (let item of folder.item) {
      parseFolder(item)
    }
  }
}

module.exports = generateExcel

