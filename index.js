#!/usr/bin/env node

const meow = require('meow');
const generateExcel = require('./generateExcel.js')
const cli = meow(`
  Usage
    $ ptoe <input>
  Options
    --file, -f Path to collection.json file
  Examples
    $ ptoe --file collection.json
    Excel file generated
`, {
  flags: {
    file: {
      type: 'boolean',
        alias: 'f'
    }
  }
});

generateExcel(cli.input[0], cli.flags)
