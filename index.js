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
    $ ptoe --help
    Excel file generated
`, {
  flags: {
    file: {
      type: 'boolean',
        alias: 'f'
    },
    help: {
      type: 'boolean',
      alias: 'h'
    }
  }
});

generateExcel(cli.input[0], cli.flags)
