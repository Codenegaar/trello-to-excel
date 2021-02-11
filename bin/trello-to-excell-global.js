#!/usr/bin/env node

const trelloToExcell = require("../lib");
const yargs = require("yargs/yargs");
const { hideBin } = require("yargs/helpers");
const argv = yargs(hideBin(process.argv)).argv;

const start = async () => {
    await trelloToExcell.convert(argv.in, argv.out, argv.lang)
};

start();