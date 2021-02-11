# Trello to Excell

This program converts a trello board content into an MS Excell spreadsheet. 

## How to use
You need to export your board to JSON and provide the JSON file. To access this feature,
open your board, open the menu, click 'more', and finally 'Print and Export'.

Then, install the package globally:

```npm i -g trello-to-excell```

and run it with `in` and `out` arguments:

```trello2excell --in file.json --out file.xlsx```

By default, the text is written in English. If you want Farsi/Persian text, provide 
the ```--lang fa``` argument.