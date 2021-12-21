const xlsx      = require('xlsx');

let path = "saida/drive.xlsx";
let csvpath = "saida/drive2.csv";
const workBook = xlsx.readFile(path);
xlsx.writeFile(workBook, csvpath, { bookType: "csv" });
console.log('csv criado.')