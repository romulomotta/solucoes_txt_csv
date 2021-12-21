const readline  = require('readline');
const fs        = require('fs')
const xl        = require('excel4node');
const LineReader = require('line-by-line');

const lr = new LineReader('entrada/drive.txt')

const wb        = new xl.Workbook();
const ws        = wb.addWorksheet('Worksheet Name');

const headingColumnNames = [
    "ID",
    "Data",
    "CNPJ_FP",
    "Nome_FP",
    "Tipo_Pessoa",
    "CPF_CNPJ_Cliente",
    "Nome_Cliente",
    "Carteira",
    "Retencao",
    "Valor_Rendimento",
    "Valor_IR"
]

let rowIndex = 2
let headingColumnIndex = 1; //diz que começará na primeira linha
headingColumnNames.forEach(heading => { //passa por todos itens do array
    // cria uma célula do tipo string para cada título
    ws.cell(1, headingColumnIndex++).string(heading);
});

/*
    "Data",              99  106 08
    "CNPJ_FP",           01 014 14
    "Nome_FP",           39 098 60
    "Tipo_Pessoa",       25 038 11(cpf) 14(cnpj)
    "CPF_CNPJ_Cliente",  25 038 14
    "Nome_Cliente",      39 098 60
    "Carteira",         380 389 09
    "Retencao",         569 576 08
    "Valor_Rendimento", 107 123 17
    "Valor_IR"          124 140 17
*/
let counter = 0
console.log("Processo iniciado: " + Date());

lr.on('line', function (line) {
    lr.pause();

    counter ++;

    let id          = counter.toString();
    let data_mov    = line.substring(98, 106);
    let cnpj_fp     = line.substring(0,016);
    let nome_fp     = "Votorantim Asset Management";
    let cnpj_cpf    = line.substring(24, 38);
    let tpessoa     = cnpj_cpf.substring(12,14) > 1 
                        ? "PJ" : "PF";
    let nome        = line.substring(38, 98);
    let carteira    = line.substring(379, 388);
    let retencao    = line.substring(568, 576);
    let vlr_rend    = line.substring(106, 123);
    let vlr_ir      = line.substring(123, 140);
    
    let data = [
        {
            "ID": id,
            "Data": convertDate(data_mov),            
            "CNPJ_FP": cnpj_fp,         
            "Nome_FP": nome_fp,         
            "Tipo_Pessoa": tpessoa,     
            "CPF_CNPJ_Cliente": cnpj_cpf,
            "Nome_Cliente": nome,
            "Carteira": carteira,    
            "Retencao": retencao,        
            "Valor_Rendimento": convertMoney(vlr_rend),
            "Valor_IR": convertMoney(vlr_ir)         
        }
    ]

    setTimeout(function () {
        // 
        lr.resume();
    }, 100);

    if (counter %100 === 0){
        console.log(counter);
    }
    
    toExcel(data);
});

lr.on('end', function() {
    console.log("Processo encerrado: " + Date());
})


function convertMoney(money) {
    let num = Number(money);
    let decimals = num /100
    let converted = decimals.toLocaleString('en-US', {maximumFractionDigits: 2});
    
    return converted;
}

function convertDate(date) {
    let year = date.substring(4);
    let month = date.substring(2,4);
    let day = date.substring(0, 2);

    return `${day}/${month}/${year}`;
}

function toExcel(data) { 
    let path = "saida/drive.xlsx";

    data.forEach( record => {
        let columnIndex = 1;
        Object.keys(record).forEach(columnName =>{
            ws.cell(rowIndex,columnIndex++)
                .string(record [columnName])
        });
        rowIndex++;
    }); 
     
    wb.write(path);
}