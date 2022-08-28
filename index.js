const express = require("express");
const app = express();
const axios = require("axios");
const xl = require("excel4node");


const wb = new xl.Workbook(); // Criando o obj tipo wb pra cria a planilha
const ws = wb.addWorksheet("Nome da Planilha"); // criando a planilha

const titleSheet = ["Countries List"];

ws.cell(1, 1).string(titleSheet[0]);

const headingColumnNames = ["Name", "Capital", "Area", "Currencies"];

let headingColumnIndex = 1; // Para informar que será escrito na primeira linha da planilha

headingColumnNames.forEach((heading) => {
  ws.cell(2, headingColumnIndex++).string(heading);
});

// Create a reusable style
var style = wb.createStyle({
  numberFormat: '#,##0.00;',
});



function getKeyName (obj) {
    for(let  key in obj){
        return key;
    }
}

app.use(express.json());

app.listen(3000, () => {
  console.log("O Servidor esta sendo executado na porta 3000");
});

app.get("/", async (req, res) => {
  const { data } = await axios("https://restcountries.com/v3.1/all");
  const countries = [];
  for (let i = 0; i < data.length; i++) {
    const obj = {
      nome: data[i].name.common,
      capital: data[i].capital,
      area: data[i].area,
      moeda: getKeyName(data[i].currencies),
    };
    countries.push(obj);
  }
  countries.sort(function (a, b) {
    return a.nome < b.nome ? -1 : a.nome > b.nome ? 1 : 0;
  });

  console.log(countries[0]);



  
  let rowIndex = 3;
  countries.forEach((record) => {
    let columnIndex = 1;
    Object.keys(record).forEach((columnName) => {
      if(record[columnName] !== undefined && typeof(record[columnName]) !== 'number'){
      ws.cell(rowIndex, columnIndex++).string(record[columnName]);
      }else if(typeof(record[columnName]) ===  'number'){
        ws.cell(rowIndex, columnIndex++).number(record[columnName]).style(style);
      }  
      
      else {
        ws.cell(rowIndex, columnIndex++).string('-');
      }
    });
    rowIndex++;
  });

   wb.write("teste.xlsx");


  return res.send({ message: "Reload to generate other datas" });
});
