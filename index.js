const express = require("express");
const app = express();
const axios = require("axios"); // importando o axios para realizar data fetching
const xl = require("excel4node"); // importando o a biblioteca excel4node para gerar a planilha

const wb = new xl.Workbook(); // Criando o obj tipo wb pra cria a planilha
const ws = wb.addWorksheet("Nome da Planilha"); // criando a planilha

// Função para extrair as chaves do Objeto de Currencies
function getKeyName(obj) {
  for (let key in obj) {
    return key;
  }
}

// Ajustando o comprimento das células da planilha

ws.column(1).setWidth(15);
ws.column(2).setWidth(15);
ws.column(3).setWidth(15);
ws.row(1).setHeight(30);

// Ajustando o Estilo das células da planilha
var headerStyle = wb.createStyle({
  font: {
    size: 24,
    bold: true,
    color: "#000000",
  },
});

var headingStyle = wb.createStyle({
  font: {
    size: 12,
    bold: true,
    color: "#808080",
  },
});

// Tornando o padrão adequado dos números
var numberStyleFormat = wb.createStyle({
  numberFormat: "#,##0.00;",
});

// Adicionado o Titulo dos dados da planilha
ws.cell(1, 2).string("Countries List").style(headerStyle);

// Definindo um array com os valores de cabeçalho da planilha
const headingColumnNames = ["Name", "Capital", "Area", "Currencies"];

// Definindo indice para percorrer as colunas através do foreach abaixo e aplicando o estilo adequado
let headingColumnIndex = 1;

headingColumnNames.forEach((heading) => {
  ws.cell(2, headingColumnIndex++)
    .string(heading)
    .style(headingStyle);
});

app.use(express.json());

// Porta 3000utilizada para subir o server
app.listen(3000, () => {
  console.log("O Servidor esta sendo executado na porta 3000");
});

// Rota raiz responsável por gerar a planilha
app.get("/", async (req, res) => {
  const { data } = await axios("https://restcountries.com/v3.1/all"); // Extraindo os dados da api
  const countries = [];
  //Percorrendo os dados da api
  for (let i = 0; i < data.length; i++) {
    const obj = {
      nome: data[i].name.common,
      capital: data[i].capital,
      area: data[i].area,
      moeda: getKeyName(data[i].currencies),
    };
    countries.push(obj); // adicionando ao meu array para colocar na planilha posteriormente
  }
  // ordenando alfabeticamente os paises
  countries.sort(function (a, b) {
    return a.nome < b.nome ? -1 : a.nome > b.nome ? 1 : 0;
  });

  // log para teste
  //console.log(countries[0]);

  // Percorrer o array adicionando os valores as celulas do excel

  let rowIndex = 3;
  countries.forEach((record) => {
    let columnIndex = 1;
    Object.keys(record).forEach((columnName) => {
      if (
        record[columnName] !== undefined &&
        typeof record[columnName] !== "number"
      ) {
        ws.cell(rowIndex, columnIndex++).string(record[columnName]);
      } else if (typeof record[columnName] === "number") {
        ws.cell(rowIndex, columnIndex++)
          .number(record[columnName])
          .style(numberStyleFormat);
      } else {
        ws.cell(rowIndex, columnIndex++).string("-");
      }
    });
    rowIndex++;
  });

  wb.write("countriesList.xlsx");

  return res.send({ message: "Reload to generate other sheet" });
});
