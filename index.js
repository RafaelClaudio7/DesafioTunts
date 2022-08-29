const express = require("express");
const app = express();
const axios = require("axios"); // importingo axios to realize data fetching
const xl = require("excel4node"); // importing excel lib

const wb = new xl.Workbook(); // Creating the object wb type to generate the sheet
const ws = wb.addWorksheet("sheetName"); // Building the sheet

// Function to get the keys from currencies
function getKeyName(obj) {
  for (let key in obj) {
    return key;
  }
}


// Setting the size of cells

ws.column(1).setWidth(15);
ws.column(2).setWidth(15);
ws.column(3).setWidth(15);
ws.row(1).setHeight(30);

// Styling the cells
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

// Formatting the numbers
var numberStyleFormat = wb.createStyle({
  numberFormat: "#,##0.00;",
});

// Adding sheet title
ws.cell(1, 2).string("Countries List").style(headerStyle);

// Creating an array with the column names
const headingColumnNames = ["Name", "Capital", "Area", "Currencies"];

// Index of columns 
let headingColumnIndex = 1;

headingColumnNames.forEach((heading) => {
  ws.cell(2, headingColumnIndex++)
    .string(heading)
    .style(headingStyle);
});

app.use(express.json());

// Server listening
app.listen(3000, () => {
  console.log("The server is listennig on port 3000");
});

// Root route that handle datas
app.get("/", async (req, res) => {
  const { data } = await axios("https://restcountries.com/v3.1/all"); // Extraindo os dados da api
  const countries = [];
  for (let i = 0; i < data.length; i++) {
    const obj = {
      name: data[i].name.common,
      capital: data[i].capital,
      area: data[i].area,
      moeda: getKeyName(data[i].currencies),
    };
    countries.push(obj); // Adding to an array with all countries
  }
  // sorting the array
  countries.sort(function (a, b) {
    return a.name < b.aome ? -1 : a.aome > b.name ? 1 : 0;
  });

  // log to test
  //console.log(countries[0]);

  // Adding all data to the sheet

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

  // Creating the sheet with the name above
  wb.write("countriesList.xlsx");

  return res.send({ message: "Reload to generate other sheet" });
});
