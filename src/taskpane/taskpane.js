/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("btnSearch").onclick = search;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function search() {
  let keyword = document.getElementById("searchBox").value;
  Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getRange("A1:A100");
      let startCell = context.workbook.getSelectedRange();
      //const rangeRS = sheet.getRange("D5");
      range.load("values");
      await context.sync();
      let result = range.values.filter(r =>
          r[0].toLowerCase().includes(keyword.toLowerCase())
      );
      //document.getElementById("result").innerText = JSON.stringify(result);
      //rangeRS.values=JSON.stringify(result);
      //console.log(result)
      for (let i = 0; i < result.length; i++) {
          let cell = startCell.getOffsetRange(i, 0);
          cell.values = [[result[i][0]]];
      }
      
  });
}

async function searchIngredient() {
 const keyword =
  document.getElementById("searchBox").value.toLowerCase();
 await Excel.run(async (context) => {
  const sheet =
   context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A2:A1000");
  range.load("values");
  await context.sync();
  const data = range.values;
  const results = data.filter(r =>
     r[0].toLowerCase().includes(keyword)
  );
  const list =
   document.getElementById("resultList");
  list.innerHTML = "";
  results.slice(0,10).forEach(r => {
    const li = document.createElement("li");
    li.innerText = r[0];
    list.appendChild(li);
  });
 });
}