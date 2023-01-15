import { chromium } from "playwright";
import async from "async";
import XLSX from "xlsx";
import Excel from "exceljs";

(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto(
    "https://admision.unmsm.edu.pe/simple/Resultados_1516_20231/A.html"
  );
  await page.waitForTimeout(500);

  const resultsLinks = new Map();

  const nodesLinks = await page
    .locator(
      "html body div.container div.row div.col-sm-12 div.table-responsive table.table.table-bordered.table-hover tbody tr td.text-center a"
    )
    .evaluateAll((nodes) => {
      return nodes.map((node) => {
        return {
          id: node.innerText,
          link: node.getAttribute("href"),
        };
      });
    });

  nodesLinks.forEach((node) => {
    resultsLinks.set(node.id, node.link);
  });

  for (var [key, value] of resultsLinks) {
    const trueLink = `https://admision.unmsm.edu.pe/simple/Resultados_1516_20231/${value}`;
    resultsLinks.set(key, trueLink);
  }

  resultsLinks.keys((id) => {
    const trueLink = `https://admision.unmsm.edu.pe/simple/Resultados_1516_20231/${resultsLinks.get(
      id
    )}`;
    resultsLinks.set(id, trueLink);
  });

  console.log("Results Links", resultsLinks);

  const career = "CIENCIA POLÍTICA";

  await page.goto(resultsLinks.get(career));

  await page.waitForTimeout(500);

  const liElementsText = await page.locator(".pagination li a").allInnerTexts();
  const finalPageText = await page
    .locator(`[data-dt-idx="${liElementsText.length - 2}"]`)
    .allInnerTexts();
  let studentCounter = 0;
  const newStudents = [];
  const maxPages = parseInt(finalPageText[0]);
  let pageCounter = 0;

  const setupNewStudentsList = async () => {
    for (let index = 0; index < maxPages; index++) {
      newStudents.push({ id: index, name: "name", score: "score" });
    }
  };

  const readDataFromTable = async () => {
    let row = await page.locator(".table-success").allInnerTexts();
    for (let data of row) {
      studentCounter++;
      let rowData = data.split("\t");
      const cachimbo = {
        id: studentCounter,
        name: rowData[1],
        score: parseFloat(rowData[3]),
      };
      newStudents.push(cachimbo);
    }
  };

  const createXLSXFile = (data, fileName) => {
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.columns = Object.keys(data[0]).map((key) => {
      return { header: key, key: key };
    });
    sheet.addRows(data);
    return workbook.xlsx.writeFile(fileName);
  };

  const clickNext = async () => {
    const idxNextButton = liElementsText.length - 1;
    await page.locator(`[data-dt-idx="${idxNextButton}"]`).click();
  };

  await setupNewStudentsList();
  console.log("Loading...(0%)");
  try {
    await async
      .eachSeries(newStudents, async (newStudents) => {
        pageCounter++;
        await readDataFromTable();

        await page.screenshot({
          path: `./screenshots/${newStudents.id}.png`,
          fullPage: true,
        });
        console.log(
          `Loading...(${Math.floor((pageCounter * 100) / maxPages)}%)`
        );
        if (pageCounter !== maxPages) {
          await clickNext();
        }
      })
      .then(async () => {
        const finalStudents = newStudents.slice(maxPages, newStudents.length);

        finalStudents.sort((a, b) => (a.score > b.score ? -1 : 1));

        console.log("Convertir JSON array to an excel file..");

        const filename = `./files/${career.split(" ").join("_")}.xlsx`;

        createXLSXFile(finalStudents, filename);
      });
  } catch (error) {
    console.log(error);
  } finally {
    await page.waitForTimeout(1000);
    await browser.close();
  }
})();
