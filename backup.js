import * as fs from "fs";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  patchDocument,
  PatchType,
} from "docx";
import ExcelJS from "exceljs";

console.log("Jasdf");

// Documents contain sections, you can have multiple sections per document, go here to learn more about sections
// This simple example will only contain one section
// const doc = new Document({
//     sections: [
//         {
//             properties: {},
//             children: [
//                 new Paragraph({
//                     children: [
//                         new TextRun("Hello World"),
//                         new TextRun({
//                             text: "Foo Bar",
//                             bold: true,
//                         }),
//                         new TextRun({
//                             text: "\tGithub is the best",
//                             bold: true,
//                         }),
//                     ],
//                 }),
//             ],
//         },
//     ],
// });

// // Used to export the file into a .docx file
// Packer.toBuffer(doc).then((buffer) => {
//     fs.writeFileSync("My Document.docx", buffer);
// });

// Done! A file called 'My Document.docx' will be in your file system.

// Lade die Excel-Datei
const workbook = new ExcelJS.Workbook();
const filePath = "Test.xlsx"; // Passe den Pfad an

const zeile = [{ schemaNr: "", function: "", typ: "" }];

//let cellValues;

workbook.xlsx
  .readFile("Test.xlsx")
  .then(() => {
    // Wähle das gewünschte Arbeitsblatt aus
    const worksheet = workbook.getWorksheet("Tabelle1");

    //Loop alle Rows
    const range = [];

    for (let row = 1; row <= 31; row++) {
      zeile.push({
        schemaNr: worksheet.getCell(row, 1),
        function: worksheet.getCell(row, 2),
        typ: worksheet.getCell(row, 3),
      });
    }

    //Jetzt kannst du auf die Werte innerhalb der Range zugreifen

    zeile.forEach((element) => {
      console.log(
        "SchemaNr: " +
          element.schemaNr +
          " Function: " +
          element.function +
          " Typ: " +
          element.typ
      );
    });

    /*
    const lastRow = worksheet.lastRow;
    console.log("letzte Adresse:" + lastRow.values);

    const startCell = worksheet.getCell("A2");
    const endCell = worksheet.getCell("B31");

    console.log("start Adresse:" + startCell.address);
    console.log("letzte Adresse:" + endCell.address);

    const range2 = worksheet.getCells
    const range = worksheet.getCells(startCell.address, endCell.address);

    // Jetzt kannst du auf die Zellen innerhalb der Range zugreifen
    range.forEach((cell) => {
      console.log(cell.value);
    });

*/

    // Lies die Daten aus einer bestimmten Zeile (z. B. Zeile 2)
    const rowNumber = 2; // Passe die Zeilennummer an
    const row = worksheet.getRow(rowNumber);

    // Hole Werte aus der Zeile
    const cellValues = row.values;
    cellValues.shift();
    return cellValues;
    // console.log(`Werte aus Zeile ${rowNumber}:`, cellValues);
  })
  .catch((error) => {
    console.error("Fehler beim Einlesen der Excel-Datei:", error);
  })
  .then((cellValues) => {
    // Hier kannst du auf cellValues zugreifen
    console.log("Außerhalb der Funktion:", cellValues);

    patchDocument(fs.readFileSync("../Etiketten/Etiketten_Vorlage.docx"), {
      patches: {
        SchemaNr: {
          type: PatchType.PARAGRAPH,
          children: [new TextRun(cellValues[0])],
        },
        Function: {
          type: PatchType.DOCUMENT,
          children: [
            new Paragraph("Lorem ipsum paragraph"),
            new Paragraph("Another paragraph"),
          ],
        },

        Typ: {
          type: PatchType.PARAGRAPH,
          children: [new TextRun(cellValues[0])],
        },
      },
    }).then((doc) => {
      fs.writeFileSync("../Etiketten/Etiketten_Test.docx", doc);
    });
  })
  .catch((error) => {
    // Fehlerbehandlung
    console.error("Ein Fehler ist aufgetreten:", error);
  });
