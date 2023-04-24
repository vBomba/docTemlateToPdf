const AdmZip = require("adm-zip");
const libre = require("libreoffice-convert");
const path = require("path");
const fs = require("fs");

function replaceTextAndSaveToPdf(entryObj, entryFileName, outputPdfName) {
  const zip = new AdmZip(entryFileName);
  const xmlContent = zip.readAsText("word/document.xml");
  const regEx = `(${Object.keys(entryObj)
    .map((v) => `{${v}}`)
    .join("|")})`;
  const modifiedContent = xmlContent.replace(
    new RegExp(regEx, "g"),
    function (c) {
      for (const [key, value] of Object.entries(entryObj)) {
        if (c === `{${key}}`) {
          if (typeof value === "string") {
            return value;
          }
        }
      }
    }
  );

  const tmpDocxPath = path.join(
    __dirname,
    `${Math.random().toString(36).substring(2, 15)}.docx`
  );
  zip.updateFile("word/document.xml", Buffer.from(modifiedContent));
  zip.writeZip(tmpDocxPath);
  const outputPath = path.join(__dirname, outputPdfName);
  const docxPdf = path.join(
    __dirname,
    `${Math.random().toString(36).substring(2, 15)}.pdf`
  );
  const enterPath = fs.readFileSync(tmpDocxPath);
  libre.convert(enterPath, ".pdf", undefined, (err, done) => {
    if (err) {
      console.log(`Error converting file: ${err}`);
    }
    fs.writeFileSync(docxPdf, done);
    fs.renameSync(docxPdf, outputPath);
    fs.unlinkSync(tmpDocxPath);
  });
}

replaceTextAndSaveToPdf(
  {
    x: "test1",
    y: "test2",
  },
  "document.docx",
  "pdf_mod.pdf"
);
