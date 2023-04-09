const AdmZip = require("adm-zip");
const libre = require("libreoffice-convert");

const path = require("path");
const fs = require("fs");

function replaceTextAndSaveToPdf(
  entryObj,
  entryFileName,
  outputDocxName,
  outputPdfName
) {
  // Load the file as a zip archive
  const zip = new AdmZip(entryFileName);

  // Read the contents of the "word/document.xml" file
  const xmlContent = zip.readAsText("word/document.xml");

  const regEx = `(${Object.keys(entryObj)
    .map((v) => `{${v}}`)
    .join("|")})`;

  // Replace the text
  const modifiedContent = xmlContent.replace(
    new RegExp(regEx, "g"),
    function (c) {
      for (const [key, value] of Object.entries(entryObj)) {
        if (c === `{${key}}`) {
          return value;
        }
      }
    }
  );

  // Write the modified content back to the zip archive
  zip.updateFile("word/document.xml", Buffer.from(modifiedContent));

  // Save the modified zip archive to a new file
  zip.writeZip(outputDocxName);

  const extend = ".pdf";
  const FilePath = path.join(__dirname, outputDocxName);
  const outputPath = path.join(__dirname, outputPdfName);

  // Read file
  const enterPath = fs.readFileSync(FilePath);
  // Convert it to pdf format with undefined filter (see Libreoffice doc about filter)
  libre.convert(enterPath, extend, undefined, (err, done) => {
    if (err) {
      console.log(`Error converting file: ${err}`);
    }

    // Here in done you have pdf file which you can save or transfer in another stream
    fs.writeFileSync(outputPath, done);
  });
}

replaceTextAndSaveToPdf(
  {
    x: "test1",
    y: "test2",
  },
  "document.docx",
  "docx_mod.docx",
  "pdf_mod.pdf"
);
