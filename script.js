const AdmZip = require("adm-zip");
const path = require("path");
const fs = require("fs");
const { promisify } = require("util");
const exec = promisify(require("child_process").exec);

async function replaceTextAndSaveToPdf(
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

  const outputPath = path.join(__dirname, outputPdfName);

  // Use LibreOffice to convert the docx file to pdf
  const { stderr } = await exec(
    `soffice --convert-to pdf --outdir "${__dirname}" "${outputDocxName}"`
  );
  if (stderr) {
    console.error(`Error converting file: ${stderr}`);
    return;
  }

  // Rename the output pdf file
  fs.renameSync(path.join(__dirname, `${outputDocxName}.pdf`), outputPath);
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
