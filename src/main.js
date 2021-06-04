importScripts('../inkapi.js');

// for Docx to editor
const mammoth = require("mammoth");

// for editor content to Docx
const htmlToDocx = require("html-to-docx-buffer");

INKAPI.ready(() => {
  const UI = INKAPI.ui;

  //creating menu items for import and export
  UI.menu.addMenuItem(exportDocxHandler, "File", "Export", "as Docx");
  UI.menu.addMenuItem(importDocxHandler, "File", "Import", "from Docx");

})

// handling export menu item click
async function exportDocxHandler() {

  const Editor = INKAPI.editor;
  const IO = INKAPI.io;

  const htmlString = await Editor.getHTML(); //retrieve editor content in docx format.
  const converted = await htmlToDocx(htmlString, null, {
    table: {
      row: {
        canSplit: true,
      }
    },
    footer: true,
    pageNumber: true,
  });

  const bufferContent = await converted.arrayBuffer();

  IO.saveFile(bufferContent, 'docx');  //open save dialog with only docx file extension

}

// handling import menu item click
function importDocxHandler() {
  INKAPI.io.openFile(openFileHandler, { ext: "docx", allowMultipleFiles: false });
}

// handling file open on import
async function openFileHandler(res) {
  mammoth.convertToHtml({ arrayBuffer: res[0]?.data })
    .then(result => {
      INKAPI.editor.loadHTML(result.value);
    })
}