importScripts('../inkapi.js');

// for Docx to editor
const mammoth = require("mammoth");

// for editor content to Docx
const htmlToDocx = require("html-to-docx");

INKAPI.ready(() => {
  const UI = INKAPI.ui;
  const IO = INKAPI.io;

  //creating menu items for import and export
  UI.menu.addMenuItem(exportDocxHandler, "File", "Export", "as Docx");
  UI.menu.addMenuItem(importDocxHandler, "File", "Import", "from Docx");

  // associating current plugin with docx extension. to trigger whenever such file is dropped over editor
  IO.associateFileType(handleAssociateFile, "docx");

})

// handling export menu item click
async function exportDocxHandler() {

  const Editor = INKAPI.editor;
  const IO = INKAPI.io;

  let htmlString = await Editor.getHTML(); //retrieve editor content in docx format.

  // removing style and width attributes from image tags in html string
  Array.from(htmlString.matchAll(/<img([\w\W]+?)\/>/g)).map(i => {
    htmlString = htmlString.replaceAll(i[0], i[0].replaceAll(/((width|style)="([\w\W]+?)")/g, ''));
  });
  //-------------------------

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
  INKAPI.editor.resolveUnsavedContent(clear => {
    if (!clear) return;
    INKAPI.io.openFile(openFileHandler, { ext: "docx", allowMultipleFiles: false });
  });
}

function handleAssociateFile(res) {
  INKAPI.editor.resolveUnsavedContent(clear => {
    if (!clear) return;
    openFileHandler(res);
  });
}

// handling file open on import
async function openFileHandler(res) {
  mammoth.convertToHtml({ arrayBuffer: res[0]?.data })
    .then(result => {
      INKAPI.editor.clearContent();
      setTimeout(() => {
        INKAPI.editor.loadHTML(result.value);
      }, 0);
    })
}