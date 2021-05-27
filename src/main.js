importScripts('../inkapi.js');
importScripts('./lib/html-docx.js');
importScripts('./lib/jszip.js');
importScripts('./lib/docx4js.js');

const nodesParser = {};

const docx4js = require('docx4js');

INKAPI.ready(() => {
  const UI = INKAPI.ui;

  UI.menu.addMenuItem(exportDocxHandler, "File", "Export", "as Docx");
  UI.menu.addMenuItem(importDocxHandler, "File", "Import", "from Docx");

  initNodeParser();
})


async function exportDocxHandler() {

  //do something on menu click
  const Editor = INKAPI.editor;
  const IO = INKAPI.io;

  const htmlString = await Editor.getHTML(); //retrieve editor content in docx format.
  const converted = await htmlDocx.asBlob(htmlString).arrayBuffer();

  IO.saveFile(converted, 'docx');  //open save dialog with only docx file extension

}

//using docx4js
function importDocxHandler() {
  INKAPI.io.openFile(openFileHandler, { ext: "docx", allowMultipleFiles: false });
}

async function openFileHandler(res) {
  let html = '';
  let doc = {
    word: {
      _refs: {}
    }
  };


  const jszip = new JSZip();
  doc.zip = jszip;
  const zip = await jszip.loadAsync(res[0].data);

  // will be called, even if content is corrupted


  //docx document with embedded html
  //if a "word/afchunk.mht" is present, just extract the HTML and render it
  if (zip.files["word/afchunk.mht"]) {
    const mht = await zip.files["word/afchunk.mht"].async('text');
    if (mht) {
      const matches = mht.match(/<body[^>]*>([^<]*(?:(?!<\/?body)<[^<]*)*)<\/body\s*>/i);
      if (matches)
        html = matches[1];
      INKAPI.editor.loadHTML(html);
      return;
    }
  }

  //we need the _refs sections in order to parse media files
  doc.word._refs = await zip.files["word/_rels/document.xml.rels"].async('text');

  const docx = await docx4js.load(res[0].data);
  doc.docx = docx;
  doc.medias = await preloadMedia(doc);


  //get docx body node
  const rootNode = doc.docx.officeDocument.content("w\\:body").get(0);

  //parse all children
  for (let child of rootNode.children) {
    if (nodesParser[child.name]) html += nodesParser[child.name](child, doc);
    else console.log('unparsed', child)
  }

  doc.html = html;
  console.log(html);

  INKAPI.editor.loadHTML(html);
}

//simple docx parsing code below ====================

async function preloadMedia(doc) {
  const medias = {};
  const blips = doc.docx.officeDocument.content("a\\:blip");
  for (let i = 0; i < blips.length; i++) {
    const blip = blips.get(i);
    if (blip.attribs && blip.attribs['r:embed']) {
      const mediaId = blip.attribs['r:embed'];
      const regex = new RegExp(`<Relationship[^>]+Id="${mediaId}"[^>]+Target="([^"]*)"[^>]*>`, 'gm')
      var matches = regex.exec(doc.word._refs);
      const mFilename = matches[1];
      const base64 = await doc.zip.files["word/" + mFilename].async('base64');
      medias[mediaId] = `data:image/jpeg;base64,${base64}`;
    }
  }
  return medias;
}

function initNodeParser() {
  let openTag = [];
  nodesParser['w:p'] = (node, doc) => {
    let res = ''

    for (let child of node.children) {
      if (nodesParser[child.name]) {

        res += nodesParser[child.name](child, doc);
      }
      else console.log('unparsed > ', child);
    }


    while (openTag.length > 0) {
      var tag = openTag.pop();
      res = `<${tag}>${res} </${tag}>`;
    }

    //console.log(res);
    return res;
  }

  nodesParser['w:pPr'] = (node, doc) => {
    let res = '';
    for (let child of node.children) {

      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }


    return res;
  }

  nodesParser['w:pStyle'] = (node, doc) => {


    const style = node.attribs ? node.attribs['w:val'] : undefined;
    let tag = 'p';
    let type = 'TXT';
    if (style && style == 'Title') {
      tag = 'h1';
      type = 'H';
    }
    else if (style && style.indexOf('Heading') == 0) {
      tag = 'h' + (parseInt('Heading1'.replace('Heading', '')) + 1);
      type = 'H';
    }

    let res = '';
    for (let child of node.children) {

      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    openTag.push(tag);
    return '';
  }

  nodesParser['w:r'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['w:t'] = (node, doc) => {
    let res = '';
    if (node.children && node.children[0] && node.children[0].type == 'text') {
      if (openTag.length == 0) {
        openTag.push('p');
      }
      res += node.children[0].data;
    }

    return res;
  }



  nodesParser['w:tbl'] = (node, doc) => {
    let res = '';


    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    //res = `<table>${res}</table>`; //tables are not supported in INK
    res = `${res}`;

    return res;
  }

  nodesParser['w:tr'] = (node, doc) => {
    let res = '';
    //openTag.push('tr');

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    //res = `<tr>${res}</tr>`;//tables are not supported in INK
    res = `<p>${res}</p>`;

    return res;
  }

  nodesParser['w:tc'] = (node, doc) => {
    let res = '';
    //openTag.push('td');
    openTag.push('span');//tables are not supported in INK

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    res = `  ${res}  `;
    return res;
  }



  nodesParser['w:drawing'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['wp:inline'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['a:graphic'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['a:graphicData'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['pic:pic'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['pic:blipFill'] = (node, doc) => {
    let res = '';

    for (let child of node.children) {
      if (nodesParser[child.name]) res += nodesParser[child.name](child, doc);
      else console.log('unparsed > ', child);
    }

    return res;
  }

  nodesParser['a:blip'] = (node, doc) => {
    let res = '';

    if (node.attribs && node.attribs['r:embed']) {
      const embed = node.attribs['r:embed'];

      res = `<img src="${doc.medias[embed]}"/>`;
    }

    return res;
  }
}
