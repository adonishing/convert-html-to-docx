/* Parse document HTML to DOCX 
*  Author: Amal Amrani
*  AUthor URL: http://amrani.es
*/
import _ from 'lodash';
import JSZip from 'jszip';
import saveAs from 'file-saver';

//  PREVIOUSLY WE CONVERT ALL DOCUMENT IMAGES TO BASE64

function funcCallback() {

}


// VARIABLES FUNCION PRINCIPAL
let zip, numImg, numLink, relsDocumentXML, countList, itemsList, ultimoItemAnidado, numberingString, 
numIdString, cols, orientationRTL, tabla, rows, rowIndex, tableWidth;

function initData() {
  numImg = 1;
  numLink = 1;
  relsDocumentXML = '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId0" Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>'; // styles.xml del document
  ultimoItemAnidado = false;
  itemsList = 0;
  countList = 0;
  numberingString = '';
  numIdString = '';
  tabla = false;
  cols = 0;
  rows = [];
  rowIndex = 0;
  tableWidth = 9280;
  orientationRTL = false;
  zip = new JSZip();
}

/**
 * recursive function to convert all document images to Base64
 */

export async function convertAllUrlImagesToBase64(docu, docxName, callBack) {
  initData();

  let allImgDoc = docu.getElementsByTagName('IMG');
  let allImg = [];

  let i = 0;
  for (let j = 0; j < allImgDoc.length; j++) {

    let url = allImgDoc[j].src;
    if (url) {
      allImg.push(allImgDoc[j]);
    }
  }

  if (allImg.length) {
    recursiva_img(allImg, i, docu, docxName, callBack);

  } else {
    let myDocx = generateDocx(docu);

    saveAs(myDocx, docxName + '.docx');
    callBack();
  }
}


/**
 * function to be executed when recursiva_img finished
 */

function ConversionCompleted(docu, docxName, callBack) {
  // Call function to generat docx
  let myDocx = generateDocx(docu);
  saveAs(myDocx, docxName + '.docx');
  callBack();
}

/**
 * Function to draw document image in canvas to convert it from url to Base64
 */

function recursiva_img(allImg, i, docu, docxName, callBack) {
  let url = allImg[i].src;

  // createCanvas
  let img = new Image();
  img.crossOrigin = 'Anonymous';
  img.src = url;

  img.onerror = function () {
    console.log(img.src);
    img.src = url + '?type=downLoad';
  }

  img.onload = function () {
    let canvas = document.createElement('CANVAS');
    let ctx = canvas.getContext('2d');
    canvas.height = img.height;
    canvas.width = img.width;
    ctx.drawImage(img, 0, 0);

    let dataURL = canvas.toDataURL('image/png');

    allImg[i].src = dataURL;
    allImg[i].width = canvas.width;
    allImg[i].height = canvas.height;

    canvas = null
    i++;

    if (i < allImg.length) {
      return recursiva_img(allImg, i, docu, docxName, callBack);
    } else {
      ConversionCompleted(docu, docxName, callBack);
    }
  };
}
//// END CONVERT IMAGES TO BASE64

function checkIfInsertTag(newEle) {
  if (newEle.nodeName === 'w:Noinsert') {
    return false;
  } else {
    return true;
  }
}

function createNodeBlockquote(node, xmlDoc) {
  let newEle = xmlDoc.createElement('w:p');
  let wpPr = xmlDoc.createElement('w:pPr');

  let wstyle = xmlDoc.createElement('w:pStyle');
  wstyle.setAttribute('w:val', 'blockQuote');
  wpPr.appendChild(wstyle);

  newEle.appendChild(wpPr);

  return newEle;
}

/**
 *  función que comprueba si existe consanguinidad en la rama con una lista para insertar la rama en el padre = w:body
 *  comprobación de cada nodo antes de sacar su equivalente en el docx
 *  ul, ol, li, y cualquier otro elemento que tenga consanguinidad con una lista
 *  Se entiende que un elemento tiene consanguinidad con lista si su padre no es body y además alguno de sus parientes (hermanos o descendientes como hijos, sobrinos, nietos o hijos de sobrinos, bisnietos, etc)
 *  @param node el nodo originar a comprobar si tiene consanguinidad con una lista
 *  @return boolen true or false
 *
 */
function consanguinidadLista(node) {
  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    return false;
  }

  // SI EL ELEMENTOS ES UNO DE ESTOS DIRECTAMENTE DEVOLVER TRUE
  if (node.nodeName === 'UL' || node.nodeName === 'OL' || node.nodeName === 'LI') {
    return true;
  } else { // SI EL PADRE NO ES BODY Y TIENEN CONSANGUINIDAD = > TRUE

    let padre = node.parentNode;

    // PREGUNTA A HIJOS Y DESCENDIENTES
    for (let hijo = 0; hijo < node.childNodes.length; hijo++)
      if (consanguinidadLista(node.childNodes[hijo])) {
        return true;
      }

    // PREGUNTA A HERMANOS Y DESCENDIENTES
    for (let hermano = 0; hermano < padre.childNodes.length; hermano++) {
      // si el hermano no soy yo preguntar
      if (padre.childNodes[hermano] != node) {

        let nodHermano = padre.childNodes[hermano];
        for (let sobrino = 0; sobrino < nodHermano.childNodes.length; sobrino++)
          if (consanguinidadLista(nodHermano.childNodes[sobrino])) {
            return true;
          }
      }
    }
  }

  return false;
}

// FUNCION RECURSIVE TO CREATE DOCX XML NODE
function createNodeXML(node, xmlDoc, padreXML) {
  let newEle = parseHTMLtoDocx(node, xmlDoc, padreXML);

  // IMPORTANTE SACAR AQUI EL ANCESTRO Y NO ANTES. PARA ASEGURARNOS DE QUE SE HA CREADO YA EL ELEMNTO BODY
  let ancestro = xmlDoc.getElementsByTagName('w:body')[0];

  let insertarTag = checkIfInsertTag(newEle);
  
  if (insertarTag) {
    // IF TABLA
    if (newEle.nodeName === 'w:tbl') {
      // insertar en ancestro
      ancestro.appendChild(newEle);
      tabla = newEle;
    } else if (newEle.nodeName === 'w:tr') {
      tabla.appendChild(newEle);
    } else if (consanguinidadLista(node)) { //  SI SE TRATA DE UNA LISTA, INSERTAR EN ANCESTRO = W:BODY
      if (newEle.nodeName === 'w:r') {
        if (ancestro.lastElementChild.nodeName === 'w:p') {
          ancestro.lastElementChild.appendChild(newEle);
        } else {
          let p = xmlDoc.createElement('w:p');
          p.appendChild(newEle);
          ancestro.appendChild(p);
        }
      } else {
        ancestro.appendChild(newEle);
      }
    } else {
      if ((node.nodeName === 'SUP' || node.nodeName === 'SUB') && padreXML.nodeName === 'w:r') {
        padreXML.parentNode.appendChild(newEle);
      } else {
        padreXML.appendChild(newEle);
      }
    }
  }

  if (newEle && node.nodeName != 'math') {
    padreXML = newEle;
  }

  if (node.nodeName != 'math') {
    for (let hijo = 0; hijo < node.childNodes.length; hijo++) {
      createNodeXML(node.childNodes[hijo], xmlDoc, padreXML);
    }
  }
 
}

function createNodeTD(node, xmlDoc, cols, status = null) {
  let colspan = Math.floor(node.getAttribute('colspan')) || 1;
  let widthPercent = tableWidth / cols * colspan;

  let newEle = xmlDoc.createElement('w:tc');
  let wtcPr = xmlDoc.createElement('w:tcPr');
  let valign = xmlDoc.createElement('w:vAlign');
  valign.setAttribute('w:val', 'center');
  if (colspan > 1) {
    let gridSpan = xmlDoc.createElement('w:gridSpan');
    gridSpan.setAttribute('w:val', colspan);
    wtcPr.appendChild(gridSpan);
  }
  let wtblBorders = xmlDoc.createElement('w:tblBorders');
  let wtop = xmlDoc.createElement('w:top');
  wtop.setAttribute('w:val', 'single');
  wtop.setAttribute('w:sz', '10');
  wtop.setAttribute('w:space', '0');
  wtop.setAttribute('w:color', '000000');
  let wstart = xmlDoc.createElement('w:start');
  wstart.setAttribute('w:val', 'single');
  wstart.setAttribute('w:sz', '10');
  wstart.setAttribute('w:space', '0');
  wstart.setAttribute('w:color', '000000');
  let wbottom = xmlDoc.createElement('w:bottom');
  wbottom.setAttribute('w:val', 'single');
  wbottom.setAttribute('w:sz', '10');
  wbottom.setAttribute('w:space', '0');
  wbottom.setAttribute('w:color', '000000');
  let wend = xmlDoc.createElement('w:end');
  wend.setAttribute('w:val', 'single');
  wend.setAttribute('w:sz', '10');
  wend.setAttribute('w:space', '0');
  wend.setAttribute('w:color', '000000');
  wtblBorders.appendChild(wtop);
  wtblBorders.appendChild(wstart);
  wtblBorders.appendChild(wbottom);
  wtblBorders.appendChild(wend);
  //  color thead si existe
  let granFather = (node.parentNode).parentNode;
  if (granFather.nodeName === 'THEAD') {
    let wshd = xmlDoc.createElement('w:shd');
    wshd.setAttribute('w:val', 'clear');
    wshd.setAttribute('w:fill', 'EEEEEE');
    wtcPr.appendChild(wshd);
  }

  let wtcW = xmlDoc.createElement('w:tcW');
  wtcW.setAttribute('w:type', 'dxa');
  wtcW.setAttribute('w:w', widthPercent);

  wtcPr.appendChild(wtblBorders);
  wtcPr.appendChild(wtcW);
  wtcPr.appendChild(valign);

  if (status) {
    let vMerge = xmlDoc.createElement('w:vMerge');
    vMerge.setAttribute('w:val', status);
    wtcPr.appendChild(vMerge);
  }

  newEle.appendChild(wtcPr);

  // SI TD VACÍO
  if (node.childNodes.length == 0 && status != 'continue') {
    let nodep = xmlDoc.createElement('w:p');
    let noder = xmlDoc.createElement('w:r');

    let nodetext = xmlDoc.createElement('w:t');
    let text = xmlDoc.createTextNode(' ');
    nodetext.appendChild(text);
    noder.appendChild(nodetext);
    
    nodep.appendChild(noder);
    newEle.appendChild(nodep);
  }

  return newEle;
}

function createNodeTR(node, xmlDoc) {
  return xmlDoc.createElement('w:tr');
}

function createTableNode(node, xmlDoc) {

  let newEle = xmlDoc.createElement('w:tbl');
  let wtblPr = xmlDoc.createElement('w:tblPr');
  let wtblStyle = xmlDoc.createElement('w:tblStyle');
  wtblStyle.setAttribute('w:val', 'TableGrid');
  let wtblW = xmlDoc.createElement('w:tblW');
  wtblW.setAttribute('w:w', tableWidth);
  wtblW.setAttribute('w:type', 'dxa');
  wtblPr.appendChild(wtblStyle);
  wtblPr.appendChild(wtblW);

  let wtblBorders = xmlDoc.createElement('w:tblBorders');
  let wtop = xmlDoc.createElement('w:top');
  wtop.setAttribute('w:val', 'single');
  wtop.setAttribute('w:sz', '10');
  wtop.setAttribute('w:space', '0');
  wtop.setAttribute('w:color', '000000');
  let wstart = xmlDoc.createElement('w:start');
  wstart.setAttribute('w:val', 'single');
  wstart.setAttribute('w:sz', '10');
  wstart.setAttribute('w:space', '0');
  wstart.setAttribute('w:color', '000000');
  let wbottom = xmlDoc.createElement('w:bottom');
  wbottom.setAttribute('w:val', 'single');
  wbottom.setAttribute('w:sz', '10');
  wbottom.setAttribute('w:space', '0');
  wbottom.setAttribute('w:color', '000000');
  let wend = xmlDoc.createElement('w:end');
  wend.setAttribute('w:val', 'single');
  wend.setAttribute('w:sz', '10');
  wend.setAttribute('w:space', '0');
  wend.setAttribute('w:color', '000000');

  let windideH = xmlDoc.createElement('w:insideH');
  windideH.setAttribute('w:val', 'single');
  windideH.setAttribute('w:sz', '5');
  windideH.setAttribute('w:space', '0');
  windideH.setAttribute('w:color', '000000');
  let windideV = xmlDoc.createElement('w:insideV');
  windideV.setAttribute('w:val', 'single');
  windideV.setAttribute('w:sz', '5');
  windideV.setAttribute('w:space', '0');
  windideV.setAttribute('w:color', '000000');


  wtblBorders.appendChild(wtop);
  wtblBorders.appendChild(wstart);
  wtblBorders.appendChild(wbottom);
  wtblBorders.appendChild(wend);

  wtblBorders.appendChild(windideH);
  wtblBorders.appendChild(windideV);

  wtblPr.appendChild(wtblBorders);

  newEle.appendChild(wtblPr);

  return newEle;
}

function createHeading(node, xmlDoc, pos) {
  let newEle = null;

  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');

    let newEle2 = xmlDoc.createElement('w:pPr');

    let newEle3 = xmlDoc.createElement('w:pStyle');
    newEle3.setAttribute('w:val', 'Heading' + pos);

    newEle.appendChild(newEle2);
    newEle2.appendChild(newEle3);

    newEle = checkRTL(newEle, xmlDoc, node);
  } else {
    newEle = xmlDoc.createElement('w:r');
  }


  return newEle;
}

function checkRTL(newEle, xmlDoc, node) {

  if (node.attributes && node.attributes[0]) {

    let direction = node.attributes[0].nodeValue.replace(' ', '');
    direction = direction.replace('direction:', '');

    let res = direction.substring(0, 3);
    if (res === 'rtl') {
      orientationRTL = true;
      //AÑADIMOS AQUI RTL
      if (newEle.childNodes && newEle.childNodes[0] && newEle.childNodes[0].nodeName === 'w:pPr') {

        let elepPr = newEle.childNodes[0];
      } else {
        let elepPr = xmlDoc.createElement('w:pPr');
        let bidi = xmlDoc.createElement('w:bidi');
        bidi.setAttribute('w:val', '1');
        elepPr.appendChild(bidi);


        newEle.appendChild(elepPr);
      }
    } else {
      orientationRTL = false;
    }
  }

  return newEle;
}

function createNodeStrong(node, xmlDoc) {
  let newEle;

  let newEleR = xmlDoc.createElement('w:r');
  let bold = xmlDoc.createElement('w:rPr');
  let propertyBold = xmlDoc.createElement('w:b');
  newEleR.appendChild(bold);
  bold.appendChild(propertyBold);

  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');
    newEle.appendChild(newEleR);
  } else {
    newEle = newEleR;
  }

  return newEle;
}

function createNodeEM(node, xmlDoc) {
  let newEle;

  let newEleR = xmlDoc.createElement('w:r');
  let nodeProperty = xmlDoc.createElement('w:rPr');
  let property = xmlDoc.createElement('w:i');
  newEleR.appendChild(nodeProperty);
  nodeProperty.appendChild(property);

  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');
    newEle.appendChild(newEleR);
  } else {
    newEle = newEleR;
  }

  return newEle;
}

function createNodeBR(node, xmlDoc) {
  let newEle;

  let newEleR = xmlDoc.createElement('w:r');
  let eleBR = xmlDoc.createElement('w:br');
  newEleR.appendChild(eleBR);

  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');
    newEle.appendChild(newEleR);
  } else {
    newEle = newEleR;
  }

  return newEle;
}

function createHiperlinkNode(node, xmlDoc, numLink) {
  // SI CONTIENE UNA IMAGEN
  for (let j = 0; j < node.childNodes.length; j++) {
    if (node.childNodes[j].nodeName === 'IMG') {

      let nR = xmlDoc.createElement('w:r');
      let nT = xmlDoc.createElement('w:t');
      nT.setAttribute('xml:space', 'preserve');
      let t = xmlDoc.createTextNode('link format not available');
      nR.appendChild(nT);
      nT.appendChild(t);

      if (node.parentNode.nodeName === 'BODY') {
        let p = xmlDoc.createElement('w:p');
        p.appendChild(nR);
        return p;
      } else {
        return nR;
      }
    }
  }

  let hiperEle = xmlDoc.createElement('w:hyperlink');
  if (node.href) {  // EXTERNAL LINK
    hiperEle.setAttribute('r:id', 'link' + numLink);

  }

  let newEle;
  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');
    newEle.appendChild(hiperEle);
  } else {
    newEle = hiperEle;
  }

  return newEle;
}

function createNodeLi(node, xmlDoc, countList, itemsList) {
  let father = node.parentNode;

  let newEle = xmlDoc.createElement('w:p');
  let pr = xmlDoc.createElement('w:pPr');
  let rstyle = xmlDoc.createElement('w:pStyle');
  rstyle.setAttribute('w:val', 'ListParagraph'); //('mystyleList');   ListParagraph

  pr.appendChild(rstyle);
  newEle.appendChild(pr);

  let numpr = xmlDoc.createElement('w:numPr');
  let wilvl = xmlDoc.createElement('w:ilvl');
  wilvl.setAttribute('w:val', itemsList);
  let numId = xmlDoc.createElement('w:numId');
  numId.setAttribute('w:val', countList);

  numpr.appendChild(wilvl);
  numpr.appendChild(numId);

  pr.appendChild(numpr);

  return newEle;
}

/**
 *  function to create a xml document node paragraph or node run
 *  @param: node html node to parse
 *  @return a new document.xml node
 */
function createNodeParagraphOrRun(node, xmlDoc) {
  let tx = node.data;

  // si texto undefined. asignar caracter blanco
  if (!tx) tx = '';

  let newEle;

  // SI EL TAG NO RECONOCIBLE Y PADRE = BODY. SE TRADUCE EN UN PÁRRAFO
  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');
    newEle = checkRTL(newEle, xmlDoc, node);
  } else {
    newEle = xmlDoc.createElement('w:r');
  }

  return newEle;
}

function createNodeSupOrSub(node, xmlDoc) {

  let newEle = xmlDoc.createElement('w:r');

  let wrPr = xmlDoc.createElement('w:rPr');

  let wszCs = xmlDoc.createElement('w:szCs');
  wszCs.setAttribute('w:val', 24);

  let alignVal = node.nodeName == 'SUP' ? 'superscript' : 'subscript';
  let vertAlign = xmlDoc.createElement('w:vertAlign');
  vertAlign.setAttribute('w:val', alignVal);

  newEle.appendChild(wrPr);

  if (node.parentNode.nodeName == 'EM') {
    let property = xmlDoc.createElement('w:i');
    wrPr.appendChild(property);
  }

  wrPr.appendChild(wszCs);
  wrPr.appendChild(vertAlign);

  return newEle;
}

function createNodeUOrS(node, xmlDoc) {
  let newEle = xmlDoc.createElement('w:r');

  let wrPr = xmlDoc.createElement('w:rPr');

  if (node.nodeName === 'U') {
    let wu = xmlDoc.createElement('w:u');
    wu.setAttribute('w:val', 'single');
    wrPr.appendChild(wu);
  } else {
    let strike = xmlDoc.createElement('w:strike');
    wrPr.appendChild(strike);
  }
  
  newEle.appendChild(wrPr);

  return newEle;
}

// 特殊节点名称
let specialNodeName = ['EM', 'SUP', 'SUB', 'U', 'S'];

/**
 *  function to create text node
 *  @param node html node to convert into docx node
 *  @param xmlDoc xml document
 *  @return docx node
 */
function createMyTextNode(node, xmlDoc) {
  let tx = node.data;

  if (!tx) tx = ''; // IF TX UNDEFINED  !!!!

  if (node.parentNode.nodeName === 'TR' || node.parentNode.nodeName === 'TABLE') {
    return xmlDoc.createElement('w:Noinsert');
  } else if (node.parentNode.nodeName === 'BODY' && (!node.data.trim())) {
    return xmlDoc.createElement('w:Noinsert');
  }

  let newEleR = xmlDoc.createElement('w:r');

  if (orientationRTL) {

    let elerPr = xmlDoc.createElement('w:rPr');
    let rtl = xmlDoc.createElement('w:rtl');
    rtl.setAttribute('w:val', '1');
    elerPr.appendChild(rtl);

    newEleR.appendChild(elerPr);
  }

  if (node.parentNode.nodeName === 'A' && node.parentNode.href) {
    let nodeStyleLink = xmlDoc.createElement('w:rStyle');
    nodeStyleLink.setAttribute('w:val', 'Hyperlink');
    newEleR.appendChild(nodeStyleLink);
  }

  let nodetext = xmlDoc.createElement('w:t');
  nodetext.setAttribute('xml:space', 'preserve');
  let texto = xmlDoc.createTextNode(tx);

  if (specialNodeName.indexOf(node.parentNode.nodeName) > -1) {
    newEleR = nodetext;
  } else {
    newEleR.appendChild(nodetext);
  }

  nodetext.appendChild(texto);

  let newEle;
  if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
    newEle = xmlDoc.createElement('w:p');
    newEle.appendChild(newEleR);
  } else {
    newEle = newEleR;
  }

  return newEle;
}

function createAstractNumListDecimal(numberingString, countList) {
  numberingString = '<w:abstractNum w:abstractNumId="' + countList + `" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
	<w:multiLevelType w:val="hybridMultilevel" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" />
		<w:lvl w:ilvl="0" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" />
			<w:numFmt w:val="decimal" />
			<w:lvlText w:val="%1." />
			<w:lvlJc w:val="left" />
			<w:pPr><w:ind w:left="720" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%2." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="1440" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="2" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%3." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="2160" w:hanging="180" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="3" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%4." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="2880" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="4" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%5." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="3600" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="5" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%6." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="4320" w:hanging="180" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="6" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="decimal" /><w:lvlText w:val="%7." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5040" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="7" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerLetter" /><w:lvlText w:val="%8." /><w:lvlJc w:val="left" /><w:pPr><w:ind w:left="5760" w:hanging="360" /></w:pPr>
		</w:lvl>
		<w:lvl w:ilvl="8" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:start w:val="1" /><w:numFmt w:val="lowerRoman" /><w:lvlText w:val="%9." /><w:lvlJc w:val="right" /><w:pPr><w:ind w:left="6480" w:hanging="180" /></w:pPr>
		</w:lvl>
</w:abstractNum>` + numberingString;

  return numberingString;
}

function createAstractNumListBullet(numberingString, countList) {
  numberingString = '<w:abstractNum xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:abstractNumId="' + countList + `">
	<w:multiLevelType xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="hybridMultilevel"/>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="0">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val=""/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="720" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="1">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val="o"/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="1440" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="2">
		<w:start w:val="1"/>
		<w:numFmt w:val="bullet"/>
		<w:lvlText w:val=""/>
		<w:lvlJc w:val="left"/>
		<w:pPr>
		<w:ind w:left="2160" w:hanging="360"/>
		</w:pPr>
		<w:rPr>
		<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
		</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="3">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="2880" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="4">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val="o"/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="3600" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="5">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="4320" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="6">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="5040" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Symbol" w:hAnsi="Symbol"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="7">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val="o"/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="5760" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Courier New" w:hAnsi="Courier New"/>
	</w:rPr>
	</w:lvl>
	<w:lvl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ilvl="8">
	<w:start w:val="1"/>
	<w:numFmt w:val="bullet"/>
	<w:lvlText w:val=""/>
	<w:lvlJc w:val="left"/>
	<w:pPr>
	<w:ind w:left="6480" w:hanging="360"/>
	</w:pPr>
	<w:rPr>
	<w:rFonts w:hint="default" w:ascii="Wingdings" w:hAnsi="Wingdings"/>
	</w:rPr>
	</w:lvl>
</w:abstractNum>` + numberingString;

  return numberingString;
}

function createNumIdList(numIdString, countList) {
  numIdString = '<w:num w:numId="' + countList + '"><w:abstractNumId w:val="' + countList + '"/></w:num>' + numIdString;

  return numIdString;
}

/**
 * function para redimensionar la imagen si es necesario y excede el tamaño de página
 * devuelve las dimensiones de la imagen en EMUs para pintarla posteriormente en el docx
 *
 *  hacer la conversión corrcetamente teniendo en cuanta DPI
 *
 *  1 inch (pulgada)   = 914400 EMUs
 *
 *  1cm = 360000 EMUs
 *  Teniendo en cuenta que la página mide 21 x 29.7 cm
 *
 *
 */
function escalarIMG(imgwidth, imgheight) {
  let dimensionIMG = {'width': 0, 'height': 0}

  // tomando  96px por pulgada . si se quiere la imagen para impresion en papel probar dividir entre 300 x pulgada
  let width_inch = imgwidth / 96;
  let height_inch = imgheight / 96;

  let width_emu = width_inch * 914400;
  let height_emu = height_inch * 914400;

  let pgSzW = (16 * 360000); // // ancho de página en EMUs
  let pgSzH = (24.7 * 360000); // // alto de página en EMUs

  if (width_emu > pgSzW) {
    let originalW = width_emu;
    width_emu = pgSzW;
    height_emu = Math.floor(width_emu * height_emu / originalW);
  }

  if (height_emu > pgSzH) {
    let originalH = height_emu;
    height_emu = pgSzH;
    width_emu = Math.floor(height_emu * width_emu / originalH);
  }

  dimensionIMG.width = Math.floor(width_emu);
  dimensionIMG.height = Math.floor(height_emu);

  return dimensionIMG;
}

/**
 *  Función que devuelve nodo vacío porque no reconoce el formato de imagen
 *  @param xmlDoc xml docx document
 *  @return node element with text = Not available image format
 */
function nodeVoid(xmlDoc) {
  let newEle = xmlDoc.createElement('w:r');
  let nodetext = xmlDoc.createElement('w:t');
  nodetext.setAttribute('xml:space', 'preserve');
  let texto = xmlDoc.createTextNode('IMAGE FORMAT NOT AVAILABLE!');

  newEle.appendChild(nodetext);
  nodetext.appendChild(texto);

  return newEle;
}

let xsl = null;

function createNodeMath(node, xmlDoc) {
  // 为避免更改原node的parentNode， 此处克隆一个新的
  let thisNode = node.cloneNode(true);

  let newEle = xmlDoc.createElement('w:r');
  newEle.appendChild(thisNode);

  if (!xsl) {
    let xhttp = new XMLHttpRequest();

    xhttp.open('GET', '/mml2omml.xsl', false);
    xhttp.send('');

    xsl = xhttp.responseXML;
  }
  let xsltProcessor = new XSLTProcessor();
  xsltProcessor.importStylesheet(xsl);

  let resultDocument = xsltProcessor.transformToDocument(newEle);
  let newNode = resultDocument.childNodes[0];

  return newNode;
}

/**
 *  función que crea el nodo picture a insertar en el document.xml de docx
 *  @param: node  el nodo IMG origen a parsear
 *
 */
function createDrawingNodeIMG(node, dataImg, xmlDoc, numImg) {
  let format = '.png';
  let nameFile = 'image' + numImg + format;

  let relashionImg = 'rId' + numImg;


  // CREO NODO IMAGEN EN EL document PARA PODER VER SU ALTO Y ANCHO
  let img = document.createElement('img');
  img.src = dataImg;
  img.width = node.width;
  img.height = node.height;

  // ESCALAMOS IMAGEN
  let dimensionImg = escalarIMG(img.width, img.height);

  // CREAR NODO DRAWING PICTURE EN DOCX

  let newEleImage = xmlDoc.createElement('w:r');

  let drawEle = xmlDoc.createElement('w:drawing');
  let wpinline = xmlDoc.createElement('wp:inline');
  wpinline.setAttribute('distR', '0');
  wpinline.setAttribute('distL', '0');
  wpinline.setAttribute('distB', '0');
  wpinline.setAttribute('distT', '0');
  let wpextent = xmlDoc.createElement('wp:extent');
  wpextent.setAttribute('cy', dimensionImg.height);
  wpextent.setAttribute('cx', dimensionImg.width);
  let wpeffectExtent = xmlDoc.createElement('wp:effectExtent');
  wpeffectExtent.setAttribute('b', '0');
  wpeffectExtent.setAttribute('r', '0');
  wpeffectExtent.setAttribute('t', '0');
  wpeffectExtent.setAttribute('l', '0');
  let wpdocPr = xmlDoc.createElement('wp:docPr');
  wpdocPr.setAttribute('name', nameFile);
  wpdocPr.setAttribute('id', numImg);

  let wpcNvGraphicFramePr = xmlDoc.createElement('wp:cNvGraphicFramePr');
  let childcNvGraphicFramePr = xmlDoc.createElement('a:graphicFrameLocks');
  childcNvGraphicFramePr.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
  childcNvGraphicFramePr.setAttribute('noChangeAspect', '1');

  wpcNvGraphicFramePr.appendChild(childcNvGraphicFramePr);

  let agraphic = xmlDoc.createElement('a:graphic');
  agraphic.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
  let agraphicdata = xmlDoc.createElement('a:graphicData');
  agraphicdata.setAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
  let pic = xmlDoc.createElement('pic:pic');
  pic.setAttribute('xmlns:pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture');
  let picnvPicPr = xmlDoc.createElement('pic:nvPicPr');
  let piccNvPr = xmlDoc.createElement('pic:cNvPr');
  piccNvPr.setAttribute('name', nameFile);
  piccNvPr.setAttribute('id', numImg);
  let piccNvPicPr = xmlDoc.createElement('pic:cNvPicPr');
  picnvPicPr.appendChild(piccNvPr);
  picnvPicPr.appendChild(piccNvPicPr);

  let picblipFill = xmlDoc.createElement('pic:blipFill');
  let ablip = xmlDoc.createElement('a:blip');
  ablip.setAttribute('cstate', 'print');
  ablip.setAttribute('r:embed', relashionImg);
  let astretch = xmlDoc.createElement('a:stretch');
  let afillRect = xmlDoc.createElement('a:fillRect');
  astretch.appendChild(afillRect);

  picblipFill.appendChild(ablip);
  picblipFill.appendChild(astretch);

  let picspPr = xmlDoc.createElement('pic:spPr');
  let axfrm = xmlDoc.createElement('a:xfrm');
  let aoff = xmlDoc.createElement('a:off');
  aoff.setAttribute('y', '0');
  aoff.setAttribute('x', '0');
  let aext = xmlDoc.createElement('a:ext');
  aext.setAttribute('cy', dimensionImg.height);
  aext.setAttribute('cx', dimensionImg.width);
  axfrm.appendChild(aoff);
  axfrm.appendChild(aext);

  let aprstGeom = xmlDoc.createElement('a:prstGeom');
  aprstGeom.setAttribute('prst', 'rect');
  let aavLst = xmlDoc.createElement('a:avLst');
  aprstGeom.appendChild(aavLst);

  picspPr.appendChild(axfrm);
  picspPr.appendChild(aprstGeom);

  pic.appendChild(picnvPicPr);
  pic.appendChild(picblipFill);
  pic.appendChild(picspPr);

  agraphicdata.appendChild(pic);

  agraphic.appendChild(agraphicdata);

  wpinline.appendChild(wpextent);
  wpinline.appendChild(wpeffectExtent);
  wpinline.appendChild(wpdocPr);
  wpinline.appendChild(wpcNvGraphicFramePr);
  wpinline.appendChild(agraphic);

  drawEle.appendChild(wpinline);

  newEleImage.appendChild(drawEle);

  return newEleImage;
}

/**
 * function that check if string starts with prefix
 */
function stringStartsWith(string, prefix) {
  return string.slice(0, prefix.length) == prefix;
}

/**
 *  función que forma el docx. Genera el zip y lo guarda
 *  @param: XmlDocumentDocx  el document.xml formado
 *
 */
function createDocx(XmlDocumentDocx, zip) {

  let relationShips = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  let contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default ContentType="image/jpeg" Extension="jpg"/>' +
    '<Default ContentType="image/png" Extension="png"/>' +
    '<Default ContentType="image/gif" Extension="gif"/>' +
    '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" PartName="/word/styles.xml"/>' +
    '<Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" PartName="/word/numbering.xml"/>' +
    '</Types>';

  let head_docx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14">';

  let footer = '</w:document>';

  // CARGAR ESTYLOS AQUI

  let estilos = `
		  <w:styles xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15 wp14">
		<w:style w:type="paragraph" w:styleId="Heading1">
			<w:name w:val="Heading 1"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="48"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading2">
			<w:name w:val="Heading 2"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="38"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading3">
			<w:name w:val="Heading 3"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="35"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading4">
			<w:name w:val="Heading 4"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="30"/>
			</w:rPr>
			</w:style>	
		<w:style w:type="paragraph" w:styleId="Heading5">
			<w:name w:val="Heading 5"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="28"/>
			</w:rPr>
			</w:style>
		<w:style w:type="paragraph" w:styleId="Heading6">
			<w:name w:val="Heading 6"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>
				<w:keepNext/>
				<w:keepLines/>
				<w:spacing w:before="480" w:after="0"/>
				<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
				
				<w:b/>
				<w:color w:val="000000"/>
				<w:sz w:val="25"/>
			</w:rPr>
			</w:style>	
			<w:style w:type="paragraph" w:styleId="blockQuote">
			<w:name w:val="blockQuote"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>			
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:pPr>			
				<w:spacing w:before="360" w:after="360"/>
				<w:ind w:left="1080" w:right="1080" />
			</w:pPr>
			<w:rPr>				
				<w:color w:val="#222222"/>
				<w:sz w:val="30"/>
			</w:rPr>			
			</w:style>			
		<w:style xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="character" w:styleId="Hyperlink" mc:Ignorable="w14">
			<w:name xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="Hyperlink"/>
			<w:basedOn xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="DefaultParagraphFont"/>
			<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="99"/>
			<w:unhideWhenUsed xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
			<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:color w:val="0563C1" />
			<w:b/>
			<w:u w:val="single"/>
			</w:rPr>
			</w:style>
		<w:style mc:Ignorable="w14" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
  w:styleId="ListParagraph" w:type="paragraph">
			<w:name xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="List Paragraph"/>
			<w:basedOn xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="Normal"/>
			<w:uiPriority xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="34"/>
			<w:qFormat xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
			<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:left="720"/>
			<w:contextualSpacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
			</w:pPr>
			</w:style>

		</w:styles>
		`;

  let cabeceraNumbering = '<?xml version="1.0" encoding="utf-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
  let footerNumbering = '</w:numbering>';

  numberingString = cabeceraNumbering + numberingString + numIdString + footerNumbering;

  // create a relationShips file
  zip.file('_rels/.rels', relationShips);
  // create a Content_tuypes.xml file
  zip.file('[Content_Types].xml', contentTypes);
  // document.xml
  XmlDocumentDocx = head_docx + XmlDocumentDocx + footer;

  zip.file('word/document.xml', XmlDocumentDocx);
  zip.file('word/styles.xml', estilos);
  zip.file('word/numbering.xml', numberingString);

  let content = zip.generate({type: 'blob'});

  return content;
}

//  FUNCION parseHTMLtoDocx node
function parseHTMLtoDocx(node, xmlDoc, padreXML) {
  let encabezados = ['H1', 'H2', 'H3', 'H4', 'H5', 'H6'];
  let newEle;

  // CREAR NODO BODY
  if (node.nodeName === 'BODY') {
    newEle = xmlDoc.createElement('w:body');
  } else if (encabezados.indexOf(node.nodeName) >= 0) { // CREAR NODO ENCABEZADO
    let pos = encabezados.indexOf(node.nodeName) + 1;
    newEle = createHeading(node, xmlDoc, pos);
  } else if (node.nodeName === 'P') {  // CREAR NODO P.
    newEle = createNodeParagraphOrRun(node, xmlDoc);
  } else if (node.nodeName === 'STRONG') { // CREAR NODO STRONG
    newEle = createNodeStrong(node, xmlDoc);
  } else if (node.nodeName === 'EM') { // CREAR NODO EM
    newEle = createNodeEM(node, xmlDoc);
  } else if (node.nodeName === 'BR') {
    newEle = createNodeBR(node, xmlDoc);
  } else if (node.nodeName === 'A') { // CREAR NODO HIPERLINK
    if (node.href) {  // EXTERNAL LINK
      let idLink = 'link' + numLink;
      relsDocumentXML = relsDocumentXML + '<Relationship Id="' + idLink + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + node.href + '" TargetMode="External"/>';

      newEle = createHiperlinkNode(node, xmlDoc, numLink);

      numLink++;
    } else {
      newEle = createNodeParagraphOrRun(node, xmlDoc);
    }
  } else if (node.nodeName === 'UL' || node.nodeName === 'OL') { // CREAR NODO LISTA . estos tags directamente se insertan en el body independientemente, en parrafos diferenciados
    // SI LISTA ANIDADA	. SIGO EN LA MISMA LISTA ITEM SE INCREMENTA EN UNO PARA EL SIGUIENTE NIVEL DE LA LISTA

    if (node.parentNode.nodeName != 'LI') {
      countList++;
    }

    if (node.nodeName === 'UL') {
      numberingString = createAstractNumListBullet(numberingString, countList);
    } else if (node.nodeName === 'OL') {
      numberingString = createAstractNumListDecimal(numberingString, countList);
    }

    numIdString = createNumIdList(numIdString, countList);

    newEle = xmlDoc.createElement('w:Noinsert');

  } else if (node.nodeName === 'LI') {
    let levelList = 0;
    let granFather = node.parentNode.parentNode;

    // CUANTOS ABUELOS TENGO
    if (granFather) {
      let abu = true;

      while (abu && granFather.nodeName === 'LI') {
        levelList++;

        if (granFather.parentNode.parentNode) {
          abu = true;
          granFather = granFather.parentNode.parentNode;
        } else abu = false;
      }
    }

    newEle = createNodeLi(node, xmlDoc, countList, levelList);
  } else if (node.nodeName === '#text') { // CREAR NODO TEXT.
    newEle = createMyTextNode(node, xmlDoc);
  } else if (node.nodeName === 'IMG') { // CREAR NODO IMAGEN
    let format = '.png';
    let nameFile = 'image' + numImg + format;
    let relashionImg = 'rId' + numImg;
    let imgEle;

    let dataImg = node.attributes[0].value;

    if (!stringStartsWith(dataImg, 'data:')) {
      imgEle = nodeVoid(xmlDoc);
    } else {
      let srcImg = dataImg.replace(/^data:image\/.+;base64,/, ' ');
      //GUARDO LA IMAGEN EN MEDIA
      zip.file('word/media/' + nameFile, srcImg, {base64: true});

      // CREO LA RELATIONSHIP DE LA IMAGEN CREADA
      relsDocumentXML = relsDocumentXML + '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + nameFile + '" Id="' + relashionImg + '" />';

      imgEle = createDrawingNodeIMG(node, dataImg, xmlDoc, numImg);

      numImg++;
    }

    if (node.parentNode.nodeName === 'BODY' || node.parentNode.nodeName === 'TD') {
      newEle = xmlDoc.createElement('w:p');
      newEle.appendChild(imgEle);
    } else {
      newEle = imgEle;
    }
  } else if (node.nodeName === 'TABLE') { // IF TABLE
    newEle = createTableNode(node, xmlDoc);
  } else if (node.nodeName === 'THEAD') { // no se inserta en el xml
    newEle = xmlDoc.createElement('w:Noinsert');
  } else if (node.nodeName === 'TBODY') {// no se inserta en el xml
    newEle = xmlDoc.createElement('w:Noinsert');
  } else if (node.nodeName === 'TR') {
    newEle = createNodeTR(node, xmlDoc);
    // calculo num columnas
    let maxCol = cols ? cols : 0; // The number of columns in the first row is the largest column in the table
    let col = 0;
    for (let c = 0; c < node.childNodes.length; c++) {
      if (node.childNodes[c].nodeName === 'TD') {
        let colspan = Math.floor(node.childNodes[c].getAttribute('colspan')) || 1;
        if (!maxCol) {
          cols += colspan;
        }
        node.childNodes[c].setAttribute('key', col);
        let rowspan = Math.floor(node.childNodes[c].getAttribute('rowspan')) || 1;

        rowIndex = 0

        if (rowspan <= 1) {
          col++;
          continue;
        }

        for (let i = 0; i < colspan; i++) {
          rows = _.unionBy([{index: col, status: 'restart', value: rowspan}], rows, 'index');
          col++;
        }
      }
    }

    // SI FILA VACÍA.CONTROLAR!
    if (cols == 0) {
      let eleTD = document.createElement('TD');
      node.appendChild(eleTD);
    }
  } else if (node.nodeName === 'TD') {
    let tdIndex = Math.floor(node.getAttribute('key'));
    let noRenderRows = _.filter(rows, (row) => {
      return row.index <= tdIndex && row.value > 0 && row.index >= rowIndex;
    });
    if (noRenderRows && noRenderRows.length) {
      for (let row of noRenderRows) {
        if (row.status == 'restart') {
          continue;
        }
        newEle = createNodeTD(node, xmlDoc, cols, row.status);
        row.value--;
        row.status = 'continue';
        padreXML.appendChild(newEle);
        rowIndex++;
      }
    }

    let thisRow = _.find(rows, {index: tdIndex, status: 'restart'});
    if (thisRow && thisRow.status) {
      newEle = createNodeTD(node, xmlDoc, cols, thisRow.status);
      thisRow.status = 'continue';
      thisRow.value--;
      rowIndex++;
    } else {
      newEle = createNodeTD(node, xmlDoc, cols);
    }
  } else if (node.nodeName === 'BLOCKQUOTE') { // BLOQUOTE
    newEle = createNodeBlockquote(node, xmlDoc);
  } else if (node.nodeName === 'math') {
    newEle = createNodeMath(node, xmlDoc);
  } else if (node.nodeName === 'SUP' || node.nodeName === 'SUB') {
    newEle = createNodeSupOrSub(node, xmlDoc);
  } else if (node.nodeName === 'U' || node.nodeName === 'S') {
    newEle = createNodeUOrS(node, xmlDoc);
  } else { // SI SE ENCUENTRA UN TAG QUE NO RECONOCE.
    newEle = createNodeParagraphOrRun(node, xmlDoc);
  }

  return newEle;
}

/**
 * function parse HTML to Docx document. Need jszip.js and FileSaver.js
 * @param: html string
 *
 */
function generateDocx(doc) {

  //   crear los nodos del xml
  let nodeParent = doc.getElementsByTagName('body');
  let xmlDoc = document.implementation.createDocument(null, 'mywordXML');

  let padreXML = xmlDoc.childNodes[0];

  createNodeXML(nodeParent[0], xmlDoc, padreXML); // primer nodo es body[0]

  relsDocumentXML = relsDocumentXML + '<Relationship Id="Rnumbering1" Target="/word/numbering.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"/>';
  // cerrar rels Document.xml.rels
  relsDocumentXML = relsDocumentXML + '</Relationships>';

  zip.file('word/_rels/document.xml.rels', relsDocumentXML);

  let xml = xmlDoc.getElementsByTagName('w:body');

  let content = createDocx(xml[0].outerHTML, zip);

  return content;
}










