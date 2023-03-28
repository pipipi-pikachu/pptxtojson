import JSZip from 'jszip'
import * as txml from 'txml/dist/txml.mjs'
import tinycolor from 'tinycolor2'

import {
  extractFileExtension,
  base64ArrayBuffer,
  eachElement,
  getTextByPathList,
  angleToDegrees,
  escapeHtml,
  getMimeType,
} from './utils'

const FACTOR = 75 / 914400

let themeContent = null
let defaultTextStyle = null

export async function parse(file) {
  const slides = []
  
  const zip = await JSZip.loadAsync(file)

  const filesInfo = await getContentTypes(zip)
  const { width, height, defaultTextStyle: _defaultTextStyle } = await getSlideSize(zip)
  themeContent = await loadTheme(zip)
  defaultTextStyle = _defaultTextStyle

  for (const filename of filesInfo.slides) {
    const singleSlide = await processSingleSlide(zip, filename)
    slides.push(singleSlide)
  }

  return {
    slides,
    size: { width, height },
  }
}

function simplifyLostLess(children, parentAttributes = {}) {
  const out = {}
  if (!children.length) return out

  if (children.length === 1 && typeof children[0] === 'string') {
    return Object.keys(parentAttributes).length ? {
      attrs: parentAttributes,
      value: children[0],
    } : children[0]
  }
  for (const child of children) {
    if (typeof child !== 'object') return
    if (child.tagName === '?xml') continue

    if (!out[child.tagName]) out[child.tagName] = []

    const kids = simplifyLostLess(child.children || [], child.attributes)
    out[child.tagName].push(kids)

    if (Object.keys(child.attributes).length) {
      kids.attrs = child.attributes
    }
  }
  for (const child in out) {
    if (out[child].length === 1) out[child] = out[child][0]
  }

  return out
}

async function readXmlFile(zip, filename) {
  const data = await zip.file(filename).async('string')
  return simplifyLostLess(txml.parse(data))
}

async function getContentTypes(zip) {
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson['Types']['Override']
  const slidesLocArray = []
  const slideLayoutsLocArray = []

  for (const item of subObj) {
    switch (item['attrs']['ContentType']) {
      case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
        slidesLocArray.push(item['attrs']['PartName'].substr(1))
        break
      case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
        slideLayoutsLocArray.push(item['attrs']['PartName'].substr(1))
        break
      default:
    }
  }
  return {
    slides: slidesLocArray,
    slideLayouts: slideLayoutsLocArray,
  }
}

async function getSlideSize(zip) {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content['p:presentation']['p:sldSz']['attrs']
  const defaultTextStyle = content['p:presentation']['p:defaultTextStyle']
  return {
    width: parseInt(sldSzAttrs['cx']) * FACTOR,
    height: parseInt(sldSzAttrs['cy']) * FACTOR,
    defaultTextStyle,
  }
}

async function loadTheme(zip) {
  const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  const relationshipArray = preResContent['Relationships']['Relationship']
  let themeURI

  if (relationshipArray.constructor === Array) {
    for (const relationshipItem of relationshipArray) {
      if (relationshipItem['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
        themeURI = relationshipItem['attrs']['Target']
        break
      }
    }
  } 
  else if (relationshipArray['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
    themeURI = relationshipArray['attrs']['Target']
  }

  if (!themeURI) throw Error(`Can't open theme file.`)

  return await readXmlFile(zip, 'ppt/' + themeURI)
}

async function processSingleSlide(zip, sldFileName) {
  const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels'
  const resContent = await readXmlFile(zip, resName)
  let relationshipArray = resContent['Relationships']['Relationship']
  let layoutFilename = ''
  let diagramFilename = ''
  const slideResObj = {}

  if (relationshipArray.constructor === Array) {
    for (const relationshipArrayItem of relationshipArray) {
      switch (relationshipArrayItem['attrs']['Type']) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
          layoutFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          break
        case 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
          diagramFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          slideResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          }
          break
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
        default:
          slideResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
      }
    }
  } 
  else layoutFilename = relationshipArray['attrs']['Target'].replace('../', 'ppt/')

  const slideLayoutContent = await readXmlFile(zip, layoutFilename)
  const slideLayoutTables = await indexNodes(slideLayoutContent)

  const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
  const slideLayoutResContent = await readXmlFile(zip, slideLayoutResFilename)
  relationshipArray = slideLayoutResContent['Relationships']['Relationship']

  let masterFilename = ''
  const layoutResObj = {}
  if (relationshipArray.constructor === Array) {
    for (const relationshipArrayItem of relationshipArray) {
      switch (relationshipArrayItem['attrs']['Type']) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
          masterFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          break
        default:
          layoutResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
      }
    }
  } 
  else masterFilename = relationshipArray['attrs']['Target'].replace('../', 'ppt/')

  const slideMasterContent = await readXmlFile(zip, masterFilename)
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
  const slideMasterTables = indexNodes(slideMasterContent)

  const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
  const slideMasterResContent = await readXmlFile(zip, slideMasterResFilename)
  relationshipArray = slideMasterResContent['Relationships']['Relationship']
  const masterResObj = {}
  if (relationshipArray.constructor === Array) {
    for (const relationshipArrayItem of relationshipArray) {
      switch (relationshipArrayItem['attrs']['Type']) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
          break
        default:
          masterResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
      }
    }
  }

  const diagramResObj = {}
  let digramFileContent = {}
  if (diagramFilename) {
    const diagName = diagramFilename.split('/').pop()
    const diagramResFileName = diagramFilename.replace(diagName, '_rels/' + diagName) + '.rels'
    digramFileContent = await readXmlFile(zip, diagramFilename)
    if (digramFileContent && digramFileContent && digramFileContent !== '') {
      let digramFileContentObjToStr = JSON.stringify(digramFileContent)
      digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, 'p:')
      digramFileContent = JSON.parse(digramFileContentObjToStr)
    }

    const digramResContent = await readXmlFile(zip, diagramResFileName)
    if (digramResContent) {
      relationshipArray = digramResContent['Relationships']['Relationship']
      if (relationshipArray.constructor === Array) {
        for (const relationshipArrayItem of relationshipArray) {
          diagramResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
        }
      } 
      else {
        diagramResObj[relationshipArray['attrs']['Id']] = {
          type: relationshipArray['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          target: relationshipArray['attrs']['Target'].replace('../', 'ppt/'),
        }
      }
    }
  }

  const slideContent = await readXmlFile(zip, sldFileName)
  const nodes = slideContent['p:sld']['p:cSld']['p:spTree']
  const warpObj = {
    zip,
    slideLayoutContent: slideLayoutContent,
    slideLayoutTables: slideLayoutTables,
    slideMasterContent: slideMasterContent,
    slideMasterTables: slideMasterTables,
    slideContent: slideContent,
    slideResObj: slideResObj,
    slideMasterTextStyles: slideMasterTextStyles,
    layoutResObj: layoutResObj,
    masterResObj: masterResObj,
    themeContent: themeContent,
    digramFileContent: digramFileContent,
    diagramResObj: diagramResObj,
    defaultTextStyle: defaultTextStyle,
  }
  const bgColor = await getSlideBackgroundFill(warpObj)

  const elements = []
  for (const nodeKey in nodes) {
    if (nodes[nodeKey].constructor === Array) {
      for (const node of nodes[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, node, warpObj)
        if (ret) elements.push(ret)
      }
    } 
    else {
      const ret = await processNodesInSlide(nodeKey, nodes[nodeKey], warpObj)
      if (ret) elements.push(ret)
    }
  }

  return {
    fill: bgColor,
    elements,
  }
}

function indexNodes(content) {

  const keys = Object.keys(content)
  const spTreeNode = content[keys[0]]['p:cSld']['p:spTree']

  const idTable = {}
  const idxTable = {}
  const typeTable = {}

  for (const key in spTreeNode) {
    if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue

    const targetNode = spTreeNode[key]

    if (targetNode.constructor === Array) {
      for (const targetNodeItem of targetNode) {
        const nvSpPrNode = targetNodeItem['p:nvSpPr']
        const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
        const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
        const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

        if (id) idTable[id] = targetNodeItem
        if (idx) idxTable[idx] = targetNodeItem
        if (type) typeTable[type] = targetNodeItem
      }
    } 
    else {
      const nvSpPrNode = targetNode['p:nvSpPr']
      const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
      const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
      const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

      if (id) idTable[id] = targetNode
      if (idx) idxTable[idx] = targetNode
      if (type) typeTable[type] = targetNode
    }
  }

  return { idTable, idxTable, typeTable }
}

async function processNodesInSlide(nodeKey, nodeValue, warpObj) {
  let json

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = processSpNode(nodeValue, warpObj)
      break
    case 'p:cxnSp': // Shape, Text (with connection)
      json = processCxnSpNode(nodeValue, warpObj)
      break
    case 'p:pic': // Picture
      json = processPicNode(nodeValue, warpObj)
      break
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj)
      break
    case 'p:grpSp':
      json = await processGroupSpNode(nodeValue, warpObj)
      break
    default:
  }

  return json
}

async function processGroupSpNode(node, warpObj) {
  const xfrmNode = node['p:grpSpPr']['a:xfrm']
  const x = parseInt(xfrmNode['a:off']['attrs']['x']) * FACTOR
  const y = parseInt(xfrmNode['a:off']['attrs']['y']) * FACTOR
  const chx = parseInt(xfrmNode['a:chOff']['attrs']['x']) * FACTOR
  const chy = parseInt(xfrmNode['a:chOff']['attrs']['y']) * FACTOR
  const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * FACTOR
  const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * FACTOR
  const chcx = parseInt(xfrmNode['a:chExt']['attrs']['cx']) * FACTOR
  const chcy = parseInt(xfrmNode['a:chExt']['attrs']['cy']) * FACTOR

  const elements = []
  for (const nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (const item of node[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, item, warpObj)
        if (ret) elements.push(ret)
      }
    }
    else {
      const ret = await processNodesInSlide(nodeKey, node[nodeKey], warpObj)
      if (ret) elements.push(ret)
    }
  }

  return {
    type: 'group',
    top: y - chy,
    left: x - chx,
    width: cx - chcx,
    height: cy - chcy,
    elements,
  }
}

function processSpNode(node, warpObj) {
  const id = node['p:nvSpPr']['p:cNvPr']['attrs']['id']
  const name = node['p:nvSpPr']['p:cNvPr']['attrs']['name']
  const idx = node['p:nvSpPr']['p:nvPr']['p:ph'] ? node['p:nvSpPr']['p:nvPr']['p:ph']['attrs']['idx'] : undefined
  let type = node['p:nvSpPr']['p:nvPr']['p:ph'] ? node['p:nvSpPr']['p:nvPr']['p:ph']['attrs']['type'] : undefined

  let slideLayoutSpNode, slideMasterSpNode

  if (type) {
    if (idx) {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    } 
    else {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
  }
  else if (idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
    slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
  }

  if (!type) {
    const txBoxVal = getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox'])
    if (txBoxVal === '1') type = 'text'
  }
  if (!type) type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!type) type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  return genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, warpObj)
}

function processCxnSpNode(node, warpObj) {
  const id = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['id']
  const name = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['name']

  return genShape(node, undefined, undefined, id, name, undefined, undefined, warpObj)
}

function genShape(node, slideLayoutSpNode, slideMasterSpNode, id, name, idx, type, warpObj) {
  const xfrmList = ['p:spPr', 'a:xfrm']
  const slideXfrmNode = getTextByPathList(node, xfrmList)
  const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
  const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

  const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])

  const { top, left } = getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const { width, height } = getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)

  let isFlipV = false
  let isFlipH = false
  if (getTextByPathList(slideXfrmNode, ['attrs', 'flipV']) === '1') {
    isFlipV = true
  }
  if (getTextByPathList(slideXfrmNode, ['attrs', 'flipH']) === '1') {
    isFlipH = true
  }

  const rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ['attrs', 'rot']))

  const txtXframeNode = getTextByPathList(node, ['p:txXfrm'])
  let txtRotate
  if (txtXframeNode) {
    const txtXframeRot = getTextByPathList(txtXframeNode, ['attrs', 'rot'])
    if (txtXframeRot) txtRotate = angleToDegrees(txtXframeRot) + 90
  } 
  else txtRotate = rotate

  let content = ''
  if (node['p:txBody']) content = genTextBody(node['p:txBody'], slideLayoutSpNode, slideMasterSpNode, type, warpObj)

  const { borderColor, borderWidth, borderType } = getBorder(node, type)
  const fillColor = getShapeFill(node) || ''

  if (shapType) {
    const ext = getTextByPathList(slideXfrmNode, ['a:ext', 'attrs'])
    const cx = parseInt(ext['cx']) * FACTOR
    const cy = parseInt(ext['cy']) * FACTOR

    return {
      type: type === 'text' ? 'text' : 'shape',
      left,
      top,
      width,
      height,
      cx,
      cy,
      borderColor,
      borderWidth,
      borderType,
      fillColor,
      content,
      isFlipV,
      isFlipH,
      rotate,
      shapType,
      id,
      name,
      idx,
    }
  }

  return {
    type: 'text',
    left,
    top,
    width,
    height,
    borderColor,
    borderWidth,
    borderType,
    fillColor,
    isFlipV,
    isFlipH,
    rotate: txtRotate,
    content,
    id,
    name,
    idx,
  }
}

async function processPicNode(node, warpObj) {
  const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
  const imgName = warpObj['slideResObj'][rid]['target']
  const imgFileExt = extractFileExtension(imgName).toLowerCase()
  const zip = warpObj['zip']
  const imgArrayBuffer = await zip.file(imgName).async('arraybuffer')
  const xfrmNode = node['p:spPr']['a:xfrm']
  let mimeType = ''

  switch (imgFileExt) {
    case 'jpg':
    case 'jpeg':
      mimeType = 'image/jpeg'
      break
    case 'png':
      mimeType = 'image/png'
      break
    case 'gif':
      mimeType = 'image/gif'
      break
    case 'emf':
      mimeType = 'image/x-emf'
      break
    case 'wmf':
      mimeType = 'image/x-wmf'
      break
    default:
      mimeType = 'image/*'
  }
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  const src = `data:${mimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`

  let rotate = 0
  const rotateNode = getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot'])
  if (rotateNode) rotate = angleToDegrees(rotateNode)

  return {
    type: 'image',
    top,
    left,
    width, 
    height,
    src,
    rotate,
  }
}

async function processGraphicFrameNode(node, warpObj) {
  const graphicTypeUri = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri'])
  
  let result
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = genTable(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = genDiagram(node, warpObj)
      break
    default:
  }
  return result
}

function genTextBody(textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, warpObj) {
  if (!textBodyNode) return ''

  let text = ''
  const slideMasterTextStyles = warpObj['slideMasterTextStyles']

  const pNode = textBodyNode['a:p']
  const pNodes = pNode.constructor === Array ? pNode : [pNode]

  let isList = ''

  for (const pNode of pNodes) {
    const rNode = pNode['a:r']
    const align = getHorizontalAlign(pNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles)

    const listType = getListType(pNode)
    if (listType) {
      if (!isList) {
        text += `<${listType}>`
        isList = listType
      }
      else if (isList && isList !== listType) {
        text += `</${isList}>`
        text += `<${listType}>`
        isList = listType
      }
      text += `<li style="text-align: ${align}; color: ${type === 'text' ? '#000000' : '#ffffff'}">`
    }
    else {
      if (isList) {
        text += `</${isList}>`
        isList = ''
      }
      text += `<p style="text-align: ${align}; color: ${type === 'text' ? '#000000' : '#ffffff'}">`
    }
    
    if (!rNode) text += genSpanElement(pNode, slideLayoutSpNode, type, warpObj)
    else if (rNode.constructor === Array) {
      for (const rNodeItem of rNode) text += genSpanElement(rNodeItem, slideLayoutSpNode, type, warpObj)
    } 
    else text += genSpanElement(rNode, slideLayoutSpNode, type, warpObj)

    if (listType) text += '</li>'
    else text += '</p>'
  }
  return text
}

function getListType(node) {
  const pPrNode = node['a:pPr']
  if (!pPrNode) return ''

  if (pPrNode['a:buChar']) return 'ul'
  if (pPrNode['a:buAutoNum']) return 'ol'
  
  return ''
}

function genSpanElement(node, slideLayoutSpNode, type, warpObj) {
  const slideMasterTextStyles = warpObj['slideMasterTextStyles']

  let text = node['a:t']
  if (typeof text !== 'string') text = getTextByPathList(node, ['a:fld', 'a:t'])
  if (typeof text !== 'string') text = '&nbsp;'

  let styleText = ''
  const fontColor = getFontColor(node)
  const fontSize = getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles)
  const fontType = getFontType(node, type)
  const fontBold = getFontBold(node)
  const fontItalic = getFontItalic(node)
  const fontDecoration = getFontDecoration(node)
  if (fontColor) styleText += `color: ${fontColor};`
  if (fontSize) styleText += `font-size: ${fontSize};`
  if (fontType) styleText += `font-family: ${fontType};`
  if (fontBold) styleText += `font-weight: ${fontBold};`
  if (fontItalic) styleText += `font-style: ${fontItalic};`
  if (fontDecoration) styleText += `text-decoration: ${fontDecoration};`

  const linkID = getTextByPathList(node, ['a:rPr', 'a:hlinkClick', 'attrs', 'r:id'])
  if (linkID) {
    const linkURL = warpObj['slideResObj'][linkID]['target']
    return `<span style="${styleText}"><a href="${linkURL}" target="_blank">${text.replace(/\s/i, '&nbsp;')}</a></span>`
  } 
  return `<span style="${styleText}">${text.replace(/\s/i, '&nbsp;')}</span>`
}

function genTable(node, warpObj) {
  const tableNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl'])
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const trNodes = tableNode['a:tr']
  
  const data = []
  if (trNodes.constructor === Array) {
    for (const trNode of trNodes) {
      const tcNodes = trNode['a:tc']
      const tr = []

      if (tcNodes.constructor === Array) {
        for (const tcNode of tcNodes) {
          const text = genTextBody(tcNode['a:txBody'], undefined, undefined, undefined, warpObj)
          const rowSpan = getTextByPathList(tcNode, ['attrs', 'rowSpan'])
          const colSpan = getTextByPathList(tcNode, ['attrs', 'gridSpan'])
          const vMerge = getTextByPathList(tcNode, ['attrs', 'vMerge'])
          const hMerge = getTextByPathList(tcNode, ['attrs', 'hMerge'])

          tr.push({ text, rowSpan, colSpan, vMerge, hMerge })
        }
      } 
      else {
        const text = genTextBody(tcNodes['a:txBody'])
        tr.push({ text })
      }

      data.push(tr)
    }
  } 
  else {
    const tcNodes = trNodes['a:tc']
    const tr = []

    if (tcNodes.constructor === Array) {
      for (const tcNode of tcNodes) {
        const text = genTextBody(tcNode['a:txBody'])
        tr.push({ text })
      }
    } 
    else {
      const text = genTextBody(tcNodes['a:txBody'])
      tr.push({ text })
    }
    data.push(tr)
  }

  return {
    type: 'table',
    top,
    left,
    width,
    height,
    data,
  }
}

async function genChart(node, warpObj) {
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
  const refName = warpObj['slideResObj'][rid]['target']
  const content = await readXmlFile(warpObj['zip'], refName)
  const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])

  let chart = null
  for (const key in plotArea) {
    switch (key) {
      case 'c:lineChart':
        chart = {
          type: 'lineChart',
          data: extractChartData(plotArea[key]['c:ser']),
        }
        break
      case 'c:barChart':
        chart = {
          type: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']) === 'stacked' ? 'stackedBarChart' : 'barChart',
          data: extractChartData(plotArea[key]['c:ser']),
        }
        break
      case 'c:pieChart':
        chart = {
          type: 'pieChart',
          data: extractChartData(plotArea[key]['c:ser']),
        }
        break
      case 'c:pie3DChart':
        chart = {
          type: 'pie3DChart',
          data: extractChartData(plotArea[key]['c:ser']),
        }
        break
      case 'c:areaChart':
        chart = {
          type: getTextByPathList(plotArea[key], ['c:grouping', 'attrs', 'val']) === 'percentStacked' ? 'stackedAreaChart' : 'areaChart',
          data: extractChartData(plotArea[key]['c:ser']),
        }
        break
      case 'c:scatterChart':
        chart = {
          type: 'scatterChart',
          data: extractChartData(plotArea[key]['c:ser']),
        }
        break
      case 'c:catAx':
        break
      case 'c:valAx':
        break
      default:
    }
  }

  if (!chart) return {}
  return {
    type: 'chart',
    top,
    left,
    width,
    height,
    data: chart.data,
    chartType: chart.type,
  }
}

function genDiagram(node) {
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { left, top } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
  }
}

function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
  let off

  if (slideSpNode) off = slideSpNode['a:off']['attrs']
  else if (slideLayoutSpNode) off = slideLayoutSpNode['a:off']['attrs']
  else if (slideMasterSpNode) off = slideMasterSpNode['a:off']['attrs']

  if (!off) return { top: 0, left: 0 }

  return {
    top: parseInt(off['y']) * FACTOR,
    left: parseInt(off['x']) * FACTOR,
  }
}

function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
  let ext

  if (slideSpNode) ext = slideSpNode['a:ext']['attrs']
  else if (slideLayoutSpNode) ext = slideLayoutSpNode['a:ext']['attrs']
  else if (slideMasterSpNode) ext = slideMasterSpNode['a:ext']['attrs']

  if (!ext) return { width: 0, height: 0 }

  return {
    width: parseInt(ext['cx']) * FACTOR,
    height: parseInt(ext['cy']) * FACTOR,
  }
}

function getHorizontalAlign(node, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles) {
  let algn = getTextByPathList(node, ['a:pPr', 'attrs', 'algn'])

  if (!algn) algn = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])
  if (!algn) algn = getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])
  if (!algn) {
    switch (type) {
      case 'title':
      case 'subTitle':
      case 'ctrTitle':
        algn = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'attrs', 'alng'])
        break
      default:
        algn = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'attrs', 'alng'])
    }
  }
  if (!algn) {
    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') return 'center'
    else if (type === 'sldNum') return 'right'
  }
  return algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left'
}

function getFontType(node, type) {
  let typeface = getTextByPathList(node, ['a:rPr', 'a:latin', 'attrs', 'typeface'])

  if (!typeface) {
    const fontSchemeNode = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:fontScheme'])

    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      typeface = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface'])
    } 
    else if (type === 'body') {
      typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
    } 
    else {
      typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
    }
  }

  return typeface || ''
}

function getFontColor(node) {
  const color = getTextByPathList(node, ['a:rPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])
  return color ? `#${color}` : ''
}

function getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles) {
  let fontSize

  if (node['a:rPr']) fontSize = parseInt(node['a:rPr']['attrs']['sz']) / 100

  if ((isNaN(fontSize) || !fontSize)) {
    const sz = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:lstStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    fontSize = parseInt(sz) / 100
  }

  if (isNaN(fontSize) || !fontSize) {
    let sz
    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      sz = getTextByPathList(slideMasterTextStyles, ['p:titleStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    } 
    else if (type === 'body') {
      sz = getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    } 
    else if (type === 'dt' || type === 'sldNum') {
      sz = '1200'
    } 
    else if (!type) {
      sz = getTextByPathList(slideMasterTextStyles, ['p:otherStyle', 'a:lvl1pPr', 'a:defRPr', 'attrs', 'sz'])
    }
    if (sz) fontSize = parseInt(sz) / 100
  }

  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (baseline && !isNaN(fontSize)) fontSize -= 10

  return (isNaN(fontSize) || !fontSize) ? '18.75px' : (fontSize / 0.75 * (75 / 96) + 'px')
}

function getFontBold(node) {
  return (node['a:rPr'] && node['a:rPr']['attrs']['b'] === '1') ? 'bold' : ''
}

function getFontItalic(node) {
  return (node['a:rPr'] && node['a:rPr']['attrs']['i'] === '1') ? 'italic' : ''
}

function getFontDecoration(node) {
  return (node['a:rPr'] && node['a:rPr']['attrs']['u'] === 'sng') ? 'underline' : ''
}

function getBorder(node, elType) {
  const lineNode = node['p:spPr']['a:ln']

  let borderWidth = parseInt(getTextByPathList(lineNode, ['attrs', 'w'])) / 12700
  if (isNaN(borderWidth)) {
    if (lineNode) borderWidth = 0
    else if (elType === 'text') borderWidth = 0
    else borderWidth = 1
  }

  let borderColor = getTextByPathList(lineNode, ['a:solidFill', 'a:srgbClr', 'attrs', 'val'])
  if (!borderColor) {
    const schemeClrNode = getTextByPathList(lineNode, ['a:solidFill', 'a:schemeClr'])
    const schemeClr = 'a:' + getTextByPathList(schemeClrNode, ['attrs', 'val'])
    borderColor = getSchemeColorFromTheme(schemeClr)
  }

  if (!borderColor) {
    const schemeClrNode = getTextByPathList(node, ['p:style', 'a:lnRef', 'a:schemeClr'])
    const schemeClr = 'a:' + getTextByPathList(schemeClrNode, ['attrs', 'val'])
    borderColor = getSchemeColorFromTheme(schemeClr)

    if (borderColor) {
      let shade = getTextByPathList(schemeClrNode, ['a:shade', 'attrs', 'val'])

      if (shade) {
        shade = parseInt(shade) / 100000
        
        const color = tinycolor('#' + borderColor).toHsl()
        borderColor = tinycolor({ h: color.h, s: color.s, l: color.l * shade, a: color.a }).toHex()
      }
    }
  }

  if (!borderColor) borderColor = '#000'
  else borderColor = `#${borderColor}`

  const type = getTextByPathList(lineNode, ['a:prstDash', 'attrs', 'val'])
  let borderType = 'solid'
  let strokeDasharray = '0'
  switch (type) {
    case 'solid':
      borderType = 'solid'
      strokeDasharray = '0'
      break
    case 'dash':
      borderType = 'dashed'
      strokeDasharray = '5'
      break
    case 'dashDot':
      borderType = 'dashed'
      strokeDasharray = '5, 5, 1, 5'
      break
    case 'dot':
      borderType = 'dotted'
      strokeDasharray = '1, 5'
      break
    case 'lgDash':
      borderType = 'dashed'
      strokeDasharray = '10, 5'
      break
    case 'lgDashDotDot':
      borderType = 'dashed'
      strokeDasharray = '10, 5, 1, 5, 1, 5'
      break
    case 'sysDash':
      borderType = 'dashed'
      strokeDasharray = '5, 2'
      break
    case 'sysDashDot':
      borderType = 'dashed'
      strokeDasharray = '5, 2, 1, 5'
      break
    case 'sysDashDotDot':
      borderType = 'dashed'
      strokeDasharray = '5, 2, 1, 5, 1, 5'
      break
    case 'sysDot':
      borderType = 'dotted'
      strokeDasharray = '2, 5'
      break
    default:
  }

  return {
    borderColor,
    borderWidth,
    borderType,
    strokeDasharray,
  }
}

function getFillType(node) {
  let fillType = ''
  if (node['a:noFill']) fillType = 'NO_FILL'
  if (node['a:solidFill']) fillType = 'SOLID_FILL'
  if (node['a:gradFill']) fillType = 'GRADIENT_FILL'
  if (node['a:pattFill']) fillType = 'PATTERN_FILL'
  if (node['a:blipFill']) fillType = 'PIC_FILL'
  if (node['a:grpFill']) fillType = 'GROUP_FILL'

  return fillType
}

async function getPicFill(type, node, warpObj) {
  let img
  const rId = node['a:blip']['attrs']['r:embed']
  let imgPath
  if (type === 'slideBg' || type === 'slide') {
    imgPath = getTextByPathList(warpObj, ['slideResObj', rId, 'target'])
  }
  else if (type === 'slideLayoutBg') {
    imgPath = getTextByPathList(warpObj, ['layoutResObj', rId, 'target'])
  }
  else if (type === 'slideMasterBg') {
    imgPath = getTextByPathList(warpObj, ['masterResObj', rId, 'target'])
  }
  else if (type === 'themeBg') {
    imgPath = getTextByPathList(warpObj, ['themeResObj', rId, 'target'])
  }
  else if (type === 'diagramBg') {
    imgPath = getTextByPathList(warpObj, ['diagramResObj', rId, 'target'])
  }
  if (!imgPath) return imgPath

  img = getTextByPathList(warpObj, ['loaded-images', imgPath])
  if (!img) {
    imgPath = escapeHtml(imgPath)

    const imgExt = imgPath.split('.').pop()
    if (imgExt === 'xml') return undefined

    const imgArrayBuffer = await warpObj['zip'].file(imgPath).async('arraybuffer')
    const imgMimeType = getMimeType(imgExt)
    img = `data:${imgMimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`
  }
  return img
}

async function getBgPicFill(bgPr, sorce, warpObj) {
  const picBase64 = await getPicFill(sorce, bgPr['a:blipFill'], warpObj)
  const aBlipNode = bgPr['a:blipFill']['a:blip']

  const aphaModFixNode = getTextByPathList(aBlipNode, ['a:alphaModFix', 'attrs'])
  let opacity = 1
  if (aphaModFixNode && aphaModFixNode['amt'] && aphaModFixNode['amt'] !== '') {
    opacity = parseInt(aphaModFixNode['amt']) / 100000
  }

  return {
    picBase64,
    opacity,
  }
}

async function getSlideBackgroundFill(warpObj) {
  const slideContent = warpObj['slideContent']
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']
  
  let bgPr = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:bg', 'p:bgPr'])

  let background = '#fff'
  let backgroundType = 'color'

  if (bgPr) {
    const bgFillTyp = getFillType(bgPr)

    if (bgFillTyp === 'SOLID_FILL') {
      const sldFill = bgPr['a:solidFill']
      let clrMapOvr
      const sldClrMapOvr = getTextByPathList(slideContent, ['p:sld', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
      if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
      else {
        const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
        if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
        else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      }
      const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
      background = `#${sldBgClr}`
    }
    else if (bgFillTyp === 'PIC_FILL') {
      background = await getBgPicFill(bgPr, 'slideBg', warpObj)
      backgroundType = 'image'
    }
  }
  else {
    bgPr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:bg', 'p:bgPr'])

    let clrMapOvr
    const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
    if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
    else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])

    if (bgPr) {
      const bgFillTyp = getFillType(bgPr)
      if (bgFillTyp === 'SOLID_FILL') {
        const sldFill = bgPr['a:solidFill']
        const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
        background = `#${sldBgClr}`
      }
      else if (bgFillTyp === 'PIC_FILL') {
        background = await getBgPicFill(bgPr, 'slideLayoutBg', warpObj)
        backgroundType = 'image'
      }
    }
    else {
      bgPr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:bg', 'p:bgPr'])

      const clrMap = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      if (bgPr) {
        const bgFillTyp = getFillType(bgPr)
        if (bgFillTyp === 'SOLID_FILL') {
          const sldFill = bgPr['a:solidFill']
          const sldBgClr = getSolidFill(sldFill, clrMap, undefined, warpObj)
          background = `#${sldBgClr}`
        }
        else if (bgFillTyp === 'PIC_FILL') {
          background = await getBgPicFill(bgPr, 'slideMasterBg', warpObj)
          backgroundType = 'image'
        }
      }
    }
  }
  return {
    type: backgroundType,
    value: background,
  }
}

function getShapeFill(node, isSvgMode) {
  if (getTextByPathList(node, ['p:spPr', 'a:noFill'])) {
    return isSvgMode ? 'none' : ''
  }

  let fillColor
  if (!fillColor) {
    fillColor = getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])
  }

  if (!fillColor) {
    const schemeClr = 'a:' + getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'attrs', 'val'])
    fillColor = getSchemeColorFromTheme(schemeClr)
  }

  if (!fillColor) {
    const schemeClr = 'a:' + getTextByPathList(node, ['p:style', 'a:fillRef', 'a:schemeClr', 'attrs', 'val'])
    fillColor = getSchemeColorFromTheme(schemeClr)
  }

  if (fillColor) {
    fillColor = `#${fillColor}`

    let lumMod = parseInt(getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'a:lumMod', 'attrs', 'val'])) / 100000
    let lumOff = parseInt(getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'a:lumOff', 'attrs', 'val'])) / 100000
    if (isNaN(lumMod)) lumMod = 1.0
    if (isNaN(lumOff)) lumOff = 0

    const color = tinycolor(fillColor).toHsl()
    const lum = color.l * (1 + lumOff)
    return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHexString()
  } 

  if (isSvgMode) return 'none'
  return fillColor
}

function getSolidFill(solidFill) {
  if (!solidFill) return solidFill

  let color = 'fff'

  if (solidFill['a:srgbClr']) {
    color = getTextByPathList(solidFill['a:srgbClr'], ['attrs', 'val'])
  } 
  else if (solidFill['a:schemeClr']) {
    const schemeClr = 'a:' + getTextByPathList(solidFill['a:schemeClr'], ['attrs', 'val'])
    color = getSchemeColorFromTheme(schemeClr)
  }

  return color
}

function getSchemeColorFromTheme(schemeClr) {
  switch (schemeClr) {
    case 'a:tx1':
      schemeClr = 'a:dk1'
      break
    case 'a:tx2':
      schemeClr = 'a:dk2'
      break
    case 'a:bg1':
      schemeClr = 'a:lt1'
      break
    case 'a:bg2':
      schemeClr = 'a:lt2'
      break
    default: 
      break
  }
  const refNode = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme', schemeClr])
  let color = getTextByPathList(refNode, ['a:srgbClr', 'attrs', 'val'])
  if (!color) color = getTextByPathList(refNode, ['a:sysClr', 'attrs', 'lastClr'])
  return color
}

function extractChartData(serNode) {
  const dataMat = []
  if (!serNode) return dataMat

  if (serNode['c:xVal']) {
    let dataRow = []
    eachElement(serNode['c:xVal']['c:numRef']['c:numCache']['c:pt'], innerNode => {
      dataRow.push(parseFloat(innerNode['c:v']))
      return ''
    })
    dataMat.push(dataRow)
    dataRow = []
    eachElement(serNode['c:yVal']['c:numRef']['c:numCache']['c:pt'], innerNode => {
      dataRow.push(parseFloat(innerNode['c:v']))
      return ''
    })
    dataMat.push(dataRow)
  } 
  else {
    eachElement(serNode, (innerNode, index) => {
      const dataRow = []
      const colName = getTextByPathList(innerNode, ['c:tx', 'c:strRef', 'c:strCache', 'c:pt', 'c:v']) || index

      const rowNames = {}
      if (getTextByPathList(innerNode, ['c:cat', 'c:strRef', 'c:strCache', 'c:pt'])) {
        eachElement(innerNode['c:cat']['c:strRef']['c:strCache']['c:pt'], innerNode => {
          rowNames[innerNode['attrs']['idx']] = innerNode['c:v']
          return ''
        })
      } 
      else if (getTextByPathList(innerNode, ['c:cat', 'c:numRef', 'c:numCache', 'c:pt'])) {
        eachElement(innerNode['c:cat']['c:numRef']['c:numCache']['c:pt'], innerNode => {
          rowNames[innerNode['attrs']['idx']] = innerNode['c:v']
          return ''
        })
      }

      if (getTextByPathList(innerNode, ['c:val', 'c:numRef', 'c:numCache', 'c:pt'])) {
        eachElement(innerNode['c:val']['c:numRef']['c:numCache']['c:pt'], innerNode => {
          dataRow.push({
            x: innerNode['attrs']['idx'],
            y: parseFloat(innerNode['c:v']),
          })
          return ''
        })
      }

      dataMat.push({
        key: colName,
        values: dataRow,
        xlabels: rowNames,
      })
      return ''
    })
  }

  return dataMat
}