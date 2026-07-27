import JSZip from 'jszip'
import { readXmlFile } from './readXmlFile'
import { getBorder } from './border'
import { getSlideBackgroundFill, getShapeFill, getSolidFill, getPicFill, getPicFilters, getImageData, getVideoData, getAudioData } from './fill'
import { getChartInfo } from './chart'
import { getVerticalAlign, getTextAutoFit } from './paragraph'
import { getTextInsets } from './textInsets'
import { getPosition, getSize } from './position'
import { genTextBody, getTextNodeValue } from './text'
import { getCustomShapePath, identifyShape, isStrokeOnlyCustomGeometry } from './shape'
import { extractFileExtension, getTextByPathList, angleToDegrees, isVideoLink, escapeHtml, hasValidText, numberToFixed, isOoxmlTrue } from './utils'
import { getShadow } from './shadow'
import { getTableBorders, getTableCellParams, getTableRowParams } from './table'
import { RATIO_EMUs_Points } from './constants'
import { findOMath, latexFormart, parseOMath } from './math'
import { getShapePath } from './shapePath'
import { parseTransition, findTransitionNode } from './animation'
import { getDiagramNodeContext, getSmartArtTextData } from './diagram'

export async function parse(file, options = {}) {
  const slides = []
  const loadedImages = {}
  const loadedVideos = {}
  const loadedAudios = {}
  const xmlCache = {}
  const parseOptions = {
    ...options,
    imageMode: options.imageMode || 'base64',
    videoMode: options.videoMode || 'none',
    audioMode: options.audioMode || 'none',
  }
  
  const zip = await JSZip.loadAsync(file)

  const filesInfo = await getContentTypes(zip)
  const { width, height, defaultTextStyle } = await getSlideInfo(zip)
  const { themeContent, themeColors } = await getTheme(zip)
  const usedFonts = await getUsedFonts(zip)

  for (const filename of filesInfo.slides) {
    const singleSlide = await processSingleSlide(zip, filename, themeContent, defaultTextStyle, loadedImages, loadedVideos, loadedAudios, parseOptions, xmlCache)
    slides.push(singleSlide)
  }

  return {
    slides,
    usedFonts,
    themeColors,
    size: {
      width,
      height,
    },
  }
}

async function getContentTypes(zip) {
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson['Types']['Override']
  let slidesLocArray = []
  let slideLayoutsLocArray = []

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
  
  const sortSlideXml = (p1, p2) => {
    const n1 = +/(\d+)\.xml/.exec(p1)[1]
    const n2 = +/(\d+)\.xml/.exec(p2)[1]
    return n1 - n2
  }
  slidesLocArray = slidesLocArray.sort(sortSlideXml)
  slideLayoutsLocArray = slideLayoutsLocArray.sort(sortSlideXml)
  
  return {
    slides: slidesLocArray,
    slideLayouts: slideLayoutsLocArray,
  }
}

async function getUsedFonts(zip) {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const embeddedFontList = getTextByPathList(content, ['p:presentation', 'p:embeddedFontLst', 'p:embeddedFont'])
  const usedFonts = []

  if (!embeddedFontList) return usedFonts

  const embeddedFonts = embeddedFontList.constructor === Array ? embeddedFontList : [embeddedFontList]
  for (const embeddedFont of embeddedFonts) {
    const typeface = getTextByPathList(embeddedFont, ['p:font', 'attrs', 'typeface'])
    if (typeface && !usedFonts.includes(typeface)) usedFonts.push(typeface)
  }

  return usedFonts
}

async function getSlideInfo(zip) {
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content['p:presentation']['p:sldSz']['attrs']
  const defaultTextStyle = content['p:presentation']['p:defaultTextStyle']
  return {
    width: parseInt(sldSzAttrs['cx']) * RATIO_EMUs_Points,
    height: parseInt(sldSzAttrs['cy']) * RATIO_EMUs_Points,
    defaultTextStyle,
  }
}

async function getTheme(zip) {
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

  const themeContent = await readXmlFile(zip, 'ppt/' + themeURI)

  const themeColors = []
  const clrScheme = getTextByPathList(themeContent, ['a:theme', 'a:themeElements', 'a:clrScheme'])
  if (clrScheme) {
    for (let i = 1; i <= 6; i++) {
      if (clrScheme[`a:accent${i}`] === undefined) break
      const color = getTextByPathList(clrScheme, [`a:accent${i}`, 'a:srgbClr', 'attrs', 'val'])
      if (color) themeColors.push('#' + color)
    }
  }

  return { themeContent, themeColors }
}

async function readXmlFileCached(zip, filename, xmlCache) {
  if (!filename) return null
  if (Object.prototype.hasOwnProperty.call(xmlCache, filename)) return xmlCache[filename]

  const content = await readXmlFile(zip, filename)
  xmlCache[filename] = content
  return content
}

async function processSingleSlide(zip, sldFileName, themeContent, defaultTextStyle, loadedImages, loadedVideos, loadedAudios, options, xmlCache) {
  const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels'
  const resContent = await readXmlFile(zip, resName)
  let relationshipArray = resContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
  
  let noteFilename = ''
  let layoutFilename = ''
  let masterFilename = ''
  let themeFilename = ''
  const slideResObj = {}
  const layoutResObj = {}
  const masterResObj = {}
  const themeResObj = {}

  for (const relationshipArrayItem of relationshipArray) {
    const relType = relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
    let relTarget = relationshipArrayItem['attrs']['Target']
    const isExternal = relationshipArrayItem['attrs']['TargetMode'] === 'External'
    if (!isExternal) {
      if (relTarget.indexOf('../') !== -1) relTarget = relTarget.replace('../', 'ppt/')
      else relTarget = 'ppt/slides/' + relTarget
    }

    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
        layoutFilename = relTarget
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
        noteFilename = relTarget
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors':
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget
        }
        break
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
      default:
        slideResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget,
        }
    }
  }
  
  const slideNotesContent = await readXmlFile(zip, noteFilename)
  const note = getNote(slideNotesContent)

  const slideLayoutContent = await readXmlFileCached(zip, layoutFilename, xmlCache)
  const slideLayoutTables = indexNodes(slideLayoutContent)
  const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
  const slideLayoutResContent = await readXmlFileCached(zip, slideLayoutResFilename, xmlCache)
  relationshipArray = slideLayoutResContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  for (const relationshipArrayItem of relationshipArray) {
    const relType = relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
    let relTarget = relationshipArrayItem['attrs']['Target']
    if (relTarget.indexOf('../') !== -1) relTarget = relTarget.replace('../', 'ppt/')
    else relTarget = 'ppt/slideLayouts/' + relTarget

    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
        masterFilename = relTarget
        break
      default:
        layoutResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget,
        }
    }
  }

  const slideMasterContent = await readXmlFileCached(zip, masterFilename, xmlCache)
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
  const slideMasterTables = indexNodes(slideMasterContent)
  const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
  const slideMasterResContent = await readXmlFileCached(zip, slideMasterResFilename, xmlCache)
  relationshipArray = slideMasterResContent['Relationships']['Relationship']
  if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]

  for (const relationshipArrayItem of relationshipArray) {
    const relType = relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', '')
    let relTarget = relationshipArrayItem['attrs']['Target']
    if (relTarget.indexOf('../') !== -1) relTarget = relTarget.replace('../', 'ppt/')
    else relTarget = 'ppt/slideMasters/' + relTarget

    switch (relationshipArrayItem['attrs']['Type']) {
      case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
        themeFilename = relTarget
        break
      default:
        masterResObj[relationshipArrayItem['attrs']['Id']] = {
          type: relType,
          target: relTarget,
        }
    }
  }

  let currentThemeContent = themeContent
  if (themeFilename) {
    currentThemeContent = currentThemeContent || await readXmlFileCached(zip, themeFilename, xmlCache)
    const themeName = themeFilename.split('/').pop()
    const themeResFileName = themeFilename.replace(themeName, '_rels/' + themeName) + '.rels'
    const themeResContent = await readXmlFile(zip, themeResFileName)
    if (themeResContent) {
      relationshipArray = themeResContent['Relationships']['Relationship']
      if (relationshipArray) {
        if (relationshipArray.constructor !== Array) relationshipArray = [relationshipArray]
        for (const relationshipArrayItem of relationshipArray) {
          themeResObj[relationshipArrayItem['attrs']['Id']] = {
            'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            'target': relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          }
        }
      }
    }
  }

  const tableStyles = await readXmlFileCached(zip, 'ppt/tableStyles.xml', xmlCache)

  const slideContent = await readXmlFile(zip, sldFileName)
  const nodes = slideContent['p:sld']['p:cSld']['p:spTree']
  const warpObj = {
    zip,
    loadedImages,
    loadedVideos,
    loadedAudios,
    options,
    slideLayoutContent,
    slideLayoutTables,
    slideMasterContent,
    slideMasterTables,
    slideContent,
    tableStyles,
    slideResObj,
    slideMasterTextStyles,
    layoutResObj,
    masterResObj,
    themeContent: currentThemeContent,
    themeResObj,
    diagramFileCache: {},
    defaultTextStyle,
  }
  const layoutElements = await getLayoutElements(warpObj)
  const fill = await getSlideBackgroundFill(warpObj)

  const elements = []
  for (const nodeKey in nodes) {
    if (nodes[nodeKey].constructor !== Array) nodes[nodeKey] = [nodes[nodeKey]]
    for (const node of nodes[nodeKey]) {
      const ret = await processNodesInSlide(nodeKey, node, warpObj, 'slide')
      if (ret) elements.push(ret)
    }
  }

  let transitionNode = findTransitionNode(slideContent, 'p:sld')
  if (!transitionNode) transitionNode = findTransitionNode(slideLayoutContent, 'p:sldLayout')
  if (!transitionNode) transitionNode = findTransitionNode(slideMasterContent, 'p:sldMaster')

  const transition = parseTransition(transitionNode)

  return {
    fill,
    elements,
    layoutElements,
    note,
    transition,
  }
}

function getHyperlinkFromCNvPr(cNvPr, warpObj) {
  const hlinkClick = getTextByPathList(cNvPr, ['a:hlinkClick', 'attrs'])
  if (!hlinkClick) return null

  const linkId = hlinkClick['r:id']
  if (!linkId) return null

  const res = warpObj['slideResObj'][linkId]
  if (!res) return null

  if (res['type'] !== 'hyperlink') return null

  const target = res['target']
  if (!target || !/^https?:\/\//.test(target)) return null

  return target
}

function getNote(noteContent) {
  let text = ''
  let spNodes = getTextByPathList(noteContent, ['p:notes', 'p:cSld', 'p:spTree', 'p:sp'])
  if (!spNodes) return ''

  if (spNodes.constructor !== Array) spNodes = [spNodes]
  for (const spNode of spNodes) {
    const phType = getTextByPathList(spNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
    if (phType !== 'body') continue

    const textBody = getTextByPathList(spNode, ['p:txBody'])
    if (!textBody) continue

    let pNode = textBody['a:p']
    if (!pNode) continue
    if (pNode.constructor !== Array) pNode = [pNode]

    const listTypes = []

    for (const p of pNode) {
      const pPr = p['a:pPr']
      const algn = getTextByPathList(pPr, ['attrs', 'algn'])
      let align = 'left'
      if (algn) {
        switch (algn) {
          case 'r': align = 'right'; break
          case 'ctr': align = 'center'; break
          case 'just': case 'dist': align = 'justify'; break
          default: break
        }
      }

      let listType = ''
      if (pPr) {
        if (pPr['a:buChar']) listType = 'ul'
        else if (pPr['a:buAutoNum']) listType = 'ol'
      }
      const lvlNode = getTextByPathList(pPr, ['attrs', 'lvl'])
      const listLevel = lvlNode !== undefined ? parseInt(lvlNode) : 0

      if (listType) {
        while (listTypes.length > listLevel + 1) {
          text += `</${listTypes.pop()}>`
        }
        if (listTypes[listLevel] === undefined) {
          text += `<${listType}>`
          listTypes[listLevel] = listType
        }
        else if (listTypes[listLevel] !== listType) {
          text += `</${listTypes[listLevel]}>`
          text += `<${listType}>`
          listTypes[listLevel] = listType
        }
        text += `<li><p style="text-align:${align};">`
      }
      else {
        while (listTypes.length > 0) {
          text += `</${listTypes.pop()}>`
        }
        text += `<p style="text-align:${align};">`
      }

      let rNodes = p['a:r']
      if (rNodes) {
        if (rNodes.constructor !== Array) rNodes = [rNodes]
        for (const r of rNodes) {
          const t = getTextNodeValue(getTextByPathList(r, ['a:t']))
          if (t && typeof t === 'string') text += t
        }
      }

      if (listType) text += '</p></li>'
      else text += '</p>'
    }
    while (listTypes.length > 0) {
      text += `</${listTypes.pop()}>`
    }
  }
  return text
}

async function getLayoutElements(warpObj) {
  const elements = []
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']
  const nodesSldLayout = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree'])
  const nodesSldMaster = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree'])

  const showMasterSp = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp'])
  if (nodesSldLayout) {
    for (const nodeKey in nodesSldLayout) {
      if (nodesSldLayout[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldLayout[nodeKey].length; i++) {
          const ph = getTextByPathList(nodesSldLayout[nodeKey][i], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
          if (!ph) {
            const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], warpObj, 'slideLayoutBg')
            if (ret) elements.push(ret)
          }
        }
      } 
      else {
        const ph = getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
        if (!ph) {
          const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], warpObj, 'slideLayoutBg')
          if (ret) elements.push(ret)
        }
      }
    }
  }
  if (nodesSldMaster && showMasterSp !== '0') {
    for (const nodeKey in nodesSldMaster) {
      if (nodesSldMaster[nodeKey].constructor === Array) {
        for (let i = 0; i < nodesSldMaster[nodeKey].length; i++) {
          const ph = getTextByPathList(nodesSldMaster[nodeKey][i], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
          if (!ph) {
            const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], warpObj, 'slideMasterBg')
            if (ret) elements.push(ret)
          }
        }
      } 
      else {
        const ph = getTextByPathList(nodesSldMaster[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph'])
        if (!ph) {
          const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], warpObj, 'slideMasterBg')
          if (ret) elements.push(ret)
        }
      }
    }
  }
  return elements
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
        if (type && !typeTable[type]) typeTable[type] = targetNodeItem
      }
    } 
    else {
      const nvSpPrNode = targetNode['p:nvSpPr']
      const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
      const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
      const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

      if (id) idTable[id] = targetNode
      if (idx) idxTable[idx] = targetNode
      if (type && !typeTable[type]) typeTable[type] = targetNode
    }
  }

  return { idTable, idxTable, typeTable }
}

async function processNodesInSlide(nodeKey, nodeValue, warpObj, source, groupHierarchy = []) {
  let json

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = await processSpNode(nodeValue, warpObj, source, groupHierarchy)
      break
    case 'p:cxnSp': // Shape, Text
      json = await processCxnSpNode(nodeValue, warpObj, source)
      break
    case 'p:pic': // Image, Video, Audio
      json = await processPicNode(nodeValue, warpObj, source)
      break
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj, source)
      break
    case 'p:grpSp':
      json = await processGroupSpNode(nodeValue, warpObj, source, groupHierarchy)
      break
    case 'mc:AlternateContent':
      if (getTextByPathList(nodeValue, ['mc:Fallback', 'p:grpSpPr', 'a:xfrm'])) {
        json = await processGroupSpNode(getTextByPathList(nodeValue, ['mc:Fallback']), warpObj, source, groupHierarchy)
      }
      else if (getTextByPathList(nodeValue, ['mc:Choice'])) {
        json = await processMathNode(nodeValue, warpObj, source)
      }
      break
    default:
  }

  return json
}

async function processMathNode(node, warpObj, source) {
  const choice = getTextByPathList(node, ['mc:Choice'])
  const fallback = getTextByPathList(node, ['mc:Fallback'])

  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(choice, ['p:sp', 'p:spPr', 'a:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const oMath = findOMath(choice)[0]
  if (!oMath) return null

  const latex = latexFormart(parseOMath(oMath))

  const blipFill = getTextByPathList(fallback, ['p:sp', 'p:spPr', 'a:blipFill'])
  const picFill = await getPicFill(source, blipFill, warpObj)

  let text = ''
  if (getTextByPathList(choice, ['p:sp', 'p:txBody', 'a:p', 'a:r'])) {
    const sp = getTextByPathList(choice, ['p:sp'])
    text = genTextBody(sp['p:txBody'], sp, undefined, undefined, undefined, warpObj)
  }

  return {
    type: 'math',
    top,
    left,
    width, 
    height,
    latex,
    picRef: picFill.ref,
    picBase64: picFill.base64,
    picBlob: picFill.blob,
    text,
    order,
  }
}

async function processGroupSpNode(node, warpObj, source, parentGroupHierarchy = []) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:grpSpPr', 'a:xfrm'])
  if (!xfrmNode) return null

  const x = parseInt(xfrmNode['a:off']['attrs']['x']) * RATIO_EMUs_Points
  const y = parseInt(xfrmNode['a:off']['attrs']['y']) * RATIO_EMUs_Points
  const chx = parseInt(xfrmNode['a:chOff']['attrs']['x']) * RATIO_EMUs_Points
  const chy = parseInt(xfrmNode['a:chOff']['attrs']['y']) * RATIO_EMUs_Points
  const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * RATIO_EMUs_Points
  const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * RATIO_EMUs_Points
  const chcx = parseInt(xfrmNode['a:chExt']['attrs']['cx']) * RATIO_EMUs_Points
  const chcy = parseInt(xfrmNode['a:chExt']['attrs']['cy']) * RATIO_EMUs_Points

  const isFlipV = isOoxmlTrue(getTextByPathList(xfrmNode, ['attrs', 'flipV']))
  const isFlipH = isOoxmlTrue(getTextByPathList(xfrmNode, ['attrs', 'flipH']))

  let rotate = getTextByPathList(xfrmNode, ['attrs', 'rot']) || 0
  if (rotate) rotate = angleToDegrees(rotate)

  // 计算缩放因子
  const ws = chcx === 0 ? 0 : cx / chcx
  const hs = chcy === 0 ? 0 : cy / chcy

  // 构建当前组合层级（将当前组合添加到父级层级中）
  const currentGroupHierarchy = [...parentGroupHierarchy, node]

  const elements = []
  for (const nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (const item of node[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, item, warpObj, source, currentGroupHierarchy)
        if (ret) elements.push(ret)
      }
    }
    else {
      const ret = await processNodesInSlide(nodeKey, node[nodeKey], warpObj, source, currentGroupHierarchy)
      if (ret) elements.push(ret)
    }
  }

  const transformGroupedElement = (element, offsetX = 0, offsetY = 0) => {
    const elementRotate = element.rotate || 0
    const normalizedRotate = ((elementRotate % 360) + 360) % 360
    const isUniformScale = Math.abs(ws - hs) < 0.000001
    const shouldSwapDimensions = normalizedRotate === 90 || normalizedRotate === 270
    const centerX = element.left + element.width / 2
    const centerY = element.top + element.height / 2
    const nextCenterX = (centerX - offsetX) * ws
    const nextCenterY = (centerY - offsetY) * hs
    const widthScale = shouldSwapDimensions && !isUniformScale ? hs : ws
    const heightScale = shouldSwapDimensions && !isUniformScale ? ws : hs
    const width = element.width * widthScale
    const height = element.height * heightScale

    const transformed = {
      ...element,
      left: numberToFixed(nextCenterX - width / 2),
      top: numberToFixed(nextCenterY - height / 2),
      width: numberToFixed(width),
      height: numberToFixed(height),
    }
    return transformed
  }

  const processedElements = elements.map(element => ({
    ...transformGroupedElement(element, chx, chy),
    ...(element.type === 'group' && element.elements ? {
      elements: processNestedGroupElements(element.elements)
    } : {})
  }))

  function processNestedGroupElements(elements, depth = 0) {
    if (depth > 10) return elements

    return elements.map(element => {
      const processed = transformGroupedElement(element)
      if (element.type === 'group' && element.elements) {
        processed.elements = processNestedGroupElements(element.elements, depth + 1)
      }
      return processed
    })
  }

  return {
    type: 'group',
    top: numberToFixed(y),
    left: numberToFixed(x),
    width: numberToFixed(cx),
    height: numberToFixed(cy),
    rotate,
    order,
    isFlipV,
    isFlipH,
    elements: processedElements,
  }
}

async function processSpNode(node, warpObj, source, groupHierarchy = []) {
  const cNvPr = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr'])
  const name = getTextByPathList(cNvPr, ['attrs', 'name'])
  const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  let type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  const order = getTextByPathList(node, ['attrs', 'order'])

  let slideLayoutSpNode, slideMasterSpNode

  if (type) {
    if (idx) {
      slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
      slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
      if (!slideLayoutSpNode) slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      if (!slideMasterSpNode) slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
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
    if (isOoxmlTrue(txBoxVal)) type = 'text'
  }
  if (!type) type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!type) type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!slideMasterSpNode && type === 'ctrTitle') slideMasterSpNode = warpObj['slideMasterTables']['typeTable']['title']

  if (!type) {
    if (source === 'diagramBg') type = 'diagram'
    else type = 'obj'
  }

  const link = getHyperlinkFromCNvPr(cNvPr, warpObj)

  return await genShape(node, slideLayoutSpNode, slideMasterSpNode, name, type, order, warpObj, source, link, groupHierarchy)
}

async function processCxnSpNode(node, warpObj, source) {
  const cNvPr = getTextByPathList(node, ['p:nvCxnSpPr', 'p:cNvPr'])
  const name = getTextByPathList(cNvPr, ['attrs', 'name'])
  const type = (node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined) ? undefined : node['p:nvCxnSpPr']['p:nvPr']['p:ph']['attrs']['type']
  const order = node['attrs']['order']
  const link = getHyperlinkFromCNvPr(cNvPr, warpObj)

  return await genShape(node, undefined, undefined, name, type, order, warpObj, source, link)
}

async function genShape(node, slideLayoutSpNode, slideMasterSpNode, name, type, order, warpObj, source, link, groupHierarchy = []) {
  const xfrmList = ['p:spPr', 'a:xfrm']
  const slideXfrmNode = getTextByPathList(node, xfrmList)
  const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
  const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

  const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  const custShapType = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  const keypoints = {}
  if (shapType) {
    const shapAdjst_ary = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'a:avLst', 'a:gd'])
    if (shapAdjst_ary) {
      const adjList = Array.isArray(shapAdjst_ary) ? shapAdjst_ary : [shapAdjst_ary]
      for (const adj of adjList) {
        const name = getTextByPathList(adj, ['attrs', 'name'])
        const fmla = getTextByPathList(adj, ['attrs', 'fmla'])
        if (name && fmla && fmla.startsWith('val ')) {
          keypoints[name] = parseInt(fmla.substring(4)) / 50000
        }
      }
    }
  }

  const { top, left } = getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const { width, height } = getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const pathViewBox = { x: 0, y: 0, width, height }

  const isFlipV = isOoxmlTrue(getTextByPathList(slideXfrmNode, ['attrs', 'flipV']))
  const isFlipH = isOoxmlTrue(getTextByPathList(slideXfrmNode, ['attrs', 'flipH']))

  const rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ['attrs', 'rot']))

  const txtXframeNode = getTextByPathList(node, ['p:txXfrm'])
  let txtRotate
  if (txtXframeNode) {
    const txtXframeRot = getTextByPathList(txtXframeNode, ['attrs', 'rot'])
    if (txtXframeRot) txtRotate = angleToDegrees(txtXframeRot) + 90
  } 
  else txtRotate = rotate

  let content = ''
  if (node['p:txBody']) content = genTextBody(node['p:txBody'], node, slideLayoutSpNode, slideMasterSpNode, type, warpObj)

  const { borderColor, borderWidth, borderType, strokeDasharray, headEnd, tailEnd } = getBorder(node, type, warpObj)
  const fill = await getShapeFill(node, warpObj, source, {
    groupHierarchy,
    slideLayoutSpNode,
    slideMasterSpNode,
  })

  let shadow
  const outerShdwNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
  if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

  const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)
  // Vertical (East-Asian) writing mode. Read once here and reported on every branch below: a
  // <a:bodyPr vert="eaVert"> is just as valid on a shape as on a plain text box, and dropping it
  // for shapes leaves the consumer laying that text out horizontally.
  const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'
  const autoFit = getTextAutoFit(node, slideLayoutSpNode, slideMasterSpNode)
  const textInset = getTextInsets(node, slideLayoutSpNode, slideMasterSpNode)

  const data = {
    left,
    top,
    width,
    height,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    fill,
    content,
    isFlipV,
    isFlipH,
    rotate,
    vAlign,
    name,
    order,
  }

  if (shadow) data.shadow = shadow
  if (autoFit) data.autoFit = autoFit
  if (link) data.link = link
  if (textInset) data.textInset = textInset
  if (headEnd) data.headEnd = headEnd
  if (tailEnd) data.tailEnd = tailEnd

  const isHasValidText = data.content && hasValidText(data.content)

  if (custShapType && type !== 'diagram') {
    const ext = getTextByPathList(slideXfrmNode, ['a:ext', 'attrs'])
    const w = parseInt(ext['cx']) * RATIO_EMUs_Points
    const h = parseInt(ext['cy']) * RATIO_EMUs_Points
    const d = getCustomShapePath(custShapType, w, h)
    if (!isHasValidText) data.content = ''

    const customShapeData = {
      ...data,
      type: 'shape',
      isVertical,
      shapType: 'custom',
      path: d,
      pathViewBox: { x: 0, y: 0, width: w, height: h },
    }
    if (isStrokeOnlyCustomGeometry(custShapType)) customShapeData.strokeOnly = true

    return customShapeData
  }

  let shapePath = ''
  if (shapType) shapePath = getShapePath(shapType, width, height, node)
  const STROKE_ONLY_PRESET_SHAPE_TYPES = ['arc', 'leftBrace', 'rightBrace', 'bracePair', 'leftBracket', 'rightBracket', 'bracketPair']
  const isStrokeOnlyPresetShape = STROKE_ONLY_PRESET_SHAPE_TYPES.includes(shapType)

  if (shapType && (type === 'obj' || !type || shapType !== 'rect')) {
    if (!isHasValidText) data.content = ''
    const shapeData = {
      ...data,
      type: 'shape',
      isVertical,
      shapType,
      path: shapePath,
      pathViewBox,
      keypoints,
    }
    if (isStrokeOnlyPresetShape) shapeData.strokeOnly = true
    return shapeData
  }
  if (shapType && !isHasValidText && (fill || borderWidth)) {
    const shapeData = {
      ...data,
      type: 'shape',
      isVertical,
      content: '',
      shapType,
      path: shapePath,
      pathViewBox,
      keypoints,
    }
    if (isStrokeOnlyPresetShape) shapeData.strokeOnly = true
    return shapeData
  }
  return {
    ...data,
    type: 'text',
    isVertical,
    rotate: txtRotate,
  }
}

async function processPicNode(node, warpObj, source) {
  let resObj
  if (source === 'slideMasterBg') resObj = warpObj['masterResObj']
  else if (source === 'slideLayoutBg') resObj = warpObj['layoutResObj']
  else resObj = warpObj['slideResObj']

  const cNvPr = getTextByPathList(node, ['p:nvPicPr', 'p:cNvPr'])
  const link = getHyperlinkFromCNvPr(cNvPr, warpObj)
  const order = node['attrs']['order']
  
  const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
  
  if (!rid || !resObj[rid]) return null

  const imgName = resObj[rid]['target']

  let xfrmNode = node['p:spPr']['a:xfrm']
  if (!xfrmNode) {
    const idx = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
    if (idx) xfrmNode = getTextByPathList(warpObj['slideLayoutTables'], ['idxTable', idx, 'p:spPr', 'a:xfrm'])
  }

  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  const imageData = await getImageData(imgName, warpObj)

  const isFlipV = isOoxmlTrue(getTextByPathList(xfrmNode, ['attrs', 'flipV']))
  const isFlipH = isOoxmlTrue(getTextByPathList(xfrmNode, ['attrs', 'flipH']))

  let rotate = 0
  const rotateNode = getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot'])
  if (rotateNode) rotate = angleToDegrees(rotateNode)

  const videoNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile'])
  let videoRid, videoFile, videoFileExt
  let videoData = {
    ref: '',
    blob: '',
  }
  let isVdeoLink = false

  if (videoNode) {
    videoRid = videoNode['attrs']['r:link']
    videoFile = resObj[videoRid]['target']
    if (isVideoLink(videoFile)) {
      videoFile = escapeHtml(videoFile)
      isVdeoLink = true
    } 
    else {
      videoFileExt = extractFileExtension(videoFile).toLowerCase()
      if (videoFileExt === 'mp4' || videoFileExt === 'webm' || videoFileExt === 'ogg') {
        videoData = await getVideoData(videoFile, warpObj)
      }
      else {
        videoData = {
          ref: videoFile,
          blob: '',
        }
      }
    }
    if (isVdeoLink) {
      videoData = {
        ref: videoFile,
        blob: '',
      }
    }
  }

  const audioNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile'])
  let audioRid, audioFile, audioFileExt
  let audioData = {
    ref: '',
    blob: '',
  }
  if (audioNode) {
    audioRid = audioNode['attrs']['r:link']
    audioFile = resObj[audioRid]['target']
    audioFileExt = extractFileExtension(audioFile).toLowerCase()
    if (audioFileExt === 'mp3' || audioFileExt === 'wav' || audioFileExt === 'ogg') {
      audioData = await getAudioData(audioFile, warpObj)
    }
    else {
      audioData = {
        ref: audioFile,
        blob: '',
      }
    }
  }

  if (videoNode && !isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width, 
      height,
      rotate,
      ref: videoData.ref,
      blob: videoData.blob,
      order,
    }
  } 
  if (videoNode && isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width, 
      height,
      rotate,
      ref: videoData.ref,
      blob: videoData.blob,
      order,
    }
  }
  if (audioNode) {
    return {
      type: 'audio',
      top,
      left,
      width, 
      height,
      rotate,
      ref: audioData.ref,
      blob: audioData.blob,
      order,
    }
  }

  let rect
  const srcRectAttrs = getTextByPathList(node, ['p:blipFill', 'a:srcRect', 'attrs'])
  if (srcRectAttrs && (srcRectAttrs.t || srcRectAttrs.b || srcRectAttrs.l || srcRectAttrs.r)) {
    rect = {}
    if (srcRectAttrs.t) rect.t = srcRectAttrs.t / 1000
    if (srcRectAttrs.b) rect.b = srcRectAttrs.b / 1000
    if (srcRectAttrs.l) rect.l = srcRectAttrs.l / 1000
    if (srcRectAttrs.r) rect.r = srcRectAttrs.r / 1000
  }
  let geom = 'rect'
  const prstGeom = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  const custGeom = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  if (prstGeom) {
    geom = prstGeom
  }
  else if (custGeom) {
    geom = identifyShape(custGeom)
    if (geom !== 'custom') geom = `custom:${geom}`
  }

  const { borderColor, borderWidth, borderType, strokeDasharray } = getBorder(node, undefined, warpObj)

  const filters = getPicFilters(node['p:blipFill'])

  const imageDataJson = {
    type: 'image',
    top,
    left,
    width,
    height,
    rotate,
    ref: imageData.ref,
    base64: imageData.base64,
    blob: imageData.blob,
    isFlipV,
    isFlipH,
    order,
    rect,
    geom,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
  }

  if (filters) imageDataJson.filters = filters
  if (link) imageDataJson.link = link

  return imageDataJson
}

async function processGraphicFrameNode(node, warpObj, source) {
  const graphicTypeUri = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri'])
  
  let result
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = await genTable(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = await genDiagram(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
      let oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj'])
      if (!oleObjNode) oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj'])
      if (oleObjNode) result = await processGroupSpNode(oleObjNode, warpObj, source)
      break
    default:
  }
  return result
}

async function genTable(node, warpObj) {
  const order = node['attrs']['order']
  const tableNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl'])
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const getTblPr = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblPr'])
  let getColsGrid = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblGrid', 'a:gridCol'])
  if (getColsGrid.constructor !== Array) getColsGrid = [getColsGrid]

  const colWidths = []
  if (getColsGrid) {
    for (const item of getColsGrid) {
      const colWidthParam = getTextByPathList(item, ['attrs', 'w']) || 0
      const colWidth = parseInt(colWidthParam) * RATIO_EMUs_Points
      colWidths.push(colWidth)
    }
  }

  const firstRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstRow'] : undefined
  const firstColAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstCol'] : undefined
  const lastRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastRow'] : undefined
  const lastColAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastCol'] : undefined
  const bandRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandRow'] : undefined
  const bandColAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandCol'] : undefined
  const tblStylAttrObj = {
    isFrstRowAttr: isOoxmlTrue(firstRowAttr) ? 1 : 0,
    isFrstColAttr: isOoxmlTrue(firstColAttr) ? 1 : 0,
    isLstRowAttr: isOoxmlTrue(lastRowAttr) ? 1 : 0,
    isLstColAttr: isOoxmlTrue(lastColAttr) ? 1 : 0,
    isBandRowAttr: isOoxmlTrue(bandRowAttr) ? 1 : 0,
    isBandColAttr: isOoxmlTrue(bandColAttr) ? 1 : 0,
  }

  let thisTblStyle
  const tbleStyleId = getTblPr['a:tableStyleId']
  if (tbleStyleId) {
    const tbleStylList = warpObj['tableStyles']['a:tblStyleLst']['a:tblStyle']
    if (tbleStylList) {
      if (tbleStylList.constructor === Array) {
        for (let k = 0; k < tbleStylList.length; k++) {
          if (tbleStylList[k]['attrs']['styleId'] === tbleStyleId) {
            thisTblStyle = tbleStylList[k]
          }
        }
      } 
      else {
        if (tbleStylList['attrs']['styleId'] === tbleStyleId) {
          thisTblStyle = tbleStylList
        }
      }
    }
  }
  if (thisTblStyle) thisTblStyle['tblStylAttrObj'] = tblStylAttrObj

  let borders = {}
  const tblStyl = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle'])
  const tblBorderStyl = getTextByPathList(tblStyl, ['a:tcBdr'])
  if (tblBorderStyl) borders = getTableBorders(tblBorderStyl, warpObj)

  let tbl_bgcolor = ''
  let tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:tblBg', 'a:fillRef'])
  if (tbl_bgFillschemeClr) {
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }
  if (tbl_bgFillschemeClr === undefined) {
    tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }

  let trNodes = tableNode['a:tr']
  if (trNodes.constructor !== Array) trNodes = [trNodes]
  
  const data = []
  const rowHeights = []
  for (let i = 0; i < trNodes.length; i++) {
    const trNode = trNodes[i]
    
    const rowHeightParam = getTextByPathList(trNodes[i], ['attrs', 'h']) || 0
    const rowHeight = parseInt(rowHeightParam) * RATIO_EMUs_Points
    rowHeights.push(rowHeight)

    const {
      fillColor,
      fontColor,
      fontBold,
    } = getTableRowParams(trNodes, i, tblStylAttrObj, thisTblStyle, warpObj)

    const tcNodes = trNode['a:tc']
    const tr = []

    if (tcNodes.constructor === Array) {
      for (let j = 0; j < tcNodes.length; j++) {
        const tcNode = tcNodes[j]
        let a_sorce
        if (j === 0 && tblStylAttrObj['isFrstColAttr'] === 1) {
          a_sorce = 'a:firstCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:seCell'])) {
            a_sorce = 'a:seCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 &&
            getTextByPathList(thisTblStyle, ['a:neCell'])) {
            a_sorce = 'a:neCell'
          }
        } 
        else if (
          (j > 0 && tblStylAttrObj['isBandColAttr'] === 1) &&
          !(tblStylAttrObj['isFrstColAttr'] === 1 && i === 0) &&
          !(tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1)) &&
          j !== (tcNodes.length - 1)
        ) {
          if ((j % 2) !== 0) {
            let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
            if (aBandNode === undefined) {
              aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
              if (aBandNode) a_sorce = 'a:band2V'
            } 
            else a_sorce = 'a:band2V'
          }
        }
        if (j === (tcNodes.length - 1) && tblStylAttrObj['isLstColAttr'] === 1) {
          a_sorce = 'a:lastCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:swCell'])) {
            a_sorce = 'a:swCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 && getTextByPathList(thisTblStyle, ['a:nwCell'])) {
            a_sorce = 'a:nwCell'
          }
        }
        const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, undefined, warpObj)
        const cell = await getTableCellParams(tcNode, thisTblStyle, a_sorce, warpObj)
        const td = { text }
        if (cell.rowSpan) td.rowSpan = cell.rowSpan
        if (cell.colSpan) td.colSpan = cell.colSpan
        if (cell.vMerge) td.vMerge = cell.vMerge
        if (cell.hMerge) td.hMerge = cell.hMerge
        if (cell.vAlign) td.vAlign = cell.vAlign
        if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
        if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
        if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor
        if (cell.borders) td.borders = cell.borders

        tr.push(td)
      }
    } 
    else {
      let a_sorce
      if (tblStylAttrObj['isFrstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:firstCol'
      } 
      else if (tblStylAttrObj['isBandColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
        if (!aBandNode) {
          aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
          if (aBandNode) a_sorce = 'a:band2V'
        } 
        else a_sorce = 'a:band2V'
      }
      if (tblStylAttrObj['isLstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:lastCol'
      }

      const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, undefined, warpObj)
      const cell = await getTableCellParams(tcNodes, thisTblStyle, a_sorce, warpObj)
      const td = { text }
      if (cell.rowSpan) td.rowSpan = cell.rowSpan
      if (cell.colSpan) td.colSpan = cell.colSpan
      if (cell.vMerge) td.vMerge = cell.vMerge
      if (cell.hMerge) td.hMerge = cell.hMerge
      if (cell.vAlign) td.vAlign = cell.vAlign
      if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
      if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
      if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor
      if (cell.borders) td.borders = cell.borders

      tr.push(td)
    }
    data.push(tr)
  }

  let actualTableWidth = colWidths.reduce((sum, width) => sum + width, 0)
  if (actualTableWidth) actualTableWidth = numberToFixed(actualTableWidth)

  return {
    type: 'table',
    top,
    left,
    width: actualTableWidth || width,
    height,
    data,
    order,
    borders,
    rowHeights,
    colWidths,
  }
}

async function genChart(node, warpObj) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
  let refName = getTextByPathList(warpObj['slideResObj'], [rid, 'target'])
  if (!refName) refName = getTextByPathList(warpObj['layoutResObj'], [rid, 'target'])
  if (!refName) refName = getTextByPathList(warpObj['masterResObj'], [rid, 'target'])
  if (!refName) return null

  const content = await readXmlFile(warpObj['zip'], refName)
  const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])

  const chart = getChartInfo(plotArea, warpObj)

  if (!chart) return null

  const data = {
    type: 'chart',
    top,
    left,
    width,
    height,
    data: chart.data,
    colors: chart.colors,
    chartType: chart.type,
    order,
  }
  if (chart.marker !== undefined) data.marker = chart.marker
  if (chart.barDir !== undefined) data.barDir = chart.barDir
  if (chart.holeSize !== undefined) data.holeSize = chart.holeSize
  if (chart.grouping !== undefined) data.grouping = chart.grouping
  if (chart.style !== undefined) data.style = chart.style

  return data
}

async function genDiagram(node, warpObj) {
  const order = node['attrs']['order']
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { left, top } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const diagramWarpObj = await getDiagramNodeContext(node, warpObj)
  const dgmDrwSpArray = getTextByPathList(diagramWarpObj['digramFileContent'], ['p:drawing', 'p:spTree', 'p:sp'])
  const elements = []
  let textList = []
  if (dgmDrwSpArray) {
    const spList = Array.isArray(dgmDrwSpArray) ? dgmDrwSpArray : [dgmDrwSpArray]

    for (const item of spList) {
      const el = await processSpNode(item, diagramWarpObj, 'diagramBg')
      if (el) elements.push(el)
    }
  }
  if (diagramWarpObj.diagramContent && diagramWarpObj.diagramContent.data) {
    textList = getSmartArtTextData(diagramWarpObj.diagramContent.data)
  }

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
    elements,
    textList,
    order,
  }
}
