import { RATIO_EMUs_Points } from './constants'
import { decodeXmlEntities, getTextByPathList, numberToFixed } from './utils'
import { getFontSize } from './fontStyle'

function getParagraphLevel(node) {
  let lvlIdx = 1
  const lvlNode = getTextByPathList(node, ['a:pPr', 'attrs', 'lvl'])
  if (lvlNode !== undefined) lvlIdx = parseInt(lvlNode) + 1
  return lvlIdx
}

function getAlignFromTextNode(node, lvlStr) {
  if (!node) return ''

  let algn = getTextByPathList(node, ['p:txBody', 'a:lstStyle', lvlStr, 'attrs', 'algn'])
  if (!algn) algn = getTextByPathList(node, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])

  return algn || ''
}

export function getHorizontalAlign(node, pNode, type, slideLayoutSpNode, slideMasterSpNode, warpObj) {
  let algn = getTextByPathList(node, ['a:pPr', 'attrs', 'algn'])

  if (!algn) algn = getTextByPathList(pNode, ['p:txBody', 'a:p', 'a:pPr', 'attrs', 'algn'])

  if (!algn) {
    const lvlIdx = getParagraphLevel(node)
    const lvlStr = 'a:lvl' + lvlIdx + 'pPr'

    algn = getAlignFromTextNode(slideLayoutSpNode, lvlStr)
    if (!algn) algn = getAlignFromTextNode(slideMasterSpNode, lvlStr)

    if (!algn && (type === 'title' || type === 'ctrTitle' || type === 'subTitle')) {
      algn = getTextByPathList(warpObj, ['slideMasterTextStyles', 'p:titleStyle', lvlStr, 'attrs', 'algn'])
      if (!algn && type === 'subTitle') {
        algn = getTextByPathList(warpObj, ['slideMasterTextStyles', 'p:bodyStyle', lvlStr, 'attrs', 'algn'])
      }
    } 
    else if (!algn && type === 'body') {
      algn = getTextByPathList(warpObj, ['slideMasterTextStyles', 'p:bodyStyle', lvlStr, 'attrs', 'algn'])
    } 
    else if (!algn) {
      algn = getTextByPathList(warpObj, ['slideMasterTextStyles', 'p:otherStyle', lvlStr, 'attrs', 'algn'])
    }
  }

  let align = 'left'
  if (algn) {
    switch (algn) {
      case 'l':
        align = 'left'
        break
      case 'r':
        align = 'right'
        break
      case 'ctr':
        align = 'center'
        break
      case 'just':
        align = 'justify'
        break
      case 'dist':
        align = 'justify'
        break
      default:
        align = 'inherit'
    }
  }
  return align
}

export function getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode) {
  let anchor = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'anchor'])
  if (!anchor) {
    anchor = getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:bodyPr', 'attrs', 'anchor'])
    if (!anchor) {
      anchor = getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:bodyPr', 'attrs', 'anchor'])
      if (!anchor) anchor = 't'
    }
  }
  return (anchor === 'ctr') ? 'mid' : ((anchor === 'b') ? 'down' : 'up')
}

export function getTextDirection(node, slideLayoutSpNode, slideMasterSpNode) {
  for (const shape of [node, slideLayoutSpNode, slideMasterSpNode]) {
    const direction = getTextByPathList(shape, ['p:txBody', 'a:bodyPr', 'attrs', 'vert'])
    if (direction !== undefined && direction !== null && direction !== '') return direction
  }
  return 'horz'
}

export function getTextAutoFit(node, slideLayoutSpNode, slideMasterSpNode) {
  function checkBodyPr(bodyPr) {
    if (!bodyPr) return null

    if (bodyPr['a:noAutofit']) return { result: null }
    else if (bodyPr['a:spAutoFit']) return { result: { type: 'shape' } }
    else if (bodyPr['a:normAutofit']) {
      const fontScale = getTextByPathList(bodyPr['a:normAutofit'], ['attrs', 'fontScale'])
      if (fontScale) {
        const scalePercent = parseInt(fontScale) / 1000
        return {
          result: {
            type: 'text',
            fontScale: scalePercent,
          }
        }
      }
      return { result: { type: 'text' } }
    }
    return null
  }

  const nodeCheck = checkBodyPr(getTextByPathList(node, ['p:txBody', 'a:bodyPr']))
  if (nodeCheck) return nodeCheck.result

  const layoutCheck = checkBodyPr(getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:bodyPr']))
  if (layoutCheck) return layoutCheck.result

  const masterCheck = checkBodyPr(getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:bodyPr']))
  if (masterCheck) return masterCheck.result

  return null
}

function pushParagraphStyleNode(styleNodes, styleNode) {
  if (styleNode) styleNodes.push(styleNode)
}

function appendTextBodyParagraphStyleNodes(styleNodes, textBodyNode, lvl) {
  if (!textBodyNode) return

  const lvlPath = `a:lvl${lvl}pPr`
  pushParagraphStyleNode(styleNodes, getTextByPathList(textBodyNode, ['a:lstStyle', lvlPath]))
}

function appendShapeParagraphStyleNodes(styleNodes, shapeNode, lvl) {
  if (!shapeNode) return

  const lvlPath = `a:lvl${lvl}pPr`
  pushParagraphStyleNode(styleNodes, getTextByPathList(shapeNode, ['p:txBody', 'a:lstStyle', lvlPath]))
  pushParagraphStyleNode(styleNodes, getTextByPathList(shapeNode, ['p:txBody', 'a:p', 'a:pPr']))
}

function appendMasterTextParagraphStyleNodes(styleNodes, type, lvl, slideMasterTextStyles) {
  if (!slideMasterTextStyles) return

  const lvlPath = `a:lvl${lvl}pPr`

  if (type === 'title' || type === 'ctrTitle' || type === 'subTitle') {
    pushParagraphStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:titleStyle', lvlPath]))
    if (type === 'subTitle') {
      pushParagraphStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath]))
    }
  }
  else if (type === 'body') {
    pushParagraphStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath]))
  }
  else {
    pushParagraphStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:otherStyle', lvlPath]))
  }
}

function appendDefaultTextParagraphStyleNodes(styleNodes, defaultTextStyle, lvl) {
  if (!defaultTextStyle) return

  const lvlPath = `a:lvl${lvl}pPr`
  pushParagraphStyleNode(styleNodes, getTextByPathList(defaultTextStyle, [lvlPath]))
  pushParagraphStyleNode(styleNodes, getTextByPathList(defaultTextStyle, ['a:defPPr']))
}

export function getParagraphStyleNodes(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) {
  if (!pNode) return null

  const pPrNode = pNode['a:pPr']
  const lvl = getParagraphLevel(pNode)
  const styleNodes = []

  pushParagraphStyleNode(styleNodes, pPrNode)
  appendTextBodyParagraphStyleNodes(styleNodes, textBodyNode, lvl)
  appendShapeParagraphStyleNodes(styleNodes, slideLayoutSpNode, lvl)
  appendShapeParagraphStyleNodes(styleNodes, slideMasterSpNode, lvl)
  appendMasterTextParagraphStyleNodes(styleNodes, type, lvl, slideMasterTextStyles)
  appendDefaultTextParagraphStyleNodes(styleNodes, getTextByPathList(warpObj, ['defaultTextStyle']), lvl)

  return styleNodes
}

export function getParagraphListInfo(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) {
  const styleNodes = getParagraphStyleNodes(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) || []
  for (const style of styleNodes) {
    if (style['a:buNone'] || style['a:buBlip']) return { type: '' }
    if (style['a:buChar']) {
      const marker = decodeXmlEntities(getTextByPathList(style, ['a:buChar', 'attrs', 'char']) || '•')
      return { type: 'ul', marker, signature: 'ul:' + marker }
    }
    if (style['a:buAutoNum']) {
      const format = getTextByPathList(style, ['a:buAutoNum', 'attrs', 'type']) || 'arabicPeriod'
      const rawStart = getTextByPathList(style, ['a:buAutoNum', 'attrs', 'startAt'])
      const parsed = Number.parseInt(rawStart, 10)
      return {
        type: 'ol',
        format,
        start: Number.isFinite(parsed) && parsed > 0 ? parsed : 1,
        explicitStart: style === pNode['a:pPr'] && rawStart !== undefined,
        signature: 'ol:' + format,
      }
    }
  }
  return { type: '' }
}

function getLineSpacingValue(spacingNode, singleLineSpacingFactor = 1) {
  const spcPct = getTextByPathList(spacingNode, ['a:spcPct', 'attrs', 'val'])
  const spcPts = getTextByPathList(spacingNode, ['a:spcPts', 'attrs', 'val'])

  if (spcPct !== undefined) {
    const value = String(spcPct).trim()
    const percentage = value.endsWith('%') ? Number.parseFloat(value) / 100 : Number.parseInt(value, 10) / 100000
    if (Number.isFinite(percentage)) return percentage * singleLineSpacingFactor
  }
  if (spcPts) return parseInt(spcPts) / 100 + 'pt'

  return undefined
}

function getParagraphSpacingValue(spacingNode, referenceLineHeight) {
  const spcPct = getTextByPathList(spacingNode, ['a:spcPct', 'attrs', 'val'])
  const spcPts = getTextByPathList(spacingNode, ['a:spcPts', 'attrs', 'val'])

  if (spcPct !== undefined && spcPct !== null && spcPct !== '') {
    const value = String(spcPct).trim()
    const ratio = value.endsWith('%') ? Number.parseFloat(value) / 100 : Number.parseInt(value, 10) / 100000
    if (Number.isFinite(ratio)) return numberToFixed(ratio * referenceLineHeight) + 'pt'
  }
  if (spcPts !== undefined && spcPts !== null && spcPts !== '') return parseInt(spcPts) / 100 + 'pt'

  return undefined
}

function getParagraphReferenceLineHeight(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj, factor) {
  const lvl = getParagraphLevel(pNode)
  const runs = []
  for (const key of ['a:r', 'a:fld']) {
    const nodes = pNode[key]
    if (nodes) runs.push(...(Array.isArray(nodes) ? nodes : [nodes]))
  }
  if (!runs.length) runs.push(pNode)

  const fontSizes = runs.map(run => parseFloat(getFontSize(run, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, getTextByPathList(warpObj, ['defaultTextStyle']))))
  const size = Math.max(...fontSizes.filter(value => Number.isFinite(value) && value > 0))
  // The host paragraph's CSS em is not the text run's size. Resolve an explicit
  // point length; factor remains a configurable approximation of font metrics.
  return (Number.isFinite(size) ? size : 18) * factor
}

function getParagraphIndentValue(styleNode, attrName) {
  const val = getTextByPathList(styleNode, ['attrs', attrName])

  if (val !== undefined && val !== '') return numberToFixed(parseInt(val) * RATIO_EMUs_Points) + 'pt'

  return undefined
}

export function getParagraphSpacing(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) {
  const styleNodes = getParagraphStyleNodes(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj)
  if (!styleNodes) return null

  const spacing = {}
  const configuredFactor = getTextByPathList(warpObj, ['options', 'singleLineSpacingFactor'])
  const singleLineSpacingFactor = Number.isFinite(configuredFactor) && configuredFactor > 0 ? configuredFactor : 1
  let referenceLineHeight
  const paragraphSpacingValue = node => {
    if (getTextByPathList(node, ['a:spcPct', 'attrs', 'val']) !== undefined && referenceLineHeight === undefined) {
      referenceLineHeight = getParagraphReferenceLineHeight(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj, singleLineSpacingFactor)
    }
    return getParagraphSpacingValue(node, referenceLineHeight)
  }

  for (const styleNode of styleNodes) {
    if (spacing.lineSpacing === undefined) {
      const lineSpacing = getLineSpacingValue(styleNode['a:lnSpc'], singleLineSpacingFactor)
      if (lineSpacing !== undefined) spacing.lineSpacing = lineSpacing
    }

    if (spacing.spaceBefore === undefined) {
      const spaceBefore = paragraphSpacingValue(styleNode['a:spcBef'])
      if (spaceBefore !== undefined) spacing.spaceBefore = spaceBefore
    }

    if (spacing.spaceAfter === undefined) {
      const spaceAfter = paragraphSpacingValue(styleNode['a:spcAft'])
      if (spaceAfter !== undefined) spacing.spaceAfter = spaceAfter
    }
  }

  return Object.keys(spacing).length > 0 ? spacing : null
}

export function getParagraphIndent(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) {
  const styleNodes = getParagraphStyleNodes(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj)
  if (!styleNodes) return null

  const indent = {}

  for (const styleNode of styleNodes) {
    if (indent.marginLeft === undefined) {
      const marginLeft = getParagraphIndentValue(styleNode, 'marL')
      if (marginLeft !== undefined) indent.marginLeft = marginLeft
    }

    if (indent.textIndent === undefined) {
      const textIndent = getParagraphIndentValue(styleNode, 'indent')
      if (textIndent !== undefined) indent.textIndent = textIndent
    }
  }

  return Object.keys(indent).length > 0 ? indent : null
}
