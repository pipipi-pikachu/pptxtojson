import { getTextByPathList } from './utils'

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

function getParagraphStyleNodes(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) {
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

// PPT `<a:lnSpc><a:spcPct val="N"/>` is a percent of "single-line-spacing",
// not a percent of font-size. Single-line-spacing is the font's natural line
// height (ascent + descent + leading) and is roughly 1.2 * font-size for the
// fonts PowerPoint ships with. Emitting the raw percentage as a unitless CSS
// `line-height` collapses every paragraph by ~20%, which breaks slides that
// overlay absolute-positioned shapes onto specific text lines (e.g. poem-pause
// connectors floating below the last text line). Bake the leading factor in
// before returning so HTML consumers get the same vertical rhythm PowerPoint
// renders.
const SINGLE_LINE_SPACING_LEADING = 1.2

function getLineSpacingValue(spacingNode) {
  const spcPct = getTextByPathList(spacingNode, ['a:spcPct', 'attrs', 'val'])
  const spcPts = getTextByPathList(spacingNode, ['a:spcPts', 'attrs', 'val'])

  if (spcPct) return (parseInt(spcPct) / 1000 / 100) * SINGLE_LINE_SPACING_LEADING
  if (spcPts) return parseInt(spcPts) / 100 + 'pt'

  return undefined
}

function getParagraphSpacingValue(spacingNode) {
  const spcPct = getTextByPathList(spacingNode, ['a:spcPct', 'attrs', 'val'])
  const spcPts = getTextByPathList(spacingNode, ['a:spcPts', 'attrs', 'val'])

  if (spcPct) return parseInt(spcPct) / 1000 + 'em'
  if (spcPts) return parseInt(spcPts) / 100 + 'pt'

  return undefined
}

export function getParagraphSpacing(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj) {
  const styleNodes = getParagraphStyleNodes(pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, warpObj)
  if (!styleNodes) return null

  const spacing = {}

  for (const styleNode of styleNodes) {
    if (spacing.lineSpacing === undefined) {
      const lineSpacing = getLineSpacingValue(styleNode['a:lnSpc'])
      if (lineSpacing !== undefined) spacing.lineSpacing = lineSpacing
    }

    if (spacing.spaceBefore === undefined) {
      const spaceBefore = getParagraphSpacingValue(styleNode['a:spcBef'])
      if (spaceBefore !== undefined) spacing.spaceBefore = spaceBefore
    }

    if (spacing.spaceAfter === undefined) {
      const spaceAfter = getParagraphSpacingValue(styleNode['a:spcAft'])
      if (spaceAfter !== undefined) spacing.spaceAfter = spaceAfter
    }
  }

  return Object.keys(spacing).length > 0 ? spacing : null
}
