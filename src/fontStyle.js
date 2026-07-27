import { getTextByPathList, isOoxmlTrue } from './utils'
import { getShadow } from './shadow'
import { getFillType, getGradientFill, getSolidFill } from './fill'

function pushStyleNode(styleNodes, styleNode) {
  if (styleNode) styleNodes.push(styleNode)
}

function getLevelPath(lvl) {
  return `a:lvl${lvl}pPr`
}

function appendTextBodyStyleNodes(styleNodes, textBodyNode, lvl) {
  if (!textBodyNode) return

  const lvlPath = getLevelPath(lvl)
  pushStyleNode(styleNodes, getTextByPathList(textBodyNode, ['a:lstStyle', lvlPath, 'a:defRPr']))
}

function appendShapeStyleNodes(styleNodes, shapeNode, lvl) {
  if (!shapeNode) return

  const lvlPath = getLevelPath(lvl)
  pushStyleNode(styleNodes, getTextByPathList(shapeNode, ['p:txBody', 'a:lstStyle', lvlPath, 'a:defRPr']))
  pushStyleNode(styleNodes, getTextByPathList(shapeNode, ['p:txBody', 'a:p', 'a:pPr', 'a:defRPr']))
}

function appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles) {
  if (!slideMasterTextStyles) return

  const lvlPath = getLevelPath(lvl)

  if (type === 'title' || type === 'ctrTitle' || type === 'subTitle') {
    pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:titleStyle', lvlPath, 'a:defRPr']))
    if (type === 'subTitle') {
      pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath, 'a:defRPr']))
    }
  }
  else if (type === 'body') {
    pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath, 'a:defRPr']))
  }
  else {
    pushStyleNode(styleNodes, getTextByPathList(slideMasterTextStyles, ['p:otherStyle', lvlPath, 'a:defRPr']))
  }
}

function appendDefaultTextStyleNodes(styleNodes, lvl, defaultTextStyle) {
  if (!defaultTextStyle) return

  const lvlPath = getLevelPath(lvl)
  pushStyleNode(styleNodes, getTextByPathList(defaultTextStyle, [lvlPath, 'a:defRPr']))
  pushStyleNode(styleNodes, getTextByPathList(defaultTextStyle, ['a:defPPr', 'a:defRPr']))
}

function getBaseFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, lvl) {
  const styleNodes = []
  const runStyleNode = getTextByPathList(node, ['a:rPr'])

  pushStyleNode(styleNodes, runStyleNode)
  if (!runStyleNode) {
    pushStyleNode(styleNodes, getTextByPathList(pNode, ['a:endParaRPr']))
  }
  pushStyleNode(styleNodes, getTextByPathList(pNode, ['a:pPr', 'a:defRPr']))

  appendTextBodyStyleNodes(styleNodes, textBodyNode, lvl)
  appendShapeStyleNodes(styleNodes, slideLayoutSpNode, lvl)
  appendShapeStyleNodes(styleNodes, slideMasterSpNode, lvl)

  return styleNodes
}

function getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getBaseFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, lvl)
  appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles)

  return styleNodes
}

function getFontAttr(styleNodes, attrName) {
  for (const styleNode of styleNodes) {
    const attrValue = getTextByPathList(styleNode, ['attrs', attrName])
    if (attrValue !== undefined && attrValue !== '') return attrValue
  }

  return ''
}

function getFontTypeface(styleNodes) {
  for (const styleNode of styleNodes) {
    const typeface = getTextByPathList(styleNode, ['a:latin', 'attrs', 'typeface']) || getTextByPathList(styleNode, ['a:ea', 'attrs', 'typeface'])
    if (typeface) return typeface
  }

  return ''
}

function getColorFromNode(node, warpObj) {
  if (!node) return ''

  const fillType = getFillType(node)
  if (fillType === 'SOLID_FILL') {
    return getSolidFill(node['a:solidFill'], undefined, undefined, warpObj)
  }
  if (fillType === 'GRADIENT_FILL') {
    return getGradientFill(node['a:gradFill'], warpObj)
  }

  return ''
}

function getFontColorFromStyleNodes(styleNodes, warpObj) {
  for (const styleNode of styleNodes) {
    const color = getColorFromNode(styleNode, warpObj)
    if (color) return color
  }

  return ''
}

function getTextShadowFromStyleNodes(styleNodes, warpObj) {
  for (const styleNode of styleNodes) {
    const txtShadow = getTextByPathList(styleNode, ['a:effectLst', 'a:outerShdw'])
    if (!txtShadow) continue

    const shadow = getShadow(txtShadow, warpObj)
    if (shadow) return shadow
  }

  return null
}

export function getFontType(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, warpObj) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  let typeface = getFontTypeface(styleNodes)

  if (!typeface || typeface.startsWith('+')) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fontScheme'])

    if (fontSchemeNode) {
      if (typeface && typeface.startsWith('+')) {
        switch (typeface) {
          case '+mj-lt': 
            return getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface'])
          case '+mn-lt': 
            return getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
          case '+mj-ea': 
            return getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:ea', 'attrs', 'typeface'])
          case '+mn-ea': 
            return getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:ea', 'attrs', 'typeface'])
          default: 
            return typeface.replace(/^\+/, '')
        }
      }
    }

    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      typeface = getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:latin', 'attrs', 'typeface']) || getTextByPathList(fontSchemeNode, ['a:majorFont', 'a:ea', 'attrs', 'typeface'])
    }
    else {
      typeface = getTextByPathList(fontSchemeNode, ['a:minorFont', 'a:latin', 'attrs', 'typeface'])
    }
  }

  return typeface || ''
}

export function getFontColor(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, pFontStyle, warpObj) {
  const styleNodes = getBaseFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, lvl)
  let color = getFontColorFromStyleNodes(styleNodes, warpObj)

  if (!color) {
    if (pFontStyle) color = getSolidFill(pFontStyle, undefined, undefined, warpObj)
    if (!color) {
      const layoutFontStyle = getTextByPathList(slideLayoutSpNode, ['p:style', 'a:fontRef'])
      if (layoutFontStyle) color = getSolidFill(layoutFontStyle, undefined, undefined, warpObj)
    }
    if (!color) {
      const masterFontStyle = getTextByPathList(slideMasterSpNode, ['p:style', 'a:fontRef'])
      if (masterFontStyle) color = getSolidFill(masterFontStyle, undefined, undefined, warpObj)
    }
  }

  if (!color) {
    appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles)
    color = getFontColorFromStyleNodes(styleNodes, warpObj)
  }

  return color || ''
}

export function getFontSize(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, defaultTextStyle) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  appendDefaultTextStyleNodes(styleNodes, lvl, defaultTextStyle)
  const sz = getFontAttr(styleNodes, 'sz')
  let fontSize = sz ? parseInt(sz) / 100 : undefined

  if ((isNaN(fontSize) || !fontSize) && (type === 'dt' || type === 'sldNum')) fontSize = 12

  fontSize = (isNaN(fontSize) || !fontSize) ? 18 : fontSize

  return fontSize + 'pt'
}

export function getFontBold(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return isOoxmlTrue(getFontAttr(styleNodes, 'b')) ? 'bold' : ''
}

export function getFontItalic(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return isOoxmlTrue(getFontAttr(styleNodes, 'i')) ? 'italic' : ''
}

export function getFontDecoration(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return getFontAttr(styleNodes, 'u') === 'sng' ? 'underline' : ''
}

export function getFontDecorationLine(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  return getFontAttr(styleNodes, 'strike') === 'sngStrike' ? 'line-through' : ''
}

export function getFontSpace(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const spc = getFontAttr(styleNodes, 'spc')
  return (spc && parseInt(spc) !== 0) ? (parseInt(spc) / 100 + 'pt') : ''
}

export function getFontSubscript(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const baseline = getFontAttr(styleNodes, 'baseline')
  if (!baseline || parseInt(baseline) === 0) return ''
  return parseInt(baseline) > 0 ? 'super' : 'sub'
}

export function getFontShadow(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl, warpObj) {
  const styleNodes = getFontStyleNodes(node, pNode, textBodyNode, slideLayoutSpNode, slideMasterSpNode, type, slideMasterTextStyles, lvl)
  const shadow = getTextShadowFromStyleNodes(styleNodes, warpObj)
  if (shadow) {
    const { h, v, blur, color } = shadow
    if (!isNaN(v) && !isNaN(h)) {
      return h + 'pt ' + v + 'pt ' + (blur ? blur + 'pt' : '') + ' ' + color
    }
  }
  return ''
}
