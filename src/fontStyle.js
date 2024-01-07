import { getTextByPathList } from './utils'
import { getShadow } from './shadow'

export function getFontType(node, type, warpObj) {
  let typeface = getTextByPathList(node, ['a:rPr', 'a:latin', 'attrs', 'typeface'])

  if (!typeface) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:fontScheme'])

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

export function getFontColor(node) {
  const color = getTextByPathList(node, ['a:rPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])
  return color ? `#${color}` : ''
}

export function getFontSize(node, slideLayoutSpNode, type, slideMasterTextStyles, fontsizeFactor) {
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

  fontSize = (isNaN(fontSize) || !fontSize) ? 18 : fontSize

  return parseFloat((fontSize * fontsizeFactor).toFixed(2)) + (fontsizeFactor === 1 ? 'pt' : 'px')
}

export function getFontBold(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'b']) === '1' ? 'bold' : ''
}

export function getFontItalic(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'i']) === '1' ? 'italic' : ''
}

export function getFontDecoration(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'u']) === 'sng' ? 'underline' : ''
}

export function getFontDecorationLine(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'strike']) === 'sngStrike' ? 'line-through' : ''
}

export function getFontSpace(node, fontsizeFactor) {
  const spc = getTextByPathList(node, ['a:rPr', 'attrs', 'spc'])
  return spc ? (parseInt(spc) / 100 * fontsizeFactor + 'px') : ''
}

export function getFontSubscript(node) {
  const baseline = getTextByPathList(node, ['a:rPr', 'attrs', 'baseline'])
  if (!baseline) return ''
  return parseInt(baseline) > 0 ? 'super' : 'sub'
}

export function getFontShadow(node, warpObj) {
  const txtShadow = getTextByPathList(node, ['a:rPr', 'a:effectLst', 'a:outerShdw'])
  if (txtShadow) {
    const shadow = getShadow(txtShadow, warpObj)
    if (shadow) {
      const { h, v, blur, color } = shadow
      if (!isNaN(v) && !isNaN(h)) {
        return h + 'px ' + v + 'px ' + (blur ? blur + 'px' : '') + ' ' + color
      }
    }
  }
  return ''
}