import { getTextByPathList } from './utils'
import { getSolidFill } from './fill'

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

  return fontSize * fontsizeFactor + 'px'
}

export function getFontBold(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'b']) === '1' ? 'bold' : ''
}

export function getFontItalic(node) {
  return getTextByPathList(node, ['a:rPr', 'attrs', 'i']) === '1' ? 'italic' : ''
}

export function getFontDecoration(node) {
  const underline = getTextByPathList(node, ['a:rPr', 'attrs', 'u']) === 'sng' ? 'underline' : ''
  const strike = getTextByPathList(node, ['a:rPr', 'attrs', 'strike']) === 'sngStrike' ? 'line-through' : ''

  if (!underline && !strike) return ''
  else if (underline && !strike) return underline
  else if (!underline && strike) return strike
  return `${underline} ${strike}`
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

export function getFontShadow(node, warpObj, slideFactor) {
  const txtShadow = getTextByPathList(node, ['a:rPr', 'a:effectLst', 'a:outerShdw'])
  if (txtShadow) {
    const shadowClr = getSolidFill(txtShadow, undefined, undefined, warpObj)
    const outerShdwAttrs = txtShadow['attrs']
    const dir = (outerShdwAttrs['dir']) ? (parseInt(outerShdwAttrs['dir']) / 60000) : 0
    const dist = parseInt(outerShdwAttrs['dist']) * slideFactor
    const blurRad = (outerShdwAttrs['blurRad']) ? (parseInt(outerShdwAttrs['blurRad']) * slideFactor + 'px') : ''
    const vx = dist * Math.sin(dir * Math.PI / 180)
    const hx = dist * Math.cos(dir * Math.PI / 180)
    if (!isNaN(vx) && !isNaN(hx)) {
      return hx + 'px ' + vx + 'px ' + blurRad + ' #' + shadowClr
    }
  }
  return ''
}