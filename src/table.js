import { getShapeFill, getSolidFill } from './fill'
import { getTextByPathList } from './utils'


export function getTableCellParams(tcNode, thisTblStyle, cellSource, warpObj) {
  const rowSpan = getTextByPathList(tcNode, ['attrs', 'rowSpan'])
  const colSpan = getTextByPathList(tcNode, ['attrs', 'gridSpan'])
  const vMerge = getTextByPathList(tcNode, ['attrs', 'vMerge'])
  const hMerge = getTextByPathList(tcNode, ['attrs', 'hMerge'])
  let fillColor
  let fontColor
  let fontBold

  const getCelFill = getTextByPathList(tcNode, ['a:tcPr'])
  if (getCelFill) {
    const cellObj = { 'p:spPr': getCelFill }
    fillColor = getShapeFill(cellObj, undefined, warpObj)
  }
  if (!fillColor) {
    let bgFillschemeClr
    if (cellSource) bgFillschemeClr = getTextByPathList(thisTblStyle, [cellSource, 'a:tcStyle', 'a:fill', 'a:solidFill'])
    if (bgFillschemeClr) {
      fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj)
    }
  }

  let rowTxtStyl
  if (cellSource) rowTxtStyl = getTextByPathList(thisTblStyle, [cellSource, 'a:tcTxStyle'])
  if (rowTxtStyl) {
    fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj)
    if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true
  }

  return {
    fillColor,
    fontColor,
    fontBold,
    rowSpan: rowSpan ? +rowSpan : undefined,
    colSpan: colSpan ? +colSpan : undefined,
    vMerge: vMerge ? +vMerge : undefined,
    hMerge: hMerge ? +hMerge : undefined,
  }
}

export function getTableRowParams(trNodes, i, tblStylAttrObj, thisTblStyle, warpObj) {
  let fillColor
  let fontColor
  let fontBold

  if (thisTblStyle && thisTblStyle['a:wholeTbl']) {
    const bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    if (bgFillschemeClr) {
      const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj)
      if (local_fillColor) fillColor = local_fillColor
    }
    const rowTxtStyl = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcTxStyle'])
    if (rowTxtStyl) {
      const local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj)
      if (local_fontColor) fontColor = local_fontColor
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true
    }
  }
  if (i === 0 && tblStylAttrObj['isFrstRowAttr'] === 1 && thisTblStyle) {
    const bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:firstRow', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    if (bgFillschemeClr) {
      const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj)
      if (local_fillColor) fillColor = local_fillColor
    }
    const rowTxtStyl = getTextByPathList(thisTblStyle, ['a:firstRow', 'a:tcTxStyle'])
    if (rowTxtStyl) {
      const local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj)
      if (local_fontColor) fontColor = local_fontColor
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true
    }
  }
  else if (i > 0 && tblStylAttrObj['isBandRowAttr'] === 1 && thisTblStyle) {
    fillColor = ''
    if ((i % 2) === 0 && thisTblStyle['a:band2H']) {
      const bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:band2H', 'a:tcStyle', 'a:fill', 'a:solidFill'])
      if (bgFillschemeClr) {
        const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj)
        if (local_fillColor) fillColor = local_fillColor
      }
      const rowTxtStyl = getTextByPathList(thisTblStyle, ['a:band2H', 'a:tcTxStyle'])
      if (rowTxtStyl) {
        const local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj)
        if (local_fontColor) fontColor = local_fontColor
      }
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true
    }
    if ((i % 2) !== 0 && thisTblStyle['a:band1H']) {
      const bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:band1H', 'a:tcStyle', 'a:fill', 'a:solidFill'])
      if (bgFillschemeClr) {
        const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj)
        if (local_fillColor) fillColor = local_fillColor
      }
      const rowTxtStyl = getTextByPathList(thisTblStyle, ['a:band1H', 'a:tcTxStyle'])
      if (rowTxtStyl) {
        const local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj)
        if (local_fontColor) fontColor = local_fontColor
        if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true
      }
    }
  }
  if (i === (trNodes.length - 1) && tblStylAttrObj['isLstRowAttr'] === 1 && thisTblStyle) {
    const bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:lastRow', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    if (bgFillschemeClr) {
      const local_fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj)
      if (local_fillColor) {
        fillColor = local_fillColor
      }
    }
    const rowTxtStyl = getTextByPathList(thisTblStyle, ['a:lastRow', 'a:tcTxStyle'])
    if (rowTxtStyl) {
      const local_fontColor = getSolidFill(rowTxtStyl, undefined, undefined, warpObj)
      if (local_fontColor) fontColor = local_fontColor
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true
    }
  }

  return {
    fillColor,
    fontColor,
    fontBold,
  }
}