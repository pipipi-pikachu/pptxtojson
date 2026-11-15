import { RATIO_EMUs_Points } from './constants'
import { getTextByPathList, numberToFixed } from './utils'

export function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
  const transform = slideSpNode || slideLayoutSpNode || slideMasterSpNode
  const off = getTextByPathList(transform, ['a:off', 'attrs'])

  if (!off) return { top: 0, left: 0 }

  return {
    top: numberToFixed(parseInt(off['y']) * RATIO_EMUs_Points),
    left: numberToFixed(parseInt(off['x']) * RATIO_EMUs_Points),
  }
}

export function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode) {
  const transform = slideSpNode || slideLayoutSpNode || slideMasterSpNode
  const ext = getTextByPathList(transform, ['a:ext', 'attrs'])

  if (!ext) return { width: 0, height: 0 }

  return {
    width: numberToFixed(parseInt(ext['cx']) * RATIO_EMUs_Points),
    height: numberToFixed(parseInt(ext['cy']) * RATIO_EMUs_Points),
  }
}