import { getSolidFill } from './fill'

export function getShadow(node, warpObj, slideFactor) {
  const chdwClrNode = getSolidFill(node, undefined, undefined, warpObj)
  const outerShdwAttrs = node['attrs']
  const dir = (outerShdwAttrs['dir']) ? (parseInt(outerShdwAttrs['dir']) / 60000) : 0
  const dist = parseInt(outerShdwAttrs['dist']) * slideFactor
  const blurRad = outerShdwAttrs['blurRad'] ? (parseInt(outerShdwAttrs['blurRad']) * slideFactor) : ''
  const vx = dist * Math.sin(dir * Math.PI / 180)
  const hx = dist * Math.cos(dir * Math.PI / 180)

  return {
    h: hx,
    v: vx,
    blur: blurRad,
    color: '#' + chdwClrNode,
  }
}