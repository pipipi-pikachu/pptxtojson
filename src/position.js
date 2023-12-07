export function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode, factor) {
  let off

  if (slideSpNode) off = slideSpNode['a:off']['attrs']
  else if (slideLayoutSpNode) off = slideLayoutSpNode['a:off']['attrs']
  else if (slideMasterSpNode) off = slideMasterSpNode['a:off']['attrs']

  if (!off) return { top: 0, left: 0 }

  return {
    top: parseFloat((parseInt(off['y']) * factor).toFixed(2)),
    left: parseFloat((parseInt(off['x']) * factor).toFixed(2)),
  }
}

export function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode, factor) {
  let ext

  if (slideSpNode) ext = slideSpNode['a:ext']['attrs']
  else if (slideLayoutSpNode) ext = slideLayoutSpNode['a:ext']['attrs']
  else if (slideMasterSpNode) ext = slideMasterSpNode['a:ext']['attrs']

  if (!ext) return { width: 0, height: 0 }

  return {
    width: parseFloat((parseInt(ext['cx']) * factor).toFixed(2)),
    height: parseFloat((parseInt(ext['cy']) * factor).toFixed(2)),
  }
}