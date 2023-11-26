export function getPosition(slideSpNode, slideLayoutSpNode, slideMasterSpNode, factor) {
  let off

  if (slideSpNode) off = slideSpNode['a:off']['attrs']
  else if (slideLayoutSpNode) off = slideLayoutSpNode['a:off']['attrs']
  else if (slideMasterSpNode) off = slideMasterSpNode['a:off']['attrs']

  if (!off) return { top: 0, left: 0 }

  return {
    top: parseInt(off['y']) * factor,
    left: parseInt(off['x']) * factor,
  }
}

export function getSize(slideSpNode, slideLayoutSpNode, slideMasterSpNode, factor) {
  let ext

  if (slideSpNode) ext = slideSpNode['a:ext']['attrs']
  else if (slideLayoutSpNode) ext = slideLayoutSpNode['a:ext']['attrs']
  else if (slideMasterSpNode) ext = slideMasterSpNode['a:ext']['attrs']

  if (!ext) return { width: 0, height: 0 }

  return {
    width: parseInt(ext['cx']) * factor,
    height: parseInt(ext['cy']) * factor,
  }
}