import { getTextByPathList } from './utils'

export function findOMath(obj) {
  let results = []
  if (!obj || typeof obj !== 'object') return results
  if (obj['m:oMath']) results = results.concat(obj['m:oMath'])
  
  Object.values(obj).forEach(value => {
    if (Array.isArray(value) || typeof value === 'object') {
      results = results.concat(findOMath(value))
    }
  })
  return results
}

function orderedMathChildren(node) {
  const children = []
  for (const [key, value] of Object.entries(node || {})) {
    if (key === 'attrs' || key === 'value') continue
    for (const child of Array.isArray(value) ? value : [value]) {
      if (child && typeof child === 'object') children.push({ key, node: child })
    }
  }
  return children.sort((a, b) => {
    const aOrder = getTextByPathList(a.node, ['attrs', 'order']) || 0
    const bOrder = getTextByPathList(b.node, ['attrs', 'order']) || 0
    return aOrder - bOrder
  })
}

function mathTextToLatex(node) {
  const value = typeof node === 'string' ? node : node && node.value
  if (typeof value !== 'string' || value === '') return ''
  const text = latexFormart(value)
  const escapes = { '\\': '\\textbackslash{}', '{': '\\{', '}': '\\}', '$': '\\$', '&': '\\&', '#': '\\#', '%': '\\%', '_': '\\_', '^': '\\textasciicircum{}', '~': '\\textasciitilde{}' }
  return '\\text{' + text.replace(/[\\{}$&#%_^~]/g, character => escapes[character]) + '}'
}

export function composeMathTextBody(content) {
  const paragraphs = []
  function collectParagraphs(node) {
    if (!node || typeof node !== 'object') return
    for (const child of orderedMathChildren(node)) {
      if (child.key === 'a:p') paragraphs.push(child.node)
      else collectParagraphs(child.node)
    }
  }
  collectParagraphs(content)

  const lines = []
  let formulaCount = 0
  let includesText = false
  let singleFormula = ''
  for (const paragraph of paragraphs) {
    const paragraphAlign = getTextByPathList(paragraph, ['a:pPr', 'attrs', 'algn']) || 'ctr'
    let align = paragraphAlign
    const initialLineCount = lines.length
    let trailingBreak = false
    let displayEnded = false
    let current = []
    const flushLine = (force = false) => {
      if (current.length || force) lines.push({ latex: current.map(item => item.latex).join('') || '\\mathstrut', align })
      current = []
    }
    const appendFormula = node => {
      const latex = latexFormart(parseOMath(node))
      formulaCount++
      singleFormula = latex
      trailingBreak = false
      displayEnded = false
      if (current.length && current[current.length - 1].isFormula) current.push({ latex: '\\quad ', isFormula: false })
      current.push({ latex, isFormula: true })
    }
    const visit = (key, node) => {
      if (key === 'm:oMath') {
        appendFormula(node)
      }
      else if (key === 'a:br') {
        flushLine(!displayEnded)
        trailingBreak = !displayEnded
        displayEnded = false
      }
      else if (key === 'a:r' || key === 'a:fld') {
        const latex = mathTextToLatex(node['a:t'])
        if (latex) {
          includesText = true
          trailingBreak = false
          displayEnded = false
          current.push({ latex, isFormula: false })
        }
      }
      else if (key === 'm:oMathPara') {
        flushLine()
        const justification = getTextByPathList(node, ['m:oMathParaPr', 'm:jc', 'attrs', 'm:val'])
        if (justification === 'left') align = 'l'
        else if (justification === 'right') align = 'r'
        else if (justification) align = 'ctr'
        const equations = orderedMathChildren(node).filter(child => child.key === 'm:oMath')
        for (const equation of equations) {
          appendFormula(equation.node)
          flushLine()
        }
        align = paragraphAlign
        displayEnded = equations.length > 0
      }
      else if (!key.endsWith('Pr') && key !== 'a:lstStyle') {
        for (const child of orderedMathChildren(node)) visit(child.key, child.node)
      }
    }
    for (const child of orderedMathChildren(paragraph)) visit(child.key, child.node)
    flushLine(trailingBreak || lines.length === initialLineCount)
  }

  if (formulaCount === 0) {
    const equations = findOMath(content)
    const latex = equations.map(node => latexFormart(parseOMath(node)))
    formulaCount = equations.length
    if (formulaCount === 1) singleFormula = latex[0]
    for (const expression of latex) lines.push({ latex: expression, align: 'ctr' })
  }
  if (formulaCount <= 1) return { latex: singleFormula, formulaCount, includesText: false }
  if (lines.length === 1) return { latex: lines[0].latex, formulaCount, includesText }

  const leftAligned = lines.every(line => line.align === 'l')
  const rightAligned = lines.every(line => line.align === 'r')
  let latex
  if (leftAligned || rightAligned) {
    const alignment = leftAligned ? 'l' : 'r'
    latex = `\\begin{array}{${alignment}}` + lines.map(line => '\\displaystyle ' + line.latex).join(' \\\\ ') + '\\end{array}'
  }
  else latex = '\\begin{gathered}' + lines.map(line => line.latex).join(' \\\\ ') + '\\end{gathered}'
  return { latex, formulaCount, includesText }
}

export function parseFraction(fraction) {
  const numerator = parseOMath(fraction['m:num'])
  const denominator = parseOMath(fraction['m:den'])
  return `\\frac{${numerator}}{${denominator}}`
}
export function parseSuperscript(superscript) {
  const base = parseOMath(superscript['m:e'])
  const sup = parseOMath(superscript['m:sup'])
  return `${base}^{${sup}}`
}
export function parseSubscript(subscript) {
  const base = parseOMath(subscript['m:e'])
  const sub = parseOMath(subscript['m:sub'])
  return `${base}_{${sub}}`
}
export function parseSubSuperscript(subSuperscript) {
  const base = parseOMath(subSuperscript['m:e'])
  const sub = parseOMath(subSuperscript['m:sub'])
  const sup = parseOMath(subSuperscript['m:sup'])
  return `${base}_{${sub}}^{${sup}}`
}
export function parseRadical(radical) {
  const degree = parseOMath(radical['m:deg'])
  const expression = parseOMath(radical['m:e'])
  return degree ? `\\sqrt[${degree}]{${expression}}` : `\\sqrt{${expression}}`
}
export function parseMatrix(matrix) {
  const rowNodes = matrix['m:mr']
  const rows = Array.isArray(rowNodes) ? rowNodes : [rowNodes]
  const matrixRows = rows.map(row => {
    const cellNodes = row['m:e']
    const cells = Array.isArray(cellNodes) ? cellNodes : [cellNodes]
    return cells.map(element => parseOMath(element)).join(' & ')
  })
  return `\\begin{matrix} ${matrixRows.join(' \\\\ ')} \\end{matrix}`
}
export function parseNary(nary) {
  const op = getTextByPathList(nary, ['m:naryPr', 'm:chr', 'attrs', 'm:val']) || '∫'
  const sub = parseOMath(nary['m:sub'])
  const sup = parseOMath(nary['m:sup'])
  const e = parseOMath(nary['m:e'])
  return `${op}_{${sub}}^{${sup}}{${e}}`
}
export function parseLimit(limit, type) {
  const base = parseOMath(limit['m:e'])
  const lim = parseOMath(limit['m:lim'])
  return type === 'low' ? `${base}_{${lim}}` : `${base}^{${lim}}`
}
export function parseDelimiter(delimiter) {
  let left = getTextByPathList(delimiter, ['m:dPr', 'm:begChr', 'attrs', 'm:val'])
  let right = getTextByPathList(delimiter, ['m:dPr', 'm:endChr', 'attrs', 'm:val'])
  if (!left && !right) {
    left = '('
    right = ')'
  }
  if (left && right) {
    left = `\\left${left}`
    right = `\\right${right}`
  }
  const e = parseOMath(delimiter['m:e'])
  return `${left}${e}${right}`
}
export function parseFunction(func) {
  const name = parseOMath(func['m:fName'])
  const arg = parseOMath(func['m:e'])
  return `\\${name}{${arg}}`
}
export function parseGroupChr(groupChr) {
  const chr = getTextByPathList(groupChr, ['m:groupChrPr', 'm:chr', 'attrs', 'm:val'])
  const e = parseOMath(groupChr['m:e'])
  return `${chr}${e}${chr}`
}
export function parseEqArr(eqArr) {
  const equations = eqArr['m:e'].map((eq) => parseOMath(eq)).join(' \\\\ ')
  return `\\begin{cases} ${equations} \\end{cases}`
}
export function parseBar(bar) {
  const e = parseOMath(bar['m:e'])
  const pos = getTextByPathList(bar, ['m:barPr', 'm:pos', 'attrs', 'm:val'])
  return pos === 'top' ? `\\overline{${e}}` : `\\underline{${e}}`
}
export function parseAccent(accent) {
  const chr = getTextByPathList(accent, ['m:accPr', 'm:chr', 'attrs', 'm:val']) || '^'
  const e = parseOMath(accent['m:e'])
  switch (chr) {
    case '\u0301':
      return `\\acute{${e}}`
    case '\u0300':
      return `\\grave{${e}}`
    case '\u0302':
      return `\\hat{${e}}`
    case '\u0303':
      return `\\tilde{${e}}`
    case '\u0304':
      return `\\bar{${e}}`
    case '\u0306':
      return `\\breve{${e}}`
    case '\u0307':
      return `\\dot{${e}}`
    case '\u0308':
      return `\\ddot{${e}}`
    case '\u030A':
      return `\\mathring{${e}}`
    case '\u030B':
      return `\\H{${e}}`
    case '\u030C':
      return `\\check{${e}}`
    case '\u0327':
      return `\\c{${e}}`
    default:
      return `\\${chr}{${e}}`
  }
}
export function parseBox(box) {
  const e = parseOMath(box['m:e'])
  return `\\boxed{${e}}`
}


export function parseOMath(oMath) {
  if (!oMath) return ''

  if (Array.isArray(oMath)) {
    return oMath.map(item => parseOMath(item)).join('')
  }

  const oMathList = []
  const keys = Object.keys(oMath)
  for (const key of keys) {
    if (Array.isArray(oMath[key])) {
      oMathList.push(...oMath[key].map(item => ({ key, value: item })))
    }
    else oMathList.push({ key, value: oMath[key] })
  }

  oMathList.sort((a, b) => {
    let oA = 0
    if (a.key === 'm:r' && a.value && a.value['a:rPr']) oA = a.value['a:rPr']['attrs']['order']
    else if (a.value[`${a.key}Pr`] && a.value[`${a.key}Pr`]['m:ctrlPr'] && a.value[`${a.key}Pr`]['m:ctrlPr']['a:rPr']) {
      oA = a.value[`${a.key}Pr`] && a.value[`${a.key}Pr`]['m:ctrlPr'] && a.value[`${a.key}Pr`]['m:ctrlPr']['a:rPr'] && a.value[`${a.key}Pr`]['m:ctrlPr']['a:rPr']['attrs']['order']
    }
    let oB = 0
    if (b.key === 'm:r' && b.value && b.value['a:rPr']) oB = b.value['a:rPr']['attrs']['order']
    else if (b.value[`${b.key}Pr`] && b.value[`${b.key}Pr`]['m:ctrlPr'] && b.value[`${b.key}Pr`]['m:ctrlPr']['a:rPr']) {
      oB = b.value[`${b.key}Pr`] && b.value[`${b.key}Pr`]['m:ctrlPr'] && b.value[`${b.key}Pr`]['m:ctrlPr']['a:rPr'] && b.value[`${b.key}Pr`]['m:ctrlPr']['a:rPr']['attrs']['order']
    }
    return oA - oB
  })

  return oMathList.map(({ key, value }) => {
    if (key === 'm:f') return parseFraction(value)
    if (key === 'm:sSup') return parseSuperscript(value)
    if (key === 'm:sSub') return parseSubscript(value)
    if (key === 'm:sSubSup') return parseSubSuperscript(value)
    if (key === 'm:rad') return parseRadical(value)
    if (key === 'm:nary') return parseNary(value)
    if (key === 'm:limLow') return parseLimit(value, 'low')
    if (key === 'm:limUpp') return parseLimit(value, 'upp')
    if (key === 'm:d') return parseDelimiter(value)
    if (key === 'm:func') return parseFunction(value)
    if (key === 'm:groupChr') return parseGroupChr(value)
    if (key === 'm:eqArr') return parseEqArr(value)
    if (key === 'm:bar') return parseBar(value)
    if (key === 'm:acc') return parseAccent(value)
    if (key === 'm:borderBox') return parseBox(value)
    if (key === 'm:m') return parseMatrix(value)
    if (key === 'm:r') return parseOMath(value)
    if (key === 'm:t') return typeof value === 'string' ? value : value.value || ''
    return ''
  }).join('')
}

export function latexFormart(latex) {
  return latex.replaceAll(/&lt;/g, '<')
    .replaceAll(/&gt;/g, '>')
    .replaceAll(/&amp;/g, '&')
    .replaceAll(/&apos;/g, "'")
    .replaceAll(/&quot;/g, '"')
}