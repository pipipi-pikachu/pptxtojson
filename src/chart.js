import { eachElement, getTextByPathList } from './utils'
import { applyTint } from './color'

function extractChartColors(serNode, warpObj) {
  if (!serNode) return []

  if (serNode.constructor !== Array) serNode = [serNode]
  const schemeClrs = []
  for (const node of serNode) {
    let schemeClr = getTextByPathList(node, ['c:spPr', 'a:solidFill', 'a:schemeClr'])
    if (!schemeClr) schemeClr = getTextByPathList(node, ['c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr'])
    if (!schemeClr) schemeClr = getTextByPathList(node, ['c:marker', 'c:spPr', 'a:ln', 'a:solidFill', 'a:schemeClr'])

    let clr = getTextByPathList(schemeClr, ['attrs', 'val'])
    if (clr) {
      clr = getTextByPathList(warpObj['themeContent'], ['a:theme', 'a:themeElements', 'a:clrScheme', `a:${clr}`, 'a:srgbClr', 'attrs', 'val'])
      const tint = getTextByPathList(schemeClr, ['a:tint', 'attrs', 'val']) / 100000
      if (clr && !isNaN(tint)) {
        clr = applyTint(clr, tint)
      }
    }
    else clr = getTextByPathList(node, ['c:spPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])

    if (clr) clr = '#' + clr
    schemeClrs.push(clr)
  }
  return schemeClrs
}

function getChartDataPoints(sourceNode) {
  const cache = getTextByPathList(sourceNode, ['c:strRef', 'c:strCache']) ||
    getTextByPathList(sourceNode, ['c:numRef', 'c:numCache']) ||
    getTextByPathList(sourceNode, ['c:strLit']) ||
    getTextByPathList(sourceNode, ['c:numLit'])

  return getTextByPathList(cache, ['c:pt'])
}

function extractChartData(serNode) {
  const dataMat = []
  if (!serNode) return dataMat

  eachElement(serNode, (innerNode, index) => {
    const dataRow = []
    const colName = getTextByPathList(innerNode, ['c:tx', 'c:strRef', 'c:strCache', 'c:pt', 'c:v']) || index

    const rowNames = {}
    eachElement(getChartDataPoints(innerNode['c:cat']), point => {
      rowNames[point['attrs']['idx']] = point['c:v']
      return ''
    })

    eachElement(getChartDataPoints(innerNode['c:val']), point => {
      dataRow.push({
        x: point['attrs']['idx'],
        y: parseFloat(point['c:v']),
      })
      return ''
    })

    dataMat.push({
      key: colName,
      values: dataRow,
      xlabels: rowNames,
    })
    return ''
  })

  return dataMat
}

function getNumericChartPoints(sourceNode) {
  const points = new Map()
  eachElement(getChartDataPoints(sourceNode), point => {
    const index = Number.parseInt(getTextByPathList(point, ['attrs', 'idx']), 10)
    const raw = point['c:v']
    const value = Number.parseFloat(typeof raw === 'object' && raw ? raw.value : raw)
    if (Number.isInteger(index) && index >= 0 && Number.isFinite(value)) points.set(index, value)
    return ''
  })
  return points
}

function getChartSeriesName(node, index) {
  const direct = getTextByPathList(node, ['c:tx', 'c:v'])
  if (typeof direct === 'string') return direct
  if (direct && typeof direct.value === 'string') return direct.value

  const points = getChartDataPoints(node['c:tx'])
  const first = Array.isArray(points) ? points[0] : points
  const name = getTextByPathList(first, ['c:v'])
  if (typeof name === 'string') return name
  if (name && typeof name.value === 'string') return name.value
  return String(index)
}

function extractScatterChartData(serNode, isBubble = false) {
  const series = []
  eachElement(serNode, (node, index) => {
    const xPoints = getNumericChartPoints(node['c:xVal'])
    const yPoints = getNumericChartPoints(node['c:yVal'])
    const sizePoints = isBubble ? getNumericChartPoints(node['c:bubbleSize']) : null
    const values = []
    for (const pointIndex of [...xPoints.keys()].sort((a, b) => a - b)) {
      if (!yPoints.has(pointIndex) || (isBubble && !sizePoints.has(pointIndex))) continue
      const point = { x: xPoints.get(pointIndex), y: yPoints.get(pointIndex) }
      if (isBubble) point.size = sizePoints.get(pointIndex)
      values.push(point)
    }
    series.push({ key: getChartSeriesName(node, index), values })
    return ''
  })
  return series
}

function getSingleChartInfo(key, node, warpObj) {
  switch (key) {
    case 'c:lineChart':
      return {
        type: 'lineChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        grouping: getTextByPathList(node, ['c:grouping', 'attrs', 'val']),
        marker: node['c:marker'] ? true : false,
      }
    case 'c:line3DChart':
      return {
        type: 'line3DChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        grouping: getTextByPathList(node, ['c:grouping', 'attrs', 'val']),
      }
    case 'c:barChart':
      return {
        type: 'barChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        grouping: getTextByPathList(node, ['c:grouping', 'attrs', 'val']),
        barDir: getTextByPathList(node, ['c:barDir', 'attrs', 'val']),
      }
    case 'c:bar3DChart':
      return {
        type: 'bar3DChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        grouping: getTextByPathList(node, ['c:grouping', 'attrs', 'val']),
        barDir: getTextByPathList(node, ['c:barDir', 'attrs', 'val']),
      }
    case 'c:pieChart':
      return {
        type: 'pieChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser']['c:dPt'], warpObj),
      }
    case 'c:pie3DChart':
      return {
        type: 'pie3DChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser']['c:dPt'], warpObj),
      }
    case 'c:doughnutChart':
      return {
        type: 'doughnutChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser']['c:dPt'], warpObj),
        holeSize: getTextByPathList(node, ['c:holeSize', 'attrs', 'val']),
      }
    case 'c:areaChart':
      return {
        type: 'areaChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        grouping: getTextByPathList(node, ['c:grouping', 'attrs', 'val']),
      }
    case 'c:area3DChart':
      return {
        type: 'area3DChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        grouping: getTextByPathList(node, ['c:grouping', 'attrs', 'val']),
      }
    case 'c:scatterChart':
      return {
        type: 'scatterChart',
        data: extractScatterChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        style: getTextByPathList(node, ['c:scatterStyle', 'attrs', 'val']),
      }
    case 'c:bubbleChart':
    {
      const chart = {
        type: 'bubbleChart',
        data: extractScatterChartData(node['c:ser'], true),
        colors: extractChartColors(node['c:ser'], warpObj),
      }
      const scale = getTextByPathList(node, ['c:bubbleScale', 'attrs', 'val'])
      if (scale !== undefined && scale !== null && scale !== '' && Number.isFinite(Number(scale))) chart.bubbleScale = Number(scale)
      const showNegative = getTextByPathList(node, ['c:showNegBubbles', 'attrs', 'val'])
      if (showNegative !== undefined) chart.showNegBubbles = showNegative === '1' || showNegative === 'true'
      const sizeRepresents = getTextByPathList(node, ['c:sizeRepresents', 'attrs', 'val'])
      if (sizeRepresents) chart.sizeRepresents = sizeRepresents
      return chart
    }
    case 'c:radarChart':
      return {
        type: 'radarChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
        style: getTextByPathList(node, ['c:radarStyle', 'attrs', 'val']),
      }
    case 'c:surfaceChart':
      return {
        type: 'surfaceChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
      }
    case 'c:surface3DChart':
      return {
        type: 'surface3DChart',
        data: extractChartData(node['c:ser']),
        colors: extractChartColors(node['c:ser'], warpObj),
      }
    case 'c:stockChart':
      return {
        type: 'stockChart',
        data: extractChartData(node['c:ser']),
        colors: [],
      }
    default:
      return null
  }
}

export function getChartInfo(plotArea, warpObj) {
  const nodes = []
  for (const [key, value] of Object.entries(plotArea || {})) {
    for (const node of Array.isArray(value) ? value : [value]) {
      if (node && node['c:ser']) nodes.push({ key, node })
    }
  }
  nodes.sort((a, b) => {
    const aOrder = getTextByPathList(a.node, ['attrs', 'order']) || 0
    const bOrder = getTextByPathList(b.node, ['attrs', 'order']) || 0
    return aOrder - bOrder
  })

  const charts = []
  for (const { key, node } of nodes) {
    const chart = getSingleChartInfo(key, node, warpObj)
    if (chart) charts.push({ chart, node })
  }
  if (!charts.length) return null
  if (charts.length === 1) return charts[0].chart

  const groups = charts.map(({ chart, node }) => {
    const { type, ...properties } = chart
    const axisIds = []
    eachElement(node['c:axId'], axis => {
      const id = getTextByPathList(axis, ['attrs', 'val'])
      if (id !== undefined) axisIds.push(String(id))
      return ''
    })
    const seriesIndices = []
    eachElement(node['c:ser'], (series, index) => {
      const value = Number.parseInt(getTextByPathList(series, ['c:idx', 'attrs', 'val']), 10)
      seriesIndices.push(Number.isFinite(value) ? value : index)
      return ''
    })
    return { chartType: type, ...properties, axisIds, seriesIndices }
  })
  return { type: 'comboChart', groups, colors: groups.flatMap(group => group.colors) }
}
