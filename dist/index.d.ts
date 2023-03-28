export interface Shape {
  type: 'shape'
  left: number
  top: number
  width: number
  height: number
  cx: number
  cy: number
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed'
  fillColor: string
  content: string
  isFlipV: boolean
  isFlipH: boolean
  rotate: number
  shapType: string
  id: string
  name: string
}

export interface Text {
  type: 'text'
  left: number
  top: number
  width: number
  height: number
  borderColor: string
  borderWidth: number
  borderType: 'solid' | 'dashed'
  fillColor: string
  isFlipV: boolean
  isFlipH: boolean
  rotate: number
  content: string
  id: string
  name: string
}

export interface Image {
  type: 'image'
  left: number
  top: number
  width: number
  height: number
  src: string
  rotate: number
}

export interface TableCell {
  text: string
  rowSpan?: number
  colSpan?: number
  vMerge?: boolean
  hMerge?: boolean
}
export interface Table {
  type: 'table'
  left: number
  top: number
  width: number
  height: number
  data: TableCell[][]
}

export type ChartType = 'lineChart' | 'stackedBarChart' | 'barChart' | 'pieChart' | 'pie3DChart' | 'stackedAreaChart' | 'areaChart'
export interface ChartValue {
  x: string
  y: number
}
export interface ChartXLabel {
  [key: string]: string
}
export interface ChartItem {
  key: string
  values: ChartValue[]
  xlabels: ChartXLabel
}
export type ScatterChartData = [number[], number[]]
export interface CommonChart {
  type: 'chart'
  left: number
  top: number
  width: number
  height: number
  data: ChartItem[],
  chartType: ChartType
}
export interface ScatterChart {
  type: 'chart'
  left: number
  top: number
  width: number
  height: number
  data: ScatterChartData,
  chartType: 'scatterChart'
}
export type Chart = CommonChart | ScatterChart

export interface Diagram {
  type: 'diagram'
  left: number
  top: number
  width: number
  height: number
}

export type BaseElement = Shape | Text | Image | Table | Chart | Diagram

export interface Group {
  type: 'group'
  left: number
  top: number
  width: number
  height: number
  elements: BaseElement[]
}
export type Element = BaseElement | Group

export interface SlideColorFill {
  type: 'color'
  value: string
}

export interface SlideImageFill {
  type: 'image'
  value: {
    picBase64: string
    opacity: number
  }
}

export type SlideFill = SlideColorFill | SlideImageFill

export interface Slide {
  fill: SlideFill
  elements: Element[]
}

export const parse: (file: ArrayBuffer) => {
  slides: Slide[]
  size: {
    width: number
    height: number
  }
}