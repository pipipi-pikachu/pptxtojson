# 🎨 PPTX2JSON
这是一个可以将 .pptx 文件转为可读的 json 数据的 JavaScript 库。

在线DEMO：https://pipipi-pikachu.github.io/pptx2json/


# 🔨安装
```
npm install pptxtojson
```

# 💿用法
```html
<input type="file" accept="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
```

```js
import { parse } from 'pptxtojson'

const options = {
	slideFactor: 75 / 914400, // 幻灯片尺寸转换因子，默认 96 / 914400
	fontsizeFactor: 100 / 96, // 字号转换因子，默认 100 / 75
}

document.querySelector('input').addEventListener('change', evt => {
	const file = evt.target.files[0]
	
	const reader = new FileReader()
	reader.onload = async e => {
		const json = await parse(e.target.result, options)
		console.log(json)
	}
	reader.readAsArrayBuffer(file)
})
```

```js
// 输出示例
{
	"slides": {
		"fill": {
			"type": "color",
			"value": "#FF0000"
		},
		"elements": [
			// element data list
		],
	},
	"size": {
		"width": 1280,
		"height": 720
	}
}
```

# 📏 输出值单位
为了方便在web应用中使用，在默认情况下，所有输出的长度值单位都是px（像素），但这个值不一定是正确的，你可能需要根据你的设备情况适当调整 `slideFactor` 和 `fontsizeFactor` 参数来获取更准确的结果。或者将这个两个参数全部设置为1，这样输出的将会是原始数据，你可以在此基础上将原数据根据具体情况进行转换。

# 📕 功能支持

### 幻灯片尺寸
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| width                  | number                         | 宽度            
| height                 | number                         | 高度  

### 页面背景
| prop                   | type                            | 描述            
|------------------------|---------------------------------|---------------
| type                   | 'color' 丨 'image' 丨 'gradient' | 背景类型            
| value                  | SlideColorFill 丨 SlideImageFill 丨 SlideGradientFill| 背景值  

### 页内元素
#### 文字
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'text'                         | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| borderColor            | string                         | 边框颜色          
| borderWidth            | number                         | 边框宽度          
| borderType             | 'solid' 丨 'dashed' 丨 'dotted' | 边框类型          
| borderStrokeDasharray  | string                         | 非实线边框样式       
| shadow                 | Shadow                         | 阴影            
| fillColor              | string                         | 填充色           
| content                | string                         | 内容文字（HTML富文本） 
| isFlipV                | boolean                        | 垂直翻转          
| isFlipH                | boolean                        | 水平翻转          
| rotate                 | number                         | 旋转角度          
| vAlign                 | string                         | 垂直对齐方向        
| isVertical             | boolean                        | 是否为竖向文本        
| name                   | string                         | 元素名  

#### 图片
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'image'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| src                    | string                         | 图片地址（base64）    
| rotate                 | number                         | 旋转角度  

#### 形状
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'shape'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| borderColor            | string                         | 边框颜色          
| borderWidth            | number                         | 边框宽度          
| borderType             | 'solid' 丨 'dashed' 丨 'dotted' | 边框类型          
| borderStrokeDasharray  | string                         | 非实线边框样式       
| shadow                 | Shadow                         | 阴影            
| fillColor              | string                         | 填充色           
| content                | string                         | 内容文字（HTML富文本） 
| isFlipV                | boolean                        | 垂直翻转          
| isFlipH                | boolean                        | 水平翻转          
| rotate                 | number                         | 旋转角度          
| shapType               | string                         | 形状类型          
| vAlign                 | string                         | 垂直对齐方向        
| path                   | string                         | 路径（仅自定义形状存在）         
| name                   | string                         | 元素名   

#### 表格
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'table'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| data                   | TableCell[][]                  | 表格数据    
| themeColor             | string                         | 主题颜色  

#### 图表
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'chart'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| data                   | ChartItem[] 丨 ScatterChartData | 图表数据    
| chartType              | ChartType                      | 图表类型    
| barDir                 | 'bar' 丨 'col'                  | 柱状图方向    
| marker                 | boolean                        | 是否带数据标记    
| holeSize               | string                         | 环形图尺寸    
| grouping               | string                         | 分组模式    
| style                  | string                         | 图表样式 

#### 视频
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'video'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| blob                   | string                         | 视频blob    
| src                    | string                         | 视频src 

#### 音频
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'audio'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| blob                   | string                         | 音频blob   

#### Smart图
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'diagram'                      | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| elements               | (Shape 丨 Text)[]               | 子元素集合  

#### 多元素组合
| prop                   | type                           | 描述            
|------------------------|--------------------------------|---------------
| type                   | 'group'                        | 类型            
| left                   | number                         | 水平坐标          
| top                    | number                         | 垂直坐标          
| width                  | number                         | 宽度            
| height                 | number                         | 高度            
| elements               | Element[]                      | 子元素集合  

### 更多类型请参考 👇
[https://github.com/pipipi-pikachu/pptx2json/blob/master/dist/index.d.ts](https://github.com/pipipi-pikachu/pptx2json/blob/master/dist/index.d.ts)

# 🙏 感谢
> 本仓库主要参考了 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 、[PPTXjs](https://github.com/meshesha/PPTXjs) 的实现

# 📄 开源协议
MIT License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)