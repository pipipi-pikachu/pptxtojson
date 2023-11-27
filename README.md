# 🎨 PPTX2JSON
这是一个可以将PPT幻灯片(.pptx)文件解析为 JSON 数据的库。

在线DEMO：https://pipipi-pikachu.github.io/pptx2json/


# 🔨安装
> npm install pptxtojson

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

# 🙏 感谢
> 本仓库主要参考了 [PPTX2HTML](https://github.com/g21589/PPTX2HTML) 、[PPTXjs](https://github.com/meshesha/PPTXjs) 的实现

# 📄 开源协议
AGPL-3.0 License | Copyright © 2020-PRESENT [pipipi-pikachu](https://github.com/pipipi-pikachu)