# PPT坐标解析问题修复方案

## 问题描述
在解析PPT文件时，多个相同类型的占位符元素（如 `type="body"`）会出现坐标错误，所有元素都使用最后一个同类型元素的坐标。

## 根本原因
在 `indexNodes` 函数中，相同 `type` 的元素会相互覆盖：
```javascript
// 问题代码：后面的元素会覆盖前面的元素
typeTable["body"] = element1
typeTable["body"] = element2  // element1 被覆盖
```

## 解决方案
使用 `type` + `idx` 组合键作为唯一标识符，避免覆盖问题。

## 代码修改

### 1. 修改存储逻辑（`indexNodes` 函数）

**文件位置**: `src/pptxtojson.js` 第401行
```javascript
// 修改前
if (type) typeTable[type] = targetNodeItem

// 修改后
if (type) {
  const compositeKey = idx ? `${type}_${idx}` : type
  typeTable[compositeKey] = targetNodeItem
}
```

**文件位置**: `src/pptxtojson.js` 第417行
```javascript
// 修改前
if (type) typeTable[type] = targetNode

// 修改后
if (type) {
  const compositeKey = idx ? `${type}_${idx}` : type
  typeTable[compositeKey] = targetNode
}
```

### 2. 修改获取逻辑（`genShape` 函数）

**文件位置**: `src/pptxtojson.js` 第574-587行
```javascript
// 修改前
if (type) {
  if (idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
    slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
  } 
  else {
    slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
    slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
  }
}
else if (idx) {
  slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
  slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
}

// 修改后
if (type) {
  const compositeKey = idx ? `${type}_${idx}` : type
  slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][compositeKey]
  slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][compositeKey]
  
  // 向后兼容：如果组合键找不到，尝试单独的 type
  if (!slideLayoutSpNode && idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
    slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
  }
}
else if (idx) {
  slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
  slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
}
```

## 修改效果

### 修改前
```javascript
typeTable = {
  "body": lastBodyElement,  // 所有body类型元素都指向最后一个
  "title": lastTitleElement
}
```

### 修改后
```javascript
typeTable = {
  "body_47": bodyElement47,  // 每个元素都有唯一标识
  "body_48": bodyElement48,
  "body": bodyElementWithoutIdx,  // 向后兼容
  "title_1": titleElement1
}
```

## 方案优势
1. **解决覆盖问题**：每个占位符都有唯一的键值
2. **向后兼容**：没有 `idx` 的元素仍然使用原有逻辑
3. **性能优秀**：O(1) 哈希查找，无额外性能开销
4. **逻辑清晰**：组合键直接对应具体元素，易于理解和维护

## 部署方案：Patch Package（推荐）

### 1. 安装 patch-package
```bash
npm install --save-dev patch-package
```

### 2. 添加 postinstall 脚本
在 `package.json` 中添加：
```json
{
  "scripts": {
    "postinstall": "patch-package"
  }
}
```

### 3. 应用修改并生成 patch
1. 按照上述代码修改，直接修改 `node_modules/pptxtojson/src/pptxtojson.js`
2. 运行 `npx patch-package pptxtojson` 生成 patch 文件
3. 提交生成的 `patches/pptxtojson+1.5.0.patch` 文件到版本控制

### 4. 团队使用
其他开发者只需要：
```bash
npm install  # 自动应用 patch
```

## 已生成的 Patch 文件
项目中已包含 `patches/pptxtojson+1.5.0.patch` 文件，包含了所有必要的修改。

## 测试验证
修复后，每个占位符元素都会使用正确的坐标：
- `type="body", idx="47"` → 使用 `typeTable["body_47"]` 的坐标
- `type="body", idx="48"` → 使用 `typeTable["body_48"]` 的坐标
- 不再出现所有元素都使用同一个错误坐标的问题

## 优势
- ✅ **自动化**：团队成员 `npm install` 时自动应用
- ✅ **版本控制**：patch 文件可以提交到 git
- ✅ **升级安全**：升级库版本时会提示冲突
- ✅ **维护简单**：只需要维护小的 patch 文件
