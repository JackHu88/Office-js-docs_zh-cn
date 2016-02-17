# WorksheetCollection 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

表示属于工作簿的 worksheet 对象的集合。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|Items|[Worksheet[]](worksheet.md)|worksheet 对象的集合。只读。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
无


## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|获取工作簿中当前处于活动状态的工作表。|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|使用其名称或 ID 获取 worksheet 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息

### add(name: string)
向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。

#### 语法
```js
worksheetCollectionObject.add(name);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|名称|string|可选。要添加的工作表的名称。如果指定，名称应唯一。如果未指定，Excel 将确定新工作表的名称。|

#### 返回
[Worksheet](worksheet.md)

#### 示例

```js
Excel.run(function (ctx) { 
	var wSheetName = 'Sample Name';
	var worksheet = ctx.workbook.worksheets.add(wSheetName);
	worksheet.load('name');
	return ctx.sync().then(function() {
		console.log(worksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getActiveWorksheet()
获取工作簿中当前处于活动状态的工作表。

#### 语法
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### 参数
无

#### 返回
[Worksheet](worksheet.md)

#### 示例

```js
Excel.run(function (ctx) {  
	var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
	activeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(activeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(key: string)
使用其名称或 ID 获取 worksheet 对象。

#### 语法
```js
worksheetCollectionObject.getItem(key);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|Key|string|工作表的名称或 ID。|

#### 返回
[Worksheet](worksheet.md)
### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void
### 属性访问示例
```js
Excel.run(function (ctx) { 
	var worksheets = ctx.workbook.worksheets;
	worksheets.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < worksheets.items.length; i++)
		{
			console.log(worksheets.items[i].name);
			console.log(worksheets.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
