# ChartFont 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

此对象表示 chart 对象的字体属性（字体名称、字体大小、颜色等）。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|bold|bool|表示字体的加粗设置。|
|color|string|文本颜色的 HTML 颜色代码表示。例如，#FF0000 表示红色。|
|italic|bool|表示字体的斜体设置。|
|name|string|字体名称（例如"Calibri"）|
|size|double|字体大小（例如 11）|
|underline|string|应用于字体的下划线类型。可能的值是：None、Single。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
无


## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息

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

使用图表标题作为示例。

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = false;
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

设置图表标题格式为 Calibri，大小为 10，加粗和红色。 

```js
Excel.run(function (ctx) { 
	var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
	title.format.font.name = "Calibri";
	title.format.font.size = 12;
	title.format.font.color = "#FF0000";
	title.format.font.italic =  false;
	title.format.font.bold = true;
	title.format.font.underline = false;
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

