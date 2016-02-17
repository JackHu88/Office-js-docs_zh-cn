# ChartAxis 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

代表图表中的单个坐标轴。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|majorUnit|对象|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|
|maximum|对象|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
|minimum|对象|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|
|minorUnit|对象|表示两个次要刻度标记之间的间隔。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|格式|[ChartAxisFormat](chartaxisformat.md)|表示 chart 对象的格式，包括线条和字体格式。只读。|
|majorGridlines|[ChartGridlines](chartgridlines.md)|返回一个表示指定坐标轴的主要网格线的网格线对象。只读。|
|minorGridlines|[ChartGridlines](chartgridlines.md)|返回一个表示指定坐标轴的次要网格线的网格线对象。只读。|
|title|[ChartAxisTitle](chartaxistitle.md)|表示坐标轴标题。只读。|

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
从 Chart1 获取图表坐标轴的 `maximum`

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var axis = chart.axes.valueaxis;
	axis.load('maximum');
	return ctx.sync().then(function() {
			console.log(axis.maximum);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

设置数值轴的 `maximum`、`minimum`、`majorunit` 或 `minorunit`。 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.axes.valueaxis.maximum = 5;
	chart.axes.valueaxis.minimum = 0;
	chart.axes.valueaxis.majorunit = 1;
	chart.axes.valueaxis.minorunit = 0.2;
	return ctx.sync().then(function() {
			console.log("Axis Settings Changed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

