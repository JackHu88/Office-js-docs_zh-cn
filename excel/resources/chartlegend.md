# ChartLegend 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

表示图表中的图例。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|重叠|bool|表示图表图例是否应该与图表的主体重叠的布尔值。|
|position|string|表示图例在图表上的位置。可能的值是：Top、Bottom、Left、Right、Corner、Custom。|
|visible|bool|表示 ChartLegend 对象的可见性的布尔值。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|格式|[ChartLegendFormat](chartlegendformat.md)|表示图表图例的格式，包括填充和字体格式。只读。|

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

从 Chart1 获取图表图例的 `position`

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var legend = chart.legend;
	legend.load('position');
	return ctx.sync().then(function() {
			console.log(legend.position);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

设置为显示 Chart1 的图例，并将其显示在图表之上。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.legend.visible = true;
	chart.legend.position = "top"; 
	chart.legend.overlay = false; 
	return ctx.sync().then(function() {
			console.log("Legend Shown ");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
``` 

