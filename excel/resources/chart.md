# Chart 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

表示工作簿中的 chart 对象。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|height|double|表示 chart 对象的高度，以磅值表示。|
|left|double|从图表左侧到工作表原点的距离，以磅值表示。|
|name|string|表示 chart 对象的名称。|
|top|double|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|
|width|double|表示 chart 对象的宽度，以磅值表示。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|表示图表坐标轴。只读。|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|表示图表上的数据标签。只读。|
|format|[ChartAreaFormat](chartareaformat.md)|封装图表区域的格式属性。只读。|
|图例|[ChartLegend](chartlegend.md)|表示图表的图例。只读。|
|series|[ChartSeriesCollection](chartseriescollection.md)|表示单个系列或图表中的系列集合。只读。|
|title|[ChartTitle](charttitle.md)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。只读。|

## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|删除 chart 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[setData(sourceData:Range or string, seriesBy: string)](#setdatasourcedata-range-or-string-seriesby-string)|void|重置图表的源数据。|
|[setPosition(startCell:Range or string, endCell:Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|相对于工作表上的单元格放置图表。|

## 方法详细信息

### delete()
删除 chart 对象。

#### 语法
```js
chartObject.delete();
```

#### 参数
无

#### 返回
无效

#### 示例
```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void


### setData(sourceData:Range or string, seriesBy: string)
重置图表的源数据。

#### 语法
```js
chartObject.setData(sourceData, seriesBy);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|sourceData|Range or string|包含源数据的区域的地址或名称。如果使用地址或工作表范围内的名称，则必须包含工作表名称（例如"Sheet1!A5:B9"）。 |
|seriesBy|string|可选。指定列或行在图表上用作数据系列的方式。可以是下列值之一：Auto（默认值）、Rows、Columns。可能的值是：Auto、Columns、Rows|

#### 返回
无效

#### 示例

将 `sourceData` to 设置为“A1:B4”，将 `seriesBy` 设置为“Columns”

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	var sourceData = "A1:B4";
	chart.setData(sourceData, "Columns");
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### setPosition(startCell:Range or string, endCell:Range or string)
相对于工作表上的单元格放置图表。

#### 语法
```js
chartObject.setPosition(startCell, endCell);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|startCell|Range or string|起始单元格。这是图表将移动到的位置。起始单元格为左上角或右上角的单元格，具体取决于用户的从左到右显示设置。|
|endCell|Range or string|可选。结束单元格。如果指定，图表的宽度和高度将设置为完全覆盖此单元格/区域。|

#### 返回
无效

#### 示例


```js
Excel.run(function (ctx) { 
	var sheetName = "Charts";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
	chart.width = 500;
	chart.height = 300;
	chart.setPosition("C2", null);
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### 属性访问示例

获取名为“Chart1”的图表

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.load('name');
	return ctx.sync().then(function() {
			console.log(chart.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

更新包括重命名、定位和大小调整的图表。

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.name="New Name";
	chart.top = 100;
	chart.left = 100;
	chart.height = 200;
	chart.weight = 200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

将图表重命名为新名称；将图表大小调整为高度和粗细均为 200 磅。将 Chart1 移动到距离顶部和左侧 100 磅。 

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
	chart.name="New Name";	
	chart.top = 100;
	chart.left = 100;
	chart.height =200;
	chart.width =200;
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

