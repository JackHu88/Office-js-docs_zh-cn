# <a name="chartaxis-object-javascript-api-for-excel"></a>ChartAxis 对象 (Excel JavaScript API)

表示图表中的单个坐标轴。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|majorUnit|对象|表示两个主要刻度标记之间的间隔。可以设置为数字值或空字符串。返回的值始终为数字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|maximum|object|表示数值轴上的最大值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minimum|object|表示数值轴上的最小值。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minorUnit|对象|表示两个次要刻度标记之间的间隔。可以设置为数字值或空字符串（对于自动坐标轴值）。返回的值始终为数字。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|format|[ChartAxisFormat](chartaxisformat.md)|表示 chart 对象的格式，包括线条和字体格式。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|majorGridlines|[ChartGridlines](chartgridlines.md)|返回一个表示指定坐标轴的主要网格线的网格线对象。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|minorGridlines|[ChartGridlines](chartgridlines.md)|返回一个表示指定坐标轴的次要网格线的网格线对象。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartAxisTitle](chartaxistitle.md)|表示坐标轴标题。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
无


## <a name="method-details"></a>方法详细信息

### <a name="property-access-examples"></a>属性访问示例
从 Chart1 获取图表坐标轴的 `maximum`

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var axis = chart.axes.valueAxis;
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

设置数值轴的 `maximum`、`minimum`、`majorunit`、`minorunit`。 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.axes.valueAxis.maximum = 5;
    chart.axes.valueAxis.minimum = 0;
    chart.axes.valueAxis.majorUnit = 1;
    chart.axes.valueAxis.minorUnit = 0.2;
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
