# <a name="charttitle-object-javascript-api-for-excel"></a>ChartTitle 对象 (Excel JavaScript API)

表示图表的图表标题对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|overlay|bool|表示图表标题是否覆盖图表的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|string|表示图表的标题文本。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|表示图表标题对象是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|format|[ChartTitleFormat](charttitleformat.md)|表示图表标题的格式，包括填充和字体格式。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
无


## <a name="method-details"></a>方法详细信息

### <a name="property-access-examples"></a>属性访问示例

从 Chart1 获取图表标题的 `text`。

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

var title = chart.title;
title.load('text');
return ctx.sync().then(function() {
        console.log(title.text);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
});
```

将图表标题的 `text` 设置为“My Chart”，并使其显示在图表之上，而不是覆盖。

```js
Excel.run(function (ctx) { 
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

return ctx.sync().then(function() {
        console.log("Char Title Changed");
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
});
```
