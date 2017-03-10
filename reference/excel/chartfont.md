# <a name="chartfont-object-javascript-api-for-excel"></a>ChartFont 对象 (Excel JavaScript API)

此对象表示 Chart 对象的字体属性（字体名称、字号、颜色等）。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|bold|bool|表示字体的加粗状态。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|文本颜色的 HTML 颜色代码表示。例如，#FF0000 表示红色。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|表示字体的斜体状态。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|字体名称（例如"Calibri"）|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|size|double|字号（例如，11）|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|underline|string|应用于字体的下划线类型。可能的值是：None、Single。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法
无


## <a name="method-details"></a>方法详细信息

### <a name="property-access-examples"></a>属性访问示例

使用图表标题作为示例。

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = "None";
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

将图表标题设置为 Calbri，大小 10，粗体，红色。 

```js
Excel.run(function (ctx) { 
    var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;
    title.format.font.name = "Calibri";
    title.format.font.size = 12;
    title.format.font.color = "#FF0000";
    title.format.font.italic =  false;
    title.format.font.bold = true;
    title.format.font.underline = "None";
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
