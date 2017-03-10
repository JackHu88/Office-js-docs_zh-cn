# <a name="chartlineformat-object-javascript-api-for-excel"></a>ChartLineFormat 对象 (Excel JavaScript API)

封装线条元素的格式选项。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|color|string|表示图表中的线条颜色的 HTML 颜色代码。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[clear()](#clear)|void|清除图表元素的线条格式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="clear"></a>clear()
清除图表元素的线条格式。

#### <a name="syntax"></a>语法
```js
chartLineFormatObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

清除“Chart1”图表上数值轴的主要网格线的线条格式。

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;    
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="property-access-examples"></a>属性访问示例

将值坐标轴上的图表主要网格线设置为红色。

```js
Excel.run(function (ctx) {
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueAxis.majorGridlines;
    gridlines.format.line.color = "#FF0000";
    return ctx.sync().then(function () {
        console.log("Chart Gridlines Color Updated");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
