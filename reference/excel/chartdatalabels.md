# <a name="chartdatalabels-object-javascript-api-for-excel"></a>ChartDataLabels 对象（适用于 Excel 的 JavaScript API）

表示图表点上的所有数据标签的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|position|string|表示数据标签的位置的 DataLabelPosition 值。可能的值是：None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|separator|string|表示用于图表中数据标签的分隔符的字符串。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBubbleSize|bool|表示数据标签气泡大小是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showCategoryName|bool|表示数据标签类别名称是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showLegendKey|bool|表示数据标签图例标示是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showPercentage|bool|表示数据标签百分比是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showSeriesName|bool|表示数据标签系列名称是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showValue|bool|表示数据标签值是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|表示图表数据标签的格式，包括填充和字体格式。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例

在数据标签中显示系列名称，并将数据标签的 `position` 设置为“top”。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.showValue = true;
    chart.datalabels.position = "top";
    chart.datalabels.showSeriesName = true;
    return ctx.sync().then(function() {
            console.log("Datalabels Shown");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
