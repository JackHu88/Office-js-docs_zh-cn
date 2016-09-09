# ChartDataLabels 对象（适用于 Excel 的 JavaScript API）

表示图表点上的所有数据标签的集合。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|position|string|表示数据标签的位置的 DataLabelPosition 值。可能的值是：None、Center、InsideEnd、InsideBase、OutsideEnd、Left、Right、Top、Bottom、BestFit、Callout。只写。|
|Separator|string|表示用于图表中数据标签的分隔符的字符串。只写。|
|showBubbleSize|bool|表示数据标签气泡大小是否可见的布尔值。只写。|
|showCategoryName|bool|表示数据标签类别名称是否可见的布尔值。只写。|
|showLegendKey|bool|表示数据标签图例标示是否可见的布尔值。只写。|
|showPercentage|bool|表示数据标签百分比是否可见的布尔值。只写。|
|showSeriesName|bool|表示数据标签系列名称是否可见的布尔值。只写。|
|showValue|bool|表示数据标签值是否可见的布尔值。只写。|

_请参阅属性访问[示例](#示例)。_

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|格式|[ChartDataLabelFormat](chartdatalabelformat.md)|表示图表数据标签的格式，包括填充和字体格式。只读。|

## 方法

| 方法           | 返回类型    |说明|
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
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void
### 属性访问示例

使系列名称显示在数据标签中，并将数据标签的 `position` 设置为“top”。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.visible = true;
    chart.datalabels.position = "top";
    chart.datalabels.ShowSeriesName = true;
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
