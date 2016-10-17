# <a name="chartseries-object-(javascript-api-for-excel)"></a>ChartSeries 对象（适用于 Excel 的 JavaScript API）

代表图表上的系列。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|name|string|表示图表中某个系列的名称。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|format|[ChartSeriesFormat](chartseriesformat.md)|表示图表系列的格式，包括填充和线条格式。只读。|
|points|[ChartPointsCollection](chartpointscollection.md)|表示系列中所有数据点的集合。只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="load(param:-object)"></a>load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例

将 Chart1 的第一个系列重命名为“New Series Name”。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.series.getItemAt(0).name = "New Series Name";
    return ctx.sync().then(function() {
            console.log("Series1 Renamed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
