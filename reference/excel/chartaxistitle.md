# <a name="chartaxistitle-object-(javascript-api-for-excel)"></a>ChartAxisTitle 对象（适用于 Excel 的 JavaScript API）

表示图表坐标轴的标题。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|text|string|表示坐标轴标题。|
|visible|bool|指定坐标轴标题的可见性的布尔值。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|format|[ChartAxisTitleFormat](chartaxistitleformat.md)|表示图表坐标轴标题的格式。只读。|

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
从 Chart1 的值坐标轴中获取图表坐标轴标题的 `text`。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var title = chart.axes.valueaxis.title;
    title.load('text');
    return ctx.sync().then(function() {
            console.log(title.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

添加“Values”作为值坐标轴的标题

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.axes.valueaxis.title.text = "Values";
    return ctx.sync().then(function() {
            console.log("Axis Title Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
