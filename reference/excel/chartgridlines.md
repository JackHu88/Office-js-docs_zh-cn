# <a name="chartgridlines-object-javascript-api-for-excel"></a>ChartGridlines 对象（适用于 Excel 的 JavaScript API）

代表图表坐标轴的主要或次要网格线。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|visible|bool|表示坐标轴网格线是否可见的布尔值。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|format|[ChartGridlinesFormat](chartgridlinesformat.md)|表示图表网格线的格式。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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

获取 Chart1 的数值轴上主要网格线的 `visible`

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var majGridlines = chart.axes.valueaxis.majorGridlines;
    majGridlines.load('visible');
    return ctx.sync().then(function() {
            console.log(majGridlines.visible);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

设置为在 Chart1 的值坐标轴上显示主要网格线

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.axes.valueAxis.majorGridlines.visible = true;
    return ctx.sync().then(function() {
            console.log("Axis Gridlines Added ");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
