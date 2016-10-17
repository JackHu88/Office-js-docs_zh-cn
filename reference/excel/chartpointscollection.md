# <a name="chartpointscollection-object-(javascript-api-for-excel)"></a>ChartPointsCollection 对象（适用于 Excel 的 JavaScript API）

图表中某个系列的所有图表点的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|count|int|返回集合中图表点的数量。只读。|
|项目|[ChartPoint[]](chartpoint.md)|chartPoints 对象的集合。只读。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getItemAt(index: number)](#getitematindex-number)|[ChartPoint](chartpoint.md)|根据其在系列中的位置检索点。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
根据其在系列中的位置检索点。

#### <a name="syntax"></a>语法
```js
chartPointsCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[ChartPoint](chartpoint.md)

#### <a name="examples"></a>示例
设置点集合中第一个点的边框颜色

```js
Excel.run(function (ctx) { 
    var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
    points.getItemAt(0).format.fill.setSolidColor("#8FBC8F");
    return ctx.sync().then(function() {
        console.log("Point Border Color Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
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

获取点集合中点的名称

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
    pointsCollection.load('items');
    return ctx.sync().then(function() {
        console.log("Points Collection loaded");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

获取点的数目

```js
Excel.run(function (ctx) { 
    var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
    pointsCollection.load('count');
    return ctx.sync().then(function() {
        console.log("points: Count= " + pointsCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
