# <a name="chartcollection-object-(javascript-api-for-excel)"></a>ChartCollection 对象（适用于 Excel 的 JavaScript API）

工作表中的所有 chart 对象的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|count|int|返回工作表中的图表数。只读。|
|项目|[Chart[]](chart.md)|chart 对象的集合。只读。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[add(type: string, sourceData:Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|创建新图表。|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|根据其在集合中的位置获取图表。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="add(type:-string,-sourcedata:-range,-seriesby:-string)"></a>add(type: string, sourceData:Range, seriesBy: string)
创建新图表。

#### <a name="syntax"></a>语法
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|type|string|表示图表的类型。可能的值是：ColumnClustered、ColumnStacked、ColumnStacked100、BarClustered、BarStacked、BarStacked100、LineStacked、LineStacked100、LineMarkers、LineMarkersStacked、LineMarkersStacked100、PieOfPie 等。|
|sourceData|Range|包含源数据的 range 对象。|
|seriesBy|string|可选。指定列或行在图表上用作数据系列的方式。可能的值是：Auto、Columns、Rows|

#### <a name="returns"></a>返回
[Chart](chart.md)

#### <a name="examples"></a>示例

在工作表“Charts”上添加图表 `chartType`“ColumnClustered”，其中区域“A1:B4”中的 `sourceData` 和 `seriesBy` 设置为“auto”。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
    return ctx.sync().then(function() {
            console.log("New Chart Added");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitem(name:-string)"></a>getItem(name: string)
使用图表名称获取图表。如果存在多个名称相同的图表，将返回第一个图表。

#### <a name="syntax"></a>语法
```js
chartCollectionObject.getItem(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|name|string|要检索的图表的名称。|

#### <a name="returns"></a>返回
[Chart](chart.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var chartname = 'Chart1';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var chartId = 'SamplChartId';
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
    return ctx.sync().then(function() {
            console.log(chart.height);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
根据其在集合中的位置获取图表。

#### <a name="syntax"></a>语法
```js
chartCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[Chart](chart.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
    return ctx.sync().then(function() {
            console.log(chart.name);
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

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < charts.items.length; i++)
        {
            console.log(charts.items[i].name);
            console.log(charts.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

获取图表的数目

```js
Excel.run(function (ctx) { 
    var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
    charts.load('count');
    return ctx.sync().then(function() {
        console.log("charts: Count= " + charts.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

