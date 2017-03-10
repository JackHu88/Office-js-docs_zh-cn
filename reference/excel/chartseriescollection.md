# <a name="chartseriescollection-object-javascript-api-for-excel"></a>ChartSeriesCollection 对象 (Excel JavaScript API)

表示一组图表系列。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|count|int|返回集合中系列的数量。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[ChartSeries[]](chartseries.md)|chartSeries 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|返回集合中的系列数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|根据其在集合中的位置检索系列|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
返回集合中的系列数量。

#### <a name="syntax"></a>语法
```js
chartSeriesCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
根据其在集合中的位置检索系列

#### <a name="syntax"></a>语法
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[ChartSeries](chartseries.md)

#### <a name="examples"></a>示例

获取系列集合中第一个系列的名称。

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        console.log(seriesCollection.items[0].name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>属性访问示例
获取系列集合中系列的名称。

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < seriesCollection.items.length; i++)
        {
            console.log(seriesCollection.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

获取集合中图表系列的数目。

```js
Excel.run(function (ctx) { 
    var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
    seriesCollection.load('count');
    return ctx.sync().then(function() {
        console.log("series: Count= " + seriesCollection.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

