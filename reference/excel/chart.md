# <a name="chart-object-javascript-api-for-excel"></a>Chart 对象 (Excel JavaScript API)

表示工作簿中的 Chart 对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|height|double|表示图表对象的高度，以磅为单位。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|根据其在集合中的位置获取图表。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|left|double|从图表左侧到工作表原点的距离，以磅为单位。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|表示图表对象的名称。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|top|double|表示从对象左边界至第 1 行顶部（在工作表上）或图表区域顶部（在图表上）的距离，以磅值表示。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|width|double|表示图表对象的宽度，以磅为单位。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|axes|[ChartAxes](chartaxes.md)|表示图表坐标轴。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|表示图表上的数据标签。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|format|[ChartAreaFormat](chartareaformat.md)|封装图表区域的格式属性。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|legend|[ChartLegend](chartlegend.md)|表示图表的图例。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|series|[ChartSeriesCollection](chartseriescollection.md)|表示单个系列或图表中的系列集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|title|[ChartTitle](charttitle.md)|表示指定图表的标题，包括标题的文本、可见性、位置和格式。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|包含当前 chart 的 worksheet 对象。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|删除图表对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|[System.IO.Stream](system.io.stream.md)|通过缩放图表以适应指定尺寸，将图表呈现为 base64 编码的图像。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[setData(sourceData:Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|重置图表的源数据。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[setPosition(startCell:Range or string, endCell:Range or string](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|设置图表相对于工作表上单元格的位置。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="delete"></a>delete()
删除 chart 对象。

#### <a name="syntax"></a>语法
```js
chartObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getimageheight-number-width-number-fittingmode-string"></a>getImage(height: number, width: number, fittingMode: string)
通过缩放图表以适合指定的尺寸，将图表呈现为 base64 编码的图像。

#### <a name="syntax"></a>语法
```js
chartObject.getImage(height, width, fittingMode);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|height|number|可选。（可选）生成的图像所需的高度。|
|width|number|可选。（可选）生成的图像所需的宽度。|
|fittingMode|string|可选。（可选）该方法用于将图表缩放到指定尺寸（如果设置了高度和宽度）。"可能的值是：Fit、FitAndCenter、Fill|

#### <a name="returns"></a>返回
[System.IO.Stream](system.io.stream.md)

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var image = chart.getImage();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```





### <a name="setdatasourcedata-range-seriesby-string"></a>setData(sourceData:Range, seriesBy: string)
重置图表的源数据。

#### <a name="syntax"></a>语法
```js
chartObject.setData(sourceData, seriesBy);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|sourceData|Range|对应于源数据的 Range 对象。|
|seriesBy|string|可选。指定列或行在图表上用作数据系列的方式。可以是下列值之一：Auto（默认值）、Rows、Columns。可能的值是：Auto、Columns、Rows|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

将 `sourceData` 设置为“A1:B4”，将 `seriesBy` 设置为“Columns”

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    var sourceData = "A1:B4";
    chart.setData(sourceData, "Columns");
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="setpositionstartcell-range-or-string-endcell-range-or-string"></a>setPosition(startCell: Range or string, endCell: Range or string)
相对于工作表上的单元格放置图表。

#### <a name="syntax"></a>语法
```js
chartObject.setPosition(startCell, endCell);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|startCell|区域或字符串|起始单元格。这是图表将移动到的位置。起始单元格为左上角或右上角的单元格，具体取决于用户的从右到左显示设置。|
|endCell|区域或字符串|可选。（可选）结束单元格。如果指定，图表的宽度和高度将设置为完全覆盖此单元格/区域。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var rangeSelection = "A1:B4";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeSelection);
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", range, "auto");
    chart.width = 500;
    chart.height = 300;
    chart.setPosition("C2", null);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>属性访问示例

获取名为“Chart1”的图表

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.load('name');
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

更新包括重命名、定位和大小调整的图表。

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");    
    chart.name="New Name";
    chart.top = 100;
    chart.left = 100;
    chart.height = 200;
    chart.width = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

将图表重命名为新名称，将图表大小调整为高度和粗细均为 200 磅。将 Chart1 移动到距离顶部和左侧 100 磅。 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    chart.name="New Name";    
    chart.top = 100;
    chart.left = 100;
    chart.height =200;
    chart.width =200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

