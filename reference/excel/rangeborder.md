# <a name="rangeborder-object-javascript-api-for-excel"></a>RangeBorder 对象（适用于 Excel 的 JavaScript API）

代表对象的边框。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|color|string|表示窗体 #RRGGBB（例如“FFA500”）的边框线条颜色或作为已命名的 HTML 颜色（例如“orange”）的 HTML 颜色代码。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|表示边框标识符。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight、InsideVertical、InsideHorizontal、DiagonalDown、DiagonalUp。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sideIndex|string|指示边框的特定边的常量值。只读。可能的值是：EdgeTop、EdgeBottom、EdgeLeft、EdgeRight、InsideVertical、InsideHorizontal、DiagonalDown、DiagonalUp。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|线条样式的常量之一，指定边框的线条样式。可能的值是：None、Continuous、Dash、DashDot、DashDotDot、Dot、Double、SlantDashDot。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|weight|string|指定区域周围的边框的粗细。可能的值是：Hairline、Thin、Medium、Thick。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


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

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    var borders = range.format.borders;
    borders.load('items');
    return ctx.sync().then(function() {
        console.log(borders.count);
        for (var i = 0; i < borders.items.length; i++)
        {
            console.log(borders.items[i].sideIndex);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
下面的示例在区域周围添加网格边框。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
    range.format.borders.getItem('InsideVertical').style = 'Continuous';
    range.format.borders.getItem('EdgeBottom').style = 'Continuous';
    range.format.borders.getItem('EdgeLeft').style = 'Continuous';
    range.format.borders.getItem('EdgeRight').style = 'Continuous';
    range.format.borders.getItem('EdgeTop').style = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

