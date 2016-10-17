# <a name="rangeformat-object-(javascript-api-for-excel)"></a>RangeFormat 对象（适用于 Excel 的 JavaScript API）

一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|columnWidth|double|获取或设置区域内的所有列的宽度。如果列宽不统一，则返回 NULL。|
|horizontalAlignment|string|表示指定对象的水平对齐方式。可能的值是：General、Left、Center、Right、Fill、Justify、CenterAcrossSelection、Distributed。|
|rowHeight|double|获取或设置区域中所有行的高度。如果行高不统一，则返回 NULL。|
|verticalAlignment|string|表示指定对象的垂直对齐方式。可能的值是：Top、Center、Bottom、Justify、Distributed。|
|wrapText|bool|指示 Excel 文本控件被设置为对象中的自动换行。指示整个区域不使用统一自动换行设置的空值。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|borders|[RangeBorderCollection](rangebordercollection.md)|应用于所选的整个区域的 border 对象的集合。只读。|
|fill|[RangeFill](rangefill.md)|返回在整个区域内定义的 fill 对象。只读。|
|font|[RangeFont](rangefont.md)|返回在所选的整个区域内定义的 font 对象。只读。|
|protection|[FormatProtection](formatprotection.md)|返回某一区域的格式 protection 对象。只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[autofitColumns()](#autofitcolumns)|void|根据列中的当前数据，更改当前区域的列宽以达到最佳宽度。|
|[autofitRows()](#autofitrows)|void|根据列中的当前数据，更改当前区域的行高以达到最佳高度。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="autofitcolumns()"></a>autofitColumns()
根据列中的当前数据，更改当前区域的列宽以达到最佳宽度。

#### <a name="syntax"></a>语法
```js
rangeFormatObject.autofitColumns();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="autofitrows()"></a>autofitRows()
根据列中的当前数据，更改当前区域的行高以达到最佳高度。

#### <a name="syntax"></a>语法
```js
rangeFormatObject.autofitRows();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

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

本示例打印某一范围的所有格式属性。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load(["format/*", "format/fill", "format/borders", "format/font"]);
    return ctx.sync().then(function() {
        console.log(range.format.wrapText);
        console.log(range.format.fill.color);
        console.log(range.format.font.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

下面的示例设置字体名称、区域中的颜色填充和自动换行。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.wrapText = true;
    range.format.font.name = 'Times New Roman';
    range.format.fill.color = '0000FF';
    return ctx.sync(); 
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
    var rangeAddress = "F:G";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
    range.format.borders('InsideVertical').lineStyle = 'Continuous';
    range.format.borders('EdgeBottom').lineStyle = 'Continuous';
    range.format.borders('EdgeLeft').lineStyle = 'Continuous';
    range.format.borders('EdgeRight').lineStyle = 'Continuous';
    range.format.borders('EdgeTop').lineStyle = 'Continuous';
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
