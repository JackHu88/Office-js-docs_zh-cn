# <a name="worksheet-object-javascript-api-for-excel"></a>Worksheet 对象（适用于 Excel 的 JavaScript API）

Excel 工作表是由单元格组成的网格。它可以包含数据、表、图表等。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|id|string|返回用于唯一标识指定工作簿中工作表的值。即使工作表被重命名或移动，标识符的值仍然相同。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|工作表的显示名称。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|position|int|工作表在工作簿中的位置（以零为基）。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visibility|string|工作表的可见性。可能的值是：Visible、Hidden、VeryHidden。|[1.1：1.1 允许查看可见性，1.2 允许设置可见性。](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|charts|[ChartCollection](chartcollection.md)|返回属于工作表的图表的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|一组属于工作表的数据透视表。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorksheetProtection](worksheetprotection.md)|返回表工作表的工作表保护对象。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|属于工作表的表的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[activate()](#activate)|void|在 Excel UI 中激活工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|从工作簿中删除工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|获取按地址或名称指定的范围对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|使用的区域是包含分配了值或格式化的任何单元格的最小区域。如果工作表为空，此函数将返回左上角的单元格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="activate"></a>activate()
在 Excel UI 中激活工作表。

#### <a name="syntax"></a>语法
```js
worksheetObject.activate();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="delete"></a>delete()
从工作簿中删除工作表。

#### <a name="syntax"></a>语法
```js
worksheetObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。

#### <a name="syntax"></a>语法
```js
worksheetObject.getCell(row, column);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|row|number|要检索的单元格的行号。从零开始编制索引。|
|column|number|要检索的单元格的列号。从零开始编制索引。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeaddress-string"></a>getRange(address: string)
获取地址或名称指定的 range 对象。

#### <a name="syntax"></a>语法
```js
worksheetObject.getRange(address);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|address|string|可选。区域的地址或名称。如果未指定，则返回整个工作表区域。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
下面的示例使用区域地址获取 range 对象。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

下面的示例使用命名范围获取范围对象。

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getusedrangevaluesonly"></a>getUsedRange(valuesOnly)
使用的区域是包含分配了值或格式化的任何单元格的最小区域。如果工作表为空，此函数将返回左上角的单元格。

#### <a name="syntax"></a>语法
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|仅将有值的单元格视为已使用的单元格（忽略格式）。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


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

根据表名称获取工作表属性。

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

设置工作表位置。 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
