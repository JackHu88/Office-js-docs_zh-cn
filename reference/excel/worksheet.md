# Worksheet 对象（适用于 Excel 的 JavaScript API）

Excel 工作表是由单元格组成的网格。它可以包含数据、表、图表等。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|id|string|返回用于唯一标识指定工作簿中工作表的值。 即使工作表被重命名或移动，标识符的值仍然相同。 值随所打开文件的每个会话更改。 只读。|
|name|string|工作表的显示名称。|
|position|int|工作表在工作簿中的位置，从零开始。|
|visibility|string|工作表的可见性。可能的值是：Visible、Hidden、VeryHidden。|

_请参阅属性访问[示例](#示例)。_

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|返回属于工作表的图表的集合。只读。|
|保护|[WorksheetProtection](worksheetprotection.md)|返回表工作表的工作表保护对象。只读。|
|表格|[TableCollection](tablecollection.md)|属于工作表的表的集合。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|在 Excel UI 中激活工作表。|
|[delete()](#delete)|void|从工作簿中删除工作表。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range 对象设置内联图片](range.md)|根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。|
|[getRange(address: string)](#getrangeaddress-string)|[Range 对象设置内联图片](range.md)|获取地址或名称指定的 range 对象。|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range 对象设置内联图片](range.md)|使用的区域是包含分配了值或格式化的任何单元格的最小区域。如果工作表为空，此函数将返回左上角的单元格。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### activate()
在 Excel UI 中激活工作表。

#### 语法
```js
worksheetObject.activate();
```

#### 参数
无

#### 返回
void

#### 示例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete()
从工作簿中删除工作表。

#### 语法
```js
worksheetObject.delete();
```

#### 参数
无

#### 返回
void

#### 示例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。

#### 语法
```js
worksheetObject.getCell(row, column);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|row|number|要检索的单元格的行号。从零开始编制索引。|
|column|number|要检索的单元格的列号。从零开始编制索引。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange(address: string)
获取地址或名称指定的 range 对象。

#### 语法
```js
worksheetObject.getRange(address);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|address|string|可选。区域的地址或名称。如果未指定，则返回整个工作表区域。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
此示例使用区域地址获取 range 对象。

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

此示例使用已命名的区域获取 range 对象。

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

### getUsedRange(valuesOnly: bool)
使用的区域是包含分配了值或格式化的任何单元格的最小区域。如果工作表为空，此函数将返回左上角的单元格。

#### 语法
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|valuesOnly|bool|可选。仅将具有值的单元格视为已使用的单元格（忽略格式）。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

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


### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void
### 属性访问示例

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

