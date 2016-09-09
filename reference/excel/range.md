# Range 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

区域表示一个或多个相邻的单元格，例如单元格、行、列、单元格块等。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|address|string|表示 A1 样式的区域引用。地址值将包含工作表引用（如 Sheet1!A1:B4）。只读。|
|addressLocal|string|以用户语言表示对指定区域的区域引用。只读。|
|cellCount|int|区域中的单元格数目。只读。|
|columnCount|int|表示区域中的列总数。只读。|
|columnHidden|bool|表示当前区域中的所有列是否隐藏。|
|columnIndex|int|表示区域中第一个单元格的列编号。从零开始编制索引。只读。|
|formulas|object[]|表示采用 A1 样式表示法的公式。|
|formulasLocal|object[][]|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|
|formulasR1C1|object[][]|表示采用 R1C1 样式表示法的公式。|
|hidden|bool|表示当前区域中的所有单元格是否隐藏。只读。|
|numberFormat|object[][]|表示指定单元格的数字格式代码。|
|rowCount|int|返回区域中的总行数。只读。|
|rowHidden|bool|表示当前区域中的所有行是否隐藏。|
|rowIndex|int|返回区域中第一个单元格的行编号。从零开始编制索引。只读。|
|text|object[][]|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|
|valueTypes|string|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|
|values|object[][]|表示指定区域的原始值。返回的数据类型可能是字符串、数字或布尔值。包含错误的单元格将返回错误的字符串。|

_请参阅属性访问[示例](#示例)。_

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|格式|[RangeFormat](rangeformat.md)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。只读。|
|排序|[RangeSort](rangesort.md)|表示区域的排序配置。只读。|
|工作表|[工作表](worksheet.md)|包含当前区域的工作表。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|清除区域值、格式、填充、边框等。|
|[delete(shift: string)](#deleteshift-string)|void|删除与区域相关的单元格。|
|[getBoundingRect(anotherRange:Range or string)](#getboundingrectanotherrange-range-or-string)|[Range 对象设置内联图片](range.md)|获取包含指定区域的最小 range 对象。例如，“B2:C5”和“D10:E15”的 getBoundingRect 为“B2:E15”。|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range 对象设置内联图片](range.md)|根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。返回的单元格位于相对于区域左上角的单元格的位置。|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range 对象设置内联图片](range.md)|获取区域中包含的列。|
|[getEntireColumn()](#getentirecolumn)|[Range 对象设置内联图片](range.md)|获取表示区域整列的对象。|
|[getEntireRow()](#getentirerow)|[Range 对象设置内联图片](range.md)|获取表示区域整行的对象。|
|[getIntersection(anotherRange:Range or string)](#getintersectionanotherrange-range-or-string)|[Range 对象设置内联图片](range.md)|获取表示指定区域的矩形交集的 range 对象。|
|[getLastCell()](#getlastcell)|[Range 对象设置内联图片](range.md)|获取区域内的最后一个单元格。例如，“B2:D5”的最后一个单元格是“D5”。|
|[getLastColumn()](#getlastcolumn)|[Range 对象设置内联图片](range.md)|获取区域内的最后一列。例如，“B2:D5”的最后一列是“D2:D5”。|
|[getLastRow()](#getlastrow)|[Range 对象设置内联图片](range.md)|获取区域内的最后一行。例如，“B2:D5”的最后一行是“B5:D5”。|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range 对象设置内联图片](range.md)|获取表示与指定区域偏移的区域的对象。返回的区域的尺寸将与该区域匹配。如果强制使生成的区域位于工作表网格的边界之外，则会引发异常。|
|[getRow(row: number)](#getrowrow-number)|[Range 对象设置内联图片](range.md)|获取区域中包含的行。|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range 对象设置内联图片](range.md)|返回 range 对象的所用子区域。|
|[insert(shift: string)](#insertshift-string)|[Range 对象设置内联图片](range.md)|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[merge(across: bool)](#mergeacross-bool)|void|在工作表中，将 range 单元格合并到一个区域中。|
|[select()](#select)|void|在 Excel UI 中选择指定的区域。|
|[unmerge()](#unmerge)|void|将 range 单元格拆分为单个单元格。|

## 方法详细信息


### clear(applyTo: string)
清除区域值、格式、填充、边框等。

#### 语法
```js
rangeObject.clear(applyTo);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|applyTo|string|可选。确定清除操作的类型。可能的值是：`All`（默认选项）、`Formats`、`Contents`|

#### 返回
void

#### 示例

以下示例将清除区域的格式和内容。 

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete(shift: string)
删除与区域相关的单元格。

#### 语法
```js
rangeObject.delete(shift);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Shift|string|指定移动单元格的方式。可能的值是：Up、Left|

#### 返回
void

#### 示例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getBoundingRect(anotherRange:Range or string)
获取包含指定区域的最小 range 对象。例如，“B2:C5”和“D10:E15”的 GetBoundingRect 为“B2:E15”。

#### 语法
```js
rangeObject.getBoundingRect(anotherRange);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|anotherRange|Range or string|range 对象或地址或区域名称。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:G6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var range = range.getBoundingRect("G4:H8");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // Prints Sheet1!D4:H8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。返回的单元格位于相对于区域左上角的单元格的位置。

#### 语法
```js
rangeObject.getCell(row, column);
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
    var range = worksheet.getRange(rangeAddress);
    var cell = range.getCell(0,0);
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


### getColumn(column: number)
获取区域中包含的列。

#### 语法
```js
rangeObject.getColumn(column);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|column|number|要检索的区域的列号。从零开始编制索引。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet19";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!B1:B8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getEntireColumn()
获取表示区域整列的对象。

#### 语法
```js
rangeObject.getEntireColumn();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

注意：由于相关区域是无限的，区域的网格属性（值、numberFormat、公式）包含 `null`。

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeEC = range.getEntireColumn();
    rangeEC.load('address');
    return ctx.sync().then(function() {
        console.log(rangeEC.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getEntireRow()
获取表示区域整行的对象。

#### 语法
```js
rangeObject.getEntireRow();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "D:F"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeER = range.getEntireRow();
    rangeER.load('address');
    return ctx.sync().then(function() {
        console.log(rangeER.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
由于相关区域是无限的，区域的网格属性（值、numberFormat、公式）包含 `null`。

### getIntersection(anotherRange:Range or string)
获取表示指定区域的矩形交集的 range 对象。

#### 语法
```js
rangeObject.getIntersection(anotherRange);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|anotherRange|Range or string|将用于确定区域交集的 range 对象或区域地址。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!D4:F6
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastCell()
获取区域内的最后一个单元格。例如，“B2:D5”的最后一个单元格是“D5”。

#### 语法
```js
rangeObject.getLastCell();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastColumn()
获取区域内的最后一列。例如，“B2:D5”的最后一列是“D2:D5”。

#### 语法
```js
rangeObject.getLastColumn();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!F1:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getLastRow()
获取区域内的最后一行。例如，“B2:D5”的最后一行是“B5:D5”。

#### 语法
```js
rangeObject.getLastRow();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A8:F8
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```



### getOffsetRange(rowOffset: number, columnOffset: number)
获取表示与指定区域偏移的区域的对象。返回的区域的尺寸将与该区域匹配。如果强制使生成的区域位于工作表网格的边界之外，则会引发异常。

#### 语法
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|rowOffset|number|区域偏移的行数（正数、负数或 0）。正数表示向下偏移，负数表示向上偏移。|
|columnOffset|number|区域偏移的列数（正数、负数或 0）。正数表示向右偏移，负数表示向左偏移。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D4:F6";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!H3:K5
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRow(row: number)
获取区域中包含的行。

#### 语法
```js
rangeObject.getRow(row);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|row|number|要检索的区域的行号。从零开始编制索引。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address); // prints Sheet1!A2:F2
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getUsedRange(valuesOnly: bool)
返回指定 range 对象的所用区域。

#### 语法
```js
rangeObject.getUsedRange(valuesOnly);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|valuesOnly|bool|可选。为 true 时，仅当前具有值的单元格被视为已使用的单元格。默认值为 false，将曾经具有值的所有单元格计入已使用的单元格。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "D:F";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var rangeUR = range.getUsedRange();
    rangeUR.load('address');
    return ctx.sync().then(function() {
        console.log(rangeUR.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### insert(shift: string)
将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。

#### 语法
```js
rangeObject.insert(shift);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Shift|string|指定移动单元格的方式。可能的值是：Down、Right|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

```js
    
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.insert();
    return ctx.sync(); 
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

### merge(across: bool)
在工作表中，将 range 单元格合并到一个区域中。

#### 语法
```js
rangeObject.merge(across);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|横向|bool|可选。如果为 True，则将指定区域中每一行的单元格合并为一个单独的合并单元格。默认值是 false。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.merge(true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### select()
在 Excel UI 中选择指定的区域。

#### 语法
```js
rangeObject.select();
```

#### 参数
无

#### 返回
void

#### 示例

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### unmerge()
将已合并的 range 单元格拆分为单个单元格。

#### 语法
```js
rangeObject.unmerge();
```

#### 参数
无

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.unmerge();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### 属性访问示例

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
    var rangeName = 'MyRange';
    var range = ctx.workbook.names.getItem(rangeName).range;
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

下面的示例在包含 2x3 网格的网格上设置 numberFormat、值和公式。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulas= formulas;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
下面的示例与上述示例相同，只是它的公式使用 R1C1 表示法。

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "F5:G7";
    var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
    var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
    var formulasR1C1 = [[null,null], [null,null], [null,"=R[-1]C-R[-2]C"]];
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.numberFormat = numberFormat;
    range.values = values;
    range.formulasR1C1= formulasR1C1;
    range.load('text');
    return ctx.sync().then(function() {
        console.log(range.text);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
获取包含区域的工作表。 

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    range = namedItem.range;
    var rangeWorksheet = range.worksheet;
    rangeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(rangeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

