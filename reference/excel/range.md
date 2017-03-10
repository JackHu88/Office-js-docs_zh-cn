# <a name="range-object-javascript-api-for-excel"></a>Range 对象 (Excel JavaScript API)

Range 表示一个或多个相邻的单元格，如单元格、行、列、单元格块等。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|地址|string|表示 A1 样式的区域引用。地址值将包含工作表引用（如 Sheet1!A1:B4）。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|addressLocal|string|以用户语言表示对指定区域的区域引用。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|cellCount|int|范围中的单元格数。如果单元格数超过 2^31-1 (2,147,483,647)，此 API 返回 -1。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|表示区域中的列总数。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|columnHidden|bool|表示当前范围的所有列均已隐藏。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|columnIndex|int|表示区域中第一个单元格的列编号。从零开始编制索引。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|表示采用 A1 表示法的公式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|表示采用 R1C1 表示法的公式。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|hidden|bool|表示当前区域中的所有单元格是否隐藏。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|获取一个值，该值指定|表示 Excel 中指定单元格的数字格式代码。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|返回区域中的总行数。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowHidden|bool|表示当前范围的所有行均已隐藏。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|rowIndex|int|返回区域中第一个单元格的行编号。从零开始编制索引。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|text|对象[][]|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|表示指定区域的原始值。返回的数据类型可能是字符串、数字或布尔值。包含一个将返回错误字符串的错误的单元格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|format|[RangeFormat](rangeformat.md)|返回一个格式对象，其中封装了区域的字体、填充、边框、对齐方式和其他属性。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[RangeSort](rangesort.md)|表示当前 range 的区域排序。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|包含当前区域的工作表。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[clear(applyTo: string)](#clearapplyto-string)|无效|清除范围值、格式、填充、边框等。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[delete(shift: string)](#deleteshift-string)|无效|删除与范围相关的单元格。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getBoundingRect(anotherRange:Range or string)](#getboundingrectanotherrange-range-or-string)|[区域](range.md)|获取包含指定区域的最小 range 对象。例如，“B2:C5”和“D10:E15”的 GetBoundingRect 为“B2:E16”。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。返回的单元格位于相对于区域左上角的单元格的位置。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|获取范围中包含的列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsAfter(count: number)](#getcolumnsaftercount-number)|[Range](range.md)|获取当前 Range 对象右侧的一定数量的列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getColumnsBefore(count: number)](#getcolumnsbeforecount-number)|[Range](range.md)|获取当前 Range 对象左侧的一定数量的列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|获取表示区域中整列的对象（例如，如果当前区域表示单元格块“B4:E11”，`getEntireColumn` 是表示列“B:E”的区域）。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getEntireRow()](#getentirerow)|[Range](range.md)|获取表示区域中整行的对象（例如，如果当前区域表示单元格块“B4:E11”，`GetEntireRow` 是表示行“4:11”的区域）。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersection(anotherRange:Range or string)](#getintersectionanotherrange-range-or-string)|[区域](range.md)|获取表示给定范围的矩形交集的范围对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getIntersectionOrNullObject(anotherRange:Range or string)](#getintersectionornullobjectanotherrange-range-or-string)|[Range](range.md)|获取表示给定范围的矩形交集的范围对象。如果找不到任何交集，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastCell()](#getlastcell)|[区域](range.md)|获取区域内的最后一个单元格。例如，“B2:D5”的最后一个单元格是“D5”。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastColumn()](#getlastcolumn)|[区域](range.md)|获取区域内的最后一列。例如，“B2:D5”的最后一列是“D2:D5”。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getLastRow()](#getlastrow)|[区域](range.md)|获取区域内的最后一行。例如，“B2:D5”的最后一行是“B5:D5”。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|获取表示与指定区域偏移的区域的对象。返回的区域的尺寸将与此区域一致。如果强制在工作表网格的边界之外生成区域，将引发错误。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getResizedRange(deltaRows: number, deltaColumns: number)](#getresizedrangedeltarows-number-deltacolumns-number)|[Range](range.md)|获取与当前范围对象类似的范围对象，但其右下角可通过一定数量的行和列进行展开（或合拢）。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|获取范围对象中包含的行。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsAbove(count: number)](#getrowsabovecount-number)|[Range](range.md)|获取当前 Range 对象上方的一定数量的行。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRowsBelow(count: number)](#getrowsbelowcount-number)|[Range](range.md)|获取当前 Range 对象下方的一定数量的行。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRange(valuesOnly: [ApiSet(Version)](#getusedrangevaluesonly-apisetversion)|[Range](range.md)|返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将引发 ItemNotFound 错误。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getUsedRangeOrNullObject(valuesOnly: bool)](#getusedrangeornullobjectvaluesonly-bool)|[Range](range.md)|返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getVisibleView()](#getvisibleview)|[RangeView](rangeview.md)|表示当前范围的可见行。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[merge(across: bool)](#mergeacross-bool)|void|将范围单元格合并到工作表的一个区域内。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[select()](#select)|void|在 Excel UI 中选择指定的范围。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[unmerge()](#unmerge)|void|将范围单元格取消合并为各个单元格。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="clearapplyto-string"></a>clear(applyTo: string)
清除区域值、格式、填充、边框等。

#### <a name="syntax"></a>语法
```js
rangeObject.clear(applyTo);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|applyTo|string|可选。确定清除操作的类型。可能的值是：`All`（默认选项）、`Formats`、`Contents` |

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

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


### <a name="deleteshift-string"></a>delete(shift: string)
删除与区域相关的单元格。

#### <a name="syntax"></a>语法
```js
rangeObject.delete(shift);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Shift|string|指定移动单元格的方式。可能的值是：Up、Left|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

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


### <a name="getboundingrectanotherrange-range-or-string"></a>getBoundingRect(anotherRange:Range or string)
获取包含指定区域的最小 range 对象。例如，“B2:C5”和“D10:E15”的 GetBoundingRect 为“B2:E16”。

#### <a name="syntax"></a>语法
```js
rangeObject.getBoundingRect(anotherRange);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|anotherRange|Range or string|range 对象或地址或区域名称。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getcellrow-number-column-number"></a>getCell(row: number, column: number)
根据行和列编号获取包含单个单元格的 range 对象。单元格可以位于父区域外部，只要其保持在工作表网格内即可。返回的单元格位于相对于区域左上角的单元格的位置。

#### <a name="syntax"></a>语法
```js
rangeObject.getCell(row, column);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
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
    var range = worksheet.getRange(rangeAddress);
    var cell = range.cell(0,0);
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


### <a name="getcolumncolumn-number"></a>getColumn(column: number)
获取区域中包含的列。

#### <a name="syntax"></a>语法
```js
rangeObject.getColumn(column);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|column|number|要检索的区域的列号。从零开始编制索引。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getcolumnsaftercount-number"></a>getColumnsAfter(count: number)
获取当前范围对象右侧的一定数量的列。

#### <a name="syntax"></a>语法
```js
rangeObject.getColumnsAfter(count);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|count|number|可选。生成的范围中要包含的列数。一般来说，使用正数可以在当前范围之外创建一个范围。也可以使用负数在当前范围之内创建一个范围。默认值为 1。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getcolumnsbeforecount-number"></a>getColumnsBefore(count: number)
获取当前范围对象左侧的一定数量的列。

#### <a name="syntax"></a>语法
```js
rangeObject.getColumnsBefore(count);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|count|number|可选。生成的范围中要包含的列数。一般来说，使用正数可以在当前范围之外创建一个范围。也可以使用负数在当前范围之内创建一个范围。默认值为 1。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getentirecolumn"></a>getEntireColumn()
获取表示区域中整列的对象（例如，如果当前区域表示单元格块“B4:E11”，`getEntireColumn` 是表示列“B:E”的区域）。

#### <a name="syntax"></a>语法
```js
rangeObject.getEntireColumn();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

注意：由于相关范围未绑定，因此范围的网格属性（values、numberFormat、formulas）包含 `null`。

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

### <a name="getentirerow"></a>getEntireRow()
获取表示区域中整行的对象（例如，如果当前区域表示单元格块“B4:E11”，`GetEntireRow` 是表示行“4:11”的区域）。

#### <a name="syntax"></a>语法
```js
rangeObject.getEntireRow();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
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
由于相关范围未绑定，因此范围的网格属性（values、numberFormat、formulas）包含 `null`。


### <a name="getintersectionanotherrange-range-or-string"></a>getIntersection(anotherRange:Range or string)
获取表示指定区域的矩形交集的 range 对象。

#### <a name="syntax"></a>语法
```js
rangeObject.getIntersection(anotherRange);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|anotherRange|Range or string|将用于确定区域交集的 range 对象或区域地址。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getintersectionornullobjectanotherrange-range-or-string"></a>getIntersectionOrNullObject(anotherRange:Range or string)
获取表示给定范围的矩形交集的范围对象。如果找不到任何交集，则此方法返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
rangeObject.getIntersectionOrNullObject(anotherRange);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|anotherRange|Range or string|将用于确定区域交集的 range 对象或区域地址。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getlastcell"></a>getLastCell()
获取区域内的最后一个单元格。例如，“B2:D5”的最后一个单元格是“D5”。

#### <a name="syntax"></a>语法
```js
rangeObject.getLastCell();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getlastcolumn"></a>getLastColumn()
获取区域内的最后一列。例如，“B2:D5”的最后一列是“D2:D5”。

#### <a name="syntax"></a>语法
```js
rangeObject.getLastColumn();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getlastrow"></a>getLastRow()
获取区域内的最后一行。例如，“B2:D5”的最后一行是“B5:D5”。

#### <a name="syntax"></a>语法
```js
rangeObject.getLastRow();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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



### <a name="getoffsetrangerowoffset-number-columnoffset-number"></a>getOffsetRange(rowOffset: number, columnOffset: number)
获取表示与指定区域偏移的区域的对象。返回的区域的尺寸将与此区域一致。如果强制在工作表网格的边界之外生成区域，将引发错误。

#### <a name="syntax"></a>语法
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|rowOffset|number|区域偏移的行数（正数、负数或 0）。正数表示向下偏移，负数表示向上偏移。|
|columnOffset|number|区域偏移的列数（正数、负数或 0）。正数表示向右偏移，负数表示向左偏移。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getresizedrangedeltarows-number-deltacolumns-number"></a>getResizedRange(deltaRows: number, deltaColumns: number)
获取与当前范围对象类似的范围对象，但其右下角可通过一定数量的行和列进行展开（或合拢）。

#### <a name="syntax"></a>语法
```js
rangeObject.getResizedRange(deltaRows, deltaColumns);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|deltaRows|number|相对于当前范围，展开右下角的行数。使用正数可展开范围，使用负数可合拢范围。|
|deltaColumns|number|相对于当前范围，右下角展开的列数。使用正数可展开范围，使用负数可合拢范围。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getrowrow-number"></a>getRow(row: number)
获取区域中包含的行。

#### <a name="syntax"></a>语法
```js
rangeObject.getRow(row);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|row|number|要检索的区域的行号。从零开始编制索引。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getrowsabovecount-number"></a>getRowsAbove(count: number)
获取当前范围对象上方的一定数量的行。

#### <a name="syntax"></a>语法
```js
rangeObject.getRowsAbove(count);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|count|number|可选。生成的范围中要包含的行数。一般来说，使用正数可以在当前范围之外创建一个范围。也可以使用负数在当前范围之内创建一个范围。默认值为 1。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getrowsbelowcount-number"></a>getRowsBelow(count: number)
获取当前范围对象下方的一定数量的行。

#### <a name="syntax"></a>语法
```js
rangeObject.getRowsBelow(count);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|count|number|可选。生成的范围中要包含的行数。一般来说，使用正数可以在当前范围之外创建一个范围。也可以使用负数在当前范围之内创建一个范围。默认值为 1。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getusedrangevaluesonly-apisetversion"></a>getUsedRange(valuesOnly: [ApiSet(Version)
返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将引发 ItemNotFound 错误。

#### <a name="syntax"></a>语法
```js
rangeObject.getUsedRange(valuesOnly);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|valuesOnly|[ApiSet(Version|仅将有值的单元格视为已使用的单元格。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="getusedrangeornullobjectvaluesonly-bool"></a>getUsedRangeOrNullObject(valuesOnly: bool)
返回指定 Range 对象的所用区域。如果区域内没有使用单元格，此函数将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
rangeObject.getUsedRangeOrNullObject(valuesOnly);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|valuesOnly|bool|可选。仅将有值的单元格视为使用的单元格。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="getvisibleview"></a>getVisibleView()
表示当前范围的可见行。

#### <a name="syntax"></a>语法
```js
rangeObject.getVisibleView();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[RangeView](rangeview.md)

### <a name="insertshift-string"></a>insert(shift: string)
将单个单元格或一系列单元格插入到工作表中取代此区域，并移动其他单元格以留出空间。在现在空白的空间返回新的 Range 对象。

#### <a name="syntax"></a>语法
```js
rangeObject.insert(shift);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Shift|string|指定移动单元格的方式。可能的值是：Down、Right|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

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


### <a name="mergeacross-bool"></a>merge(across: bool)
在工作表中，将 range 单元格合并到一个区域中。

#### <a name="syntax"></a>语法
```js
rangeObject.merge(across);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|横向|bool|可选。如果为 True，则将指定区域中每一行的单元格合并为一个单独的合并单元格。默认值是 false。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
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



#### <a name="examples"></a>示例
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


### <a name="select"></a>select()
在 Excel UI 中选择指定的区域。

#### <a name="syntax"></a>语法
```js
rangeObject.select();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

```js

Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "F5:F10"; 
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.select();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="unmerge"></a>unmerge()
将范围单元格取消合并为各个单元格。

#### <a name="syntax"></a>语法
```js
rangeObject.unmerge();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
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

### <a name="property-access-examples"></a>属性访问示例

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

下面的示例使用已命名的区域获取 range 对象。

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

下面的示例在包含 2x3 网格的网格上设置数字格式、值和公式。

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
获取包含区域的工作表。 

```js
/* This might be broken still - it was broken before because it 
    it was missing 'var', but might still be wrong because of
    getting information without loading properly. */
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    var range = namedItem.range;
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

