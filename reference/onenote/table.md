﻿# <a name="table-object-javascript-api-for-onenote"></a>表对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示 OneNote 页中的表。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|borderVisible|bool|获取或设置边框是否可见。如果为 true，则边框可见；如果为 false，则边框不可见。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-borderVisible)|
|columnCount|int|获取表格中的列数。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-columnCount)|
|id|string|获取表格的 ID。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-id)|
|rowCount|int|获取表格中的行数。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rowCount)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|获取包含表对象的段落对象。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-paragraph)|
|rows|[TableRowCollection](tablerowcollection.md)|获取所有表行。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rows)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|将列添加至表格结尾。指定的值在新列中设置。否则，列为空。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendColumn)|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|在表末尾添加一行。指定的值在新行中设置。否则，行为空。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendRow)|
|[clear()](#clear)|void|清除表内容。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-clear)|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|获取指定行和列的表单元格。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-getCell)|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|在表中给定索引处插入一列。指定的值在新列中设置。否则，列为空。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertColumn)|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|在表中给定索引处插入一行。指定的值在新行中设置。否则，行为空。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertRow)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|设置表中所有单元格的底纹色。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-setShadingColor)|

## <a name="method-details"></a>方法详细信息


### <a name="appendcolumnvalues-string"></a>appendColumn(values: string[])
将列添加至表格结尾。指定的值在新列中设置。否则该列为空。

#### <a name="syntax"></a>语法
```js
tableObject.appendColumn(values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|值|string[]|可选。可选。在新列中插入的字符串，指定为数组。表中的值不能多于行。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.appendColumn(["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="appendrowvalues-string"></a>appendRow(values: string[])
将行添加至表格结尾。指定的值在新行中设置。否则该行为空。

#### <a name="syntax"></a>语法
```js
tableObject.appendRow(values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|值|string[]|可选。可选。在新行中插入的字符串，指定为数组。表中的值不能多于列。|

#### <a name="returns"></a>返回
[TableRow](tablerow.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.appendRow(["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="clear"></a>clear()
清除表格的内容。

#### <a name="syntax"></a>语法
```js
tableObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="getcellrowindex-number-cellindex-number"></a>getCell(rowIndex: number, cellIndex: number)
获取位于指定行和列的表格单元格。

#### <a name="syntax"></a>语法
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|rowIndex|数字|行的索引。|
|cellIndex|数字|行中单元格的索引。|

#### <a name="returns"></a>返回
[TableCell](tablecell.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a cell in the second row and third column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(2 /*Row Index*/, 3 /*Column Index*/);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertcolumnindex-number-values-string"></a>insertColumn(index: number, values: string[])
在表中给定索引处插入一列。指定的值在新列中设置。否则该列为空。

#### <a name="syntax"></a>语法
```js
tableObject.insertColumn(index, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|数字|表格中插入列位置的索引。|
|值|string[]|可选。可选。在新列中插入的字符串，指定为数组。表中的值不能多于行。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, insert a column at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.insertColumn(2, ["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertrowindex-number-values-string"></a>insertRow(index: number, values: string[])
在表中给定索引处插入一行。指定的值在新行中设置。否则该行为空。

#### <a name="syntax"></a>语法
```js
tableObject.insertRow(index, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|数字|表格中插入行位置的索引。|
|值|string[]|可选。可选。在新行中插入的字符串，指定为数组。表中的值不能多于列。|

#### <a name="returns"></a>返回
[TableRow](tablerow.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, insert a row at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.insertRow(2, ["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
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
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void

### <a name="setshadingcolorcolorcode-string"></a>setShadingColor(colorCode: string)
设置表格中所有单元格的底纹色。

#### <a name="syntax"></a>语法
```js
tableObject.setShadingColor(colorCode);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|colorCode|string|要为单元格设置的颜色代码。/参数|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
**columnCount、rowCount、id**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // For each table, log properties.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table);
                return ctx.sync().then(function() {
                    console.log("Table Id: " + table.id);
                    console.log("Row Count: " + table.rowCount);
                    console.log("Column Count: " + table.columnCount);
                    return ctx.sync();
                });
            }
        }
    });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**paragraph、rows**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, log its paragraph id.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table, "paragraph/id, rows/id");
                return ctx.sync().then(function() {
                    console.log("Paragraph Id: " + table.paragraph.id);
                    var rows = table.rows;
                    
                    // for each rows in the table, log row index and id.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

