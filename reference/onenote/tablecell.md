# <a name="tablecell-object-(javascript-api-for-onenote)"></a>TableCell 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示 OneNote 表中的一个单元格。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|cellIndex|int|获取单元格行中的单元格索引。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-cellIndex)|
|id|字符串|获取单元格的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-id)|
|rowIndex|int|获取表中单元格行的索引。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-rowIndex)|
|shadingColor|string|获取并设置单元格的底纹色|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-shadingColor)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|获取 TableCell 中 Paragraph 对象的集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-paragraphs)|
|parentRow|[TableRow](tablerow.md)|获取单元格的父行。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-parentRow)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|将指定的 HTML 添加到 TableCell 的底部。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|将指定图像添加到表格单元格中。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|将指定文本添加到表格单元格中。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|将带有指定行数和列数的表格添加到表格单元格中。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-appendTable)|
|[clear()](#clear)|void|清除单元格的内容。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-clear)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCell-load)|

## <a name="method-details"></a>方法详细信息


### <a name="appendhtml(html:-string)"></a>appendHtml(html: string)
将指定的 HTML 添加到 TableCell 的底部。

#### <a name="syntax"></a>语法
```js
tableCellObject.appendHtml(html);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Html|字符串|要追加的 HTML 字符串。请查看 OneNote 外接程序 JavaScript API [支持的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

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
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                cell.appendHtml("<p>Hello</p>");
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


### appendImage(base64EncodedImage: string, width: double, height: double)
Adds the specified image to table cell.

#### Syntax
```js
tableCellObject.appendImage(base64EncodedImage, width, height);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64EncodedImage|字符串|要追加的 HTML 字符串。|
|宽度|double|可选。以磅为单位的宽度。默认值为 null，将考虑图像宽度。|
|高度|double|可选。以磅为单位的高度。默认值为 null，将考虑图像高度。|

#### <a name="returns"></a>返回
[Image](image.md)

### <a name="appendrichtext(paragraphtext:-string)"></a>appendRichText(paragraphText: string)
将指定文本添加到表格单元格中。

#### <a name="syntax"></a>语法
```js
tableCellObject.appendRichText(paragraphText);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|paragraphText|字符串|要追加的 HTML 字符串。|

#### <a name="returns"></a>返回
[RichText](richtext.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    var appendedRichText = null;
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two and add "Hello".
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                appendedRichText = cell.appendRichText("Hello");
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

### <a name="appendtable(rowcount:-number,-columncount:-number,-values:-string[][])"></a>appendTable(rowCount: number, columnCount: number, values: string[][])
将带有指定行数和列数的表格添加到表格单元格中。

#### <a name="syntax"></a>语法
```js
tableCellObject.appendTable(rowCount, columnCount, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|rowCount|数字|必需。表格的行数。|
|columnCount|数字|必需。表格的列数。|
|值|string[][]|可选。可选的二维数组。如果指定数组中的对应字符串，则填充单元格。|

#### <a name="returns"></a>返回
[Table](table.md)

### <a name="clear()"></a>clear()
清除单元格的内容。

#### <a name="syntax"></a>语法
```js
tableCellObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="load(param:-object)"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
**id、cellIndex、rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load the table cell.
                ctx.load(cell);
                ctx.sync().then(function() {
                    console.log("Cell Id: " + cell.id);
                    console.log("Cell Index: " + cell.cellIndex);
                    console.log("Cell's Row Index: " + cell.rowIndex);
                });
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

**parentTable、cells**
```js
ParentTable, ParentRow, Paragraphs
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a table cell at row one and column two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(1 /*Row Index*/, 2 /*Column Index*/);
                
                // Queue a command to load parentTable, parentRow and paragraphs of the table cell.
                ctx.load(cell, "parentTable, parentRow, paragraphs");
                
                ctx.sync().then(function() {
                    console.log("Parent Table Id: " + cell.parentTable.id);
                    console.log("Parent Row Id: " + cell.parentRow.id);
                    var paragraphs = cell.paragraphs;
                    
                    for (var i = 0; i < paragraphs.items.length; i++) {
                        console.log("Paragraph Id: " + paragraphs.items[i].id);
                    }
                });
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

