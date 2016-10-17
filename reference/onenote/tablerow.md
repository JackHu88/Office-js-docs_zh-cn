# <a name="tablerow-object-(javascript-api-for-onenote)"></a>TableRow 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示表中的行。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|cellCount|int|获取行中的单元格数。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cellCount)|
|id|字符串|获取行的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-id)|
|rowIndex|int|获取其父表中的行索引。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-rowIndex)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|cells|[TableCellCollection](tablecellcollection.md)|获取行中的单元格。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-cells)|
|parentTable|[Table](table.md)|获取父表。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-parentTable)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[clear()](#clear)|void|清除行的内容。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-clear)|
|[insertRowAsSibling(insertLocation: string, values: string[])](#insertrowassiblinginsertlocation-string-values-string)|[TableRow](tablerow.md)|在当前行之前或之后插入一行。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-insertRowAsSibling)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|设置行中所有单元格的底纹色。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRow-setShadingColor)|

## <a name="method-details"></a>方法详细信息


### <a name="clear()"></a>clear()
清除行的内容。

#### <a name="syntax"></a>语法
```js
tableRowObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="insertrowassibling(insertlocation:-string,-values:-string[])"></a>insertRowAsSibling(insertLocation: string, values: string[])
在当前行之前或之后插入一行。

#### <a name="syntax"></a>语法
```js
tableRowObject.insertRowAsSibling(insertLocation, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|insertLocation|字符串|新行应插入的位置对应于当前行。可能的值是：Before、After|
|值|string[]|可选。在新行中插入的字符串，指定为数组。单元格不能多于当前行中的单元格。可选。|

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
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                
                // Run the queued commands
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    rows.items[1].insertRowAsSibling("Before", ["cell0", "cell1"]);
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

### <a name="setshadingcolor(colorcode:-string)"></a>setShadingColor(colorCode: string)
设置行中所有单元格的底纹色。

#### <a name="syntax"></a>语法
```js
tableRowObject.setShadingColor(colorCode);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|colorCode|string|要为单元格设置的颜色代码。/参数|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
**id、cellCount、rowIndex**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load table.rows.
                ctx.load(table, "rows");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each table row, log cell count and row index.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
                        console.log("Row " + i + " Cell Count: " + rows.items[i].cellCount);
                        console.log("Row " + i + " Row Index: " + rows.items[i].rowIndex);
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

**parentTable、cells**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get table rows.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                
                // Queue a command to load parentTable and cells of each row in the table.
                ctx.load(table, "rows/parentTable, rows/cells");
                return ctx.sync().then(function() {
                    var rows = table.rows;
                    
                    // for each row, log parentTable and cells
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Parent Table Id: " + rows.items[i].parentTable.id);
                        var cells = rows.items[i].cells;
                        for (var j = 0 ; j < cells.items.length; j++) {
                            console.log("Row " + i + " Cell " + j + " Id: " + cells.items[j].id);
                        }
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

