# <a name="paragraph-object-javascript-api-for-onenote"></a>Paragraph 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


页面上可见内容的容器。一个 Paragraph 可包含任意一个 ParagraphType 类型的内容。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|string|获取 Paragraph 对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-id)|
|type|string|获取 Paragraph 对象的类型。只读。可能的值是：RichText、Image、Table、Other。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-type)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|image|[Image](image.md)|获取 Paragraph 中的 Image 对象。如果 ParagraphType 不是 Image，则引发异常。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-image)|
|inkWords|[InkWordCollection](inkwordcollection.md)|获取 Paragraph 中的 Ink 集合。如果 ParagraphType 不为 Ink，则引发异常。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-inkWords)|
|outline|[Outline](outline.md)|获取包含“段落”的“边框”对象。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-outline)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|此段落下的段落集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-paragraphs)|
|parentParagraph|[Paragraph](paragraph.md)|获取父段落对象。如果父段落不存在，则引发。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraph)|
|parentParagraphOrNull|[Paragraph](paragraph.md)|获取父段落对象。如果父段落不存在，则返回 null。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentParagraphOrNull)|
|parentTableCell|[TableCell](tablecell.md)|获取包含 Paragraph 的 TableCell 对象（如果存在）。如果父级不为 TableCell，则引发 ItemNotFound。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCell)|
|parentTableCellOrNull|[TableCell](tablecell.md)|获取包含 Paragraph 的 TableCell 对象（如果存在）。如果父级不为 TableCell，则返回 null。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-parentTableCellOrNull)|
|richText|[RichText](richtext.md)|获取 Paragraph 中的 RichText 对象。如果 ParagraphType 不为 RichText，则引发异常。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-richText)|
|table|[表](table.md)|获取 Paragraph 中的 Table 对象。如果 ParagraphType 不为 Table，则引发异常。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-table)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[delete()](#delete)|void|删除 paragraph|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-delete)|
|[insertHtmlAsSibling(insertLocation: string, html: string)](#inserthtmlassiblinginsertlocation-string-html-string)|void|插入指定的 HTML 内容|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertHtmlAsSibling)|
|[insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)](#insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double)|[Image](image.md)|在指定的插入位置插入图像。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertImageAsSibling)|
|[insertRichTextAsSibling(insertLocation: string, paragraphText: string)](#insertrichtextassiblinginsertlocation-string-paragraphtext-string)|[RichText](richtext.md)|在指定的插入位置插入段落文本。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertRichTextAsSibling)|
|[insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])](#inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string)|[Table](table.md)|将具有指定行数和列数的表格添加到当前段落的之前或之后。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-insertTableAsSibling)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-paragraph-load)|

## <a name="method-details"></a>方法详细信息


### <a name="delete"></a>delete()
删除 paragraph

#### <a name="syntax"></a>语法
```js
paragraphObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    
    var paragraphs = pageContent.outline.paragraphs;
    
    var firstParagraph = paragraphs.getItemAt(0);
    
    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Queue a command to delete the first paragraph                 
            firstParagraph.delete();
            
            // Run the command to delete it
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="inserthtmlassiblinginsertlocation-string-html-string"></a>insertHtmlAsSibling(insertLocation: string, html: string)
插入指定的 HTML 内容

#### <a name="syntax"></a>语法
```js
paragraphObject.insertHtmlAsSibling(insertLocation, html);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|insertLocation|string|相对于当前 Paragraph 的新内容的位置。可能的值是：Before、After|
|Html|string|描述内容的可视化演示文稿的 HTML 字符串。请查看 OneNote 外接程序 JavaScript API [支持的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertHtmlAsSibling("Before", "<p>ContentBeforeFirstParagraph</p>");
            firstParagraph.insertHtmlAsSibling("After", "<p>ContentAfterFirstParagraph</p>");
            
            // Run the command to run inserts
            return context.sync();
        });
))
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertimageassiblinginsertlocation-string-base64encodedimage-string-width-double-height-double"></a>insertImageAsSibling(insertLocation: string, base64EncodedImage: string, width: double, height: double)
在指定的插入位置插入图像。

#### <a name="syntax"></a>语法
```js
paragraphObject.insertImageAsSibling(insertLocation, base64EncodedImage, width, height);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|insertLocation|string|相对于当前段落的表格的位置。可能的值是：Before、After|
|base64EncodedImage|string|要追加的 HTML 字符串。|
|宽度|double|可选。以磅为单位的宽度。默认值为 null，将考虑图像宽度。|
|height|double|可选。以磅为单位的高度。默认值为 null，将考虑图像高度。|

#### <a name="returns"></a>返回
[Image](image.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertImageAsSibling("Before", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            firstParagraph.insertImageAsSibling("After", "R0lGODlhDwAPAKECAAAAzMzM/////wAAACwAAAAADwAPAAACIISPeQHsrZ5ModrLlN48CXF8m2iQ3YmmKqVlRtW4MLwWACH+H09wdGltaXplZCBieSBVbGVhZCBTbWFydFNhdmVyIQAAOw==");
            
            // Run the command to insert images
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### <a name="insertrichtextassiblinginsertlocation-string-paragraphtext-string"></a>insertRichTextAsSibling(insertLocation: string, paragraphText: string)
在指定的插入位置插入段落文本。

#### <a name="syntax"></a>语法
```js
paragraphObject.insertRichTextAsSibling(insertLocation, paragraphText);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|insertLocation|string|相对于当前段落的表格的位置。可能的值是：Before、After|
|paragraphText|string|要追加的 HTML 字符串。|

#### <a name="returns"></a>返回
[RichText](richtext.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page
    // Assuming its an outline, get the outline's paragraphs.
    var pageContent = pageContents.getItemAt(0);
    var paragraphs = pageContent.outline.paragraphs;
    var firstParagraph = paragraphs.getItemAt(0);

    // Queue a command to load the id and type of the first paragraph
    firstParagraph.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            // Queue commands to insert before and after the first paragraph
            firstParagraph.insertRichTextAsSibling("Before", "Text Appears Before Paragraph");
            firstParagraph.insertRichTextAsSibling("After", "Text Appears After Paragraph");
            
            // Run the command to insert text contents
            return context.sync();
        });
})  
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```


### <a name="inserttableassiblinginsertlocation-string-rowcount-number-columncount-number-values-string"></a>insertTableAsSibling(insertLocation: string, rowCount: number, columnCount: number, values: string[][])
将具有指定行数和列数的表格添加到当前段落的之前或之后。

#### <a name="syntax"></a>语法
```js
paragraphObject.insertTableAsSibling(insertLocation, rowCount, columnCount, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|insertLocation|string|相对于当前段落的表格的位置。可能的值是：Before、After|
|rowCount|数字|表格的行数。|
|columnCount|数字|表格的列数。|
|值|string[][]|可选。可选的二维数组。如果指定数组中的对应字符串，则填充单元格。|

#### <a name="returns"></a>返回
[Table](table.md)

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
### <a name="property-access-examples"></a>属性访问示例

**id 和 type**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;
    
    // Queue a command to load the outline property of each pageContent.
    pageContents.load("outline");
        
    // Get the first PageContent on the page, and then get its Outline.
    var pageContent = pageContents._GetItem(0);
    var paragraphs = pageContent.outline.paragraphs;
            
    // Queue a command to load the id and type of each paragraph.
    paragraphs.load("id,type");
            
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            
            // Write the text.                  
            $.each(paragraphs.items, function(index, paragraph) {
                console.log("Paragraph type: " + paragraph.type);
                console.log("Paragraph ID: " + paragraph.id);
            });
        });
})      
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```

**paragraphs**
```js
OneNote.run(function(context) {
    var app = context.application;
    
    // Gets the active outline
    var outline = app.getActiveOutline();
    
    // load nested paragraphs and their types.
    outline.load("paragraphs/type");
    
    return context.sync().then(function () {
        var paragraphs = outline.paragraphs.items;
        
        var promise;
        // for each nested paragraphs, load tables only
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];
            if (paragraph.type == "Table") {
                paragraph.load("table/id");
                promise =  context.sync().then(function() {
                    console.log(paragraph.table.id);
                });
            }
        }
        return promise;
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

