# Paragraph 对象（适用于 Word 的 JavaScript API）

表示选定内容、区域、内容控件或文档正文中的单个段落。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
| 属性   | 类型|说明
|:---------------|:--------|:----------|
|outlineLevel|int|获取或设置段落的大纲级别。|
|style|string|获取或设置用于段落的样式。这是预安装样式或自定义样式的名称。[Word-Add-in-DocumentAssembly][paragraph.style] 示例显示如何设置段落样式。|
|text|string|获取段落的文本。只读。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|Alignment|**Alignment**|获取或设置段落的对齐方式。值可以为“left”、“centered”、“right”或“justified”。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|获取段落中的内容控件对象集合。只读。|
|firstLineIndent|浮点型。|获取或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。|
|font|[Font](font.md)|获取段落的文本格式。使用此对象获取和设置字体名称、大小、颜色和其他属性。只读。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|获取段落中的 inlinePicture 对象的集合。集合不包括浮动图像。只读。|
|leftIndent|**float**|获取或设置段落的向左缩进值（以磅值表示）。|
|lineSpacing|**float**|获取或设置指定段落的行间距（以磅值表示）。在 Word UI 中，该值应除以 12。|
|lineUnitAfter|**float**|获取或设置段落后面的网格线中的间隔量。|
|lineUnitBefore|**float**|获取或设置段落前面的网格线中的间隔量。|
|parentContentControl|[ContentControl](contentcontrol.md)|获取包含段落的内容控件。如果不存在父内容控件，返回 null。只读。|
|rightIndent|**float**|获取或设置段落的向右缩进值（以磅值表示）。|
|spaceAfter|**float**|获取或设置段落后面的间距（以磅值表示）。|
|spaceBefore|**float**|获取或设置段落前面的间距（以磅值表示）。|

## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除 paragraph 对象的内容。用户可以对已清除的内容执行撤消操作。|
|[delete()](#delete)|void|从文档中删除段落及其内容。|
|[getHtml()](#gethtml)|string|获取 paragraph 对象的 HTML 表示形式。|
|[getOoxml()](#getooxml)|string|获取 paragraph 对象的 Office Open XML (OOXML) 表示形式。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定位置插入分隔符。分隔符只能插入到包含在主文档正文内的段落中，除非它是可以插入到任何 body 对象的换行符。insertLocation 值可以为“After”或“Before”。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|使用富文本内容控件封装 paragraph 对象。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|将文档插入到当前段落中的指定位置。insertLocation 值可以为“Start”或“End”。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|将 HTML 插入到段落中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|将图片插入到段落中的指定位置。insertLocation 值可以为“Before”、“After”、“Start”或“End”。|
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|将 OOXML 或 wordProcessingML 插入到段落中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定位置插入段落。insertLocation 值可以为“Before”或“After”。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|将文本插入到段落中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|使用指定搜索选项搜索 paragraph 对象的范围。搜索结果是 range 对象的集合。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|选择并在 Word UI 中导航到段落。选择模式可以为“Select”、“Start”或“End”。“Select”为默认值。|

## 方法详细信息

### clear()
清除 paragraph 对象的内容。用户可以对已清除的内容执行撤消操作。

#### 语法
```js
paragraphObject.clear();
```

#### 参数
无

#### 返回
无效

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### delete()
从文档中删除段落及其内容。

#### 语法
```js
paragraphObject.delete();
```

#### 参数
无

#### 返回
无效

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getHtml()
获取 paragraph 对象的 HTML 表示形式。

#### 语法
```js
paragraphObject.getHtml();
```

#### 参数
无

#### 返回
字符串

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
获取 paragraph 对象的 Office Open XML (OOXML) 表示形式。

#### 语法
```js
paragraphObject.getOoxml();
```

#### 参数
无

#### 返回
字符串

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
在指定位置插入分隔符。分隔符只能插入到包含在主文档正文内的段落中，除非它是可以插入到任何 body 对象的换行符。insertLocation 值可以为“Before”或“After”。

#### 语法
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|breakType|BreakType|必需。要添加到文档的分隔符类型。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### 返回
void

#### 其他详细信息
您不能在标头、页脚、脚注、尾注、注释和文本框中插入分隔符。 

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertContentControl()
使用富文本内容控件封装 paragraph 对象。

#### 语法
```js
paragraphObject.insertContentControl();
```

#### 参数
无

#### 返回
[ContentControl](contentcontrol.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl(); 
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### 其他信息
[Word-Add-in-DocumentAssembly][paragraph.insertContentControl] 示例演示如何使用 insertContentControl 方法。

### insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
将文档插入到当前段落中的指定位置。insertLocation 值可以为“Start”或“End”。

#### 语法
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|base64File|string|必需。要插入的 base64 编码的文件内容。|
|insertLocation|InsertLocation|必需。此值可以为“Start”或“End”。|

#### 返回
[Range](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert base64 encoded .docx at the beginning of the first paragraph.
        // This won't work unless you have a definition for getBase64().
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted base64 encoded content at the beginning of the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertHtml(html: string, insertLocation:InsertLocation)
将 HTML 插入到段落中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|Html|string|必需。要插入到段落中的 HTML。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
将图片插入到段落中的指定位置。insertLocation 值可以为“Before”、“After”、“Start”或“End”。

#### 语法
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必需。要插入到段落中的 HTML。|
|insertLocation|InsertLocation|必需。此值可以为“Before”、“After”、“Start”或“End”。|

#### 返回
[InlinePicture](inlinepicture.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### 其他信息
[Word-Add-in-DocumentAssembly][paragraph.insertpicture] 示例提供了另一个示例，演示如何将一个图像插入到一个段落中。

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
将 OOXML 或 wordProcessingML 插入到段落中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|ooxml|string|必需。要插入到段落中的 OOXML 或 wordProcessingML。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert Ooxml content into the first paragraph.
        var ooxmlContent = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted OOXML at the end of the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### 其他信息
阅读[使用 Office Open XML 创建更好的 Word 外接程序](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx)以获取使用 OOXML 的指南。

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
在指定位置插入段落。insertLocation 值可以为“Before”或“After”。

#### 语法
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|paragraphText|string|必需。要插入的段落文本。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### 返回
[Paragraph](paragraph.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertText(text: string, insertLocation:InsertLocation)
将文本插入到段落中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
paragraphObject.insertText(text, insertLocation);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|text|string|必需。要插入的文本。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
无效

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we 
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong>--" +
                          "--Font size: " + paragraph.font.size +
                          "--Font name: " + paragraph.font.name +
                          "--Font color: " + paragraph.font.color +
                          "--Style: " + paragraph.style;

            console.log(results);
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
使用指定搜索选项搜索 paragraph 对象的范围。搜索结果是 range 对象的集合。

#### 语法
```js
paragraphObject.search(searchText, searchOptions);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|searchText|string|必须。搜索文本。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|可选。用于搜索的选项。|

#### 返回
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
选择并在 Word UI 中导航到段落。

#### 语法
```js
paragraphObject.select(selectionMode);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|可选。选择模式可以为“Select”、“Start”或“End”。“Select”为默认值。|

#### 返回
无效

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the last paragraph a create a 
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1]; 
        
        // Queue a command to select the paragraph. The Word UI will 
        // move to the selected paragraph.
        paragraph.select();
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
        });      
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## 支持详细信息

在运行时检查过程中使用[要求设置](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)。 


[paragraph.insertContentControl]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L161 "insert content control"[paragraph.style]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L172 "set style" [paragraph.insertpicture]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L236 "insert picture"
