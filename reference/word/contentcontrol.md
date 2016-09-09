# ContentControl 对象（适用于 Word 的 JavaScript API）

表示内容控件。内容控件是文档中绑定的、有可能添加标签的区域，它们充当特定类型的内容的容器。单个内容控件可能包含诸如图像、表或格式化文本段落等内容。当前仅支持富文本内容控件。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|cannotDelete|bool|获取或设置指示用户是否可以删除内容控件的值。与 removeWhenEdited 互相排斥。|
|cannotEdit|bool|获取或设置指示用户是否可以编辑内容控件的内容的值。|
|color|string|获取或设置内容控件的颜色。颜色以“#RRGGBB”格式或使用颜色名称设置。|
|placeholderText|string|获取或设置内容控件的占位符文本。内容控件为空时，将显示灰色的文本。|
|removeWhenEdited|bool|获取或设置指示内容控件在编辑后是否可以删除的值。与 cannotDelete 互相排斥。|
|style|string|获取或设置用于内容控件的样式。这是预安装样式或自定义样式的名称。|
|tag|string|获取或设置用于标识内容控件的标记。[Silly stories](https://aka.ms/sillystorywordaddin) 外接程序示例说明如何使用 **tag** 属性。|
|text|string|获取内容控件的文本。只读。|
|title|string|获取或设置内容控件的标题。|

_请参阅属性访问[示例](#示例)。_

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|外观|**ContentControlAppearance**|获取或设置内容控件的外观。值可以为“boundingBox”、“tags”或“hidden”。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|获取内容控件中内容控件对象的集合。只读。|
|font|[字体](font.md)|获取内容控件的文本格式。使用此对象获取和设置字体名称、大小、颜色和其他属性。只读。|
|id|**[UINT]**|获取表示内容控件标识符的整数。只读。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|获取内容控件中的 inlinePicture 对象的集合。集合不包括浮动图像。只读。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|获取内容控件中的 paragraph 对象的集合。只读。|
|parentContentControl|[ContentControl](contentcontrol.md)|获取包含此内容控件的内容控件。如果不存在父内容控件，返回 null。只读。|
|类型|**ContentControlType**|获取内容控件的类型。当前仅支持富文本内容控件。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除内容控件的内容。用户可以对已清除的内容执行撤消操作。|
|[delete(keepContent: bool)](#deletekeepcontent-bool)|void|删除内容控件及其内容。如果将 keepContent 设置为 true，则不删除内容。|
|[getHtml()](#gethtml)|string|获取内容控件对象的 HTML 表示形式。|
|[getOoxml()](#getooxml)|string|获取内容控件对象的 Office Open XML (OOXML) 表示形式。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定位置插入分隔符。分隔符只能插入到包含在主文档正文内的对象中，除非它是可以插入到任何 body 对象的换行符。insertLocation 值可以为“Before”、“After”、“Start”或“End”。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range 对象设置内联图片](range.md)|将文档插入到当前内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range 对象设置内联图片](range.md)|将 HTML 插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|将嵌入式图片插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。 |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range 对象设置内联图片](range.md)|将 OOXML 或 wordProcessingML 插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph ](paragraph.md)|在指定位置插入段落。insertLocation 值可以为“Before”、“After”、“Start”或“End”。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range 对象设置内联图片](range.md)|将文本插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)|[SearchResultCollection](searchresultcollection.md)|使用指定搜索选项搜索内容控件对象的范围。搜索结果是 range 对象的集合。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|选择内容控件。这会导致 Word 滚动到选定内容。选择模式可以为“Select”、“Start”或“End”。|

## 方法详细信息

### Clear
清除内容控件的内容。用户可以对已清除的内容执行撤消操作。

#### 语法
```js
contentControlObject.clear();
```

#### 参数
无

#### 返回
void

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### delete(keepContent: bool)
删除内容控件及其内容。如果将 keepContent 设置为 true，则不删除内容。

#### 语法
```js
contentControlObject.delete(keepContent);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|keepContent|bool|必需。指示是否应使用内容控件删除内容。如果将 keepContent 设置为 true，则不删除内容。|

#### 返回
void

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {
            
            // Queue a command to delete the first content control. The
            // contents will remain in the document.
            contentControls.items[0].delete(true);
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
            });      
        }
            
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
获取内容控件对象的 HTML 表示形式。

#### 语法
```js
contentControlObject.getHtml();
```

#### 参数
无

#### 返回
string

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls. 
    context.load(contentControlsWithTag, 'tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the HTML contents of the first content control.
            var html = contentControlsWithTag.items[0].getHtml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control HTML: ' + html.value);
            });
        }
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
获取内容控件对象的 Office Open XML (OOXML) 表示形式。

#### 语法
```js
contentControlObject.getOoxml();
```

#### 参数
无

#### 返回
string

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to get the OOXML contents of the first content control.
            var ooxml = contentControls.items[0].getOoxml();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Content control OOXML: ' + ooxml.value);
            });
        }
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
在指定位置插入分隔符。分隔符只能插入到包含在主文档正文内的对象中，除非它是可以插入到任何 body 对象的换行符。insertLocation 值可以为“Before”、“After”、“Start”或“End”。

#### 语法
```js
contentControlObject.insertBreak(breakType, insertLocation);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|breakType|BreakType|必需。分隔符的类型 (breakType.md)|
|insertLocation|InsertLocation|必需。此值可以为“Before”、“After”、“Start”或“End”。|

#### 返回
void

#### 其他详细信息
除了换行符之外，您不能在标头、页脚、脚注、尾注、注释和文本框包含的对象中插入分隔符。  

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a commmand to load the id property for all of content controls. 
    context.load(contentControls, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion. We now will have 
    // access to the content control collection.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a page break after the first content control. 
            contentControls.items[0].insertBreak('page', "After");
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion. 
            return context.sync()
                .then(function () {
                    console.log('Inserted a page break after the first content control.');    
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
将文档插入到当前内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
contentControlObject.insertFileFromBase64(base64File, insertLocation);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64File|string|必需。要插入的文件的 base64 编码内容。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range 对象设置内联图片](range.md)

### insertHtml(html: string, insertLocation:InsertLocation)
将 HTML 插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
contentControlObject.insertHtml(html, insertLocation);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Html|string|必需。要插入到内容控件中的 HTML。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put HTML into the contents of the first content control.
            contentControls.items[0].insertHtml('<strong>HTML content inserted into the content control.</strong>', 'Start');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted HTML in the first content control.');
            });
        }
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
将嵌入式图片插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
contentControlObject.insertInlinePictureFromBase64(image, insertLocation);

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必需。将 base64 编码的图像插入内容控件中。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[InlinePicture](inlinepicture.md)



### insertOoxml(ooxml: string, insertLocation: InsertLocation)
将 OOXML 或 wordProcessingML 插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
contentControlObject.insertOoxml(ooxml, insertLocation);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|ooxml|string|必需。要插入到内容控件中的 OOXML 或 wordProcessingML。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to put OOXML into the contents of the first content control.
            contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted OOXML in the first content control.');
            });
        }
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
在指定位置插入段落。insertLocation 值可以为“Before”、“After”、“Start”或“End”。

#### 语法
```js
contentControlObject.insertParagraph(paragraphText, insertLocation);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|paragraphText|string|必需。要插入的段落文本。|
|insertLocation|InsertLocation|必需。此值可以为“Before”、“After”、“Start”或“End”。|

#### 返回
[Paragraph ](paragraph.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to insert a paragraph after the first content control. 
            contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Inserted a paragraph after the first content control.');
            });
        }
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
将文本插入到内容控件中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### 语法
```js
contentControlObject.insertText(text, insertLocation);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|text|string|必需。要插入到内容控件中的文本。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to replace text in the first content control. 
            contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Replaced text in the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) 外接程序示例说明如何使用 **insertText** 方法。

### load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy range object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to create the content control.
    var myContentControl = range.insertContentControl();
    myContentControl.tag = 'Customer-Address';
    myContentControl.title = ' has t';
    myContentControl.style = 'Heading 2';
    myContentControl.insertText('One Microsoft Way, Redmond, WA 98052', 'replace');
    myContentControl.cannotEdit = true;
    
    // Queue a command to load the id property for the content control you created.
    context.load(myContentControl, 'id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Created content control with id: ' + myContentControl.id);
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
使用指定搜索选项搜索内容控件对象的范围。搜索结果是 range 对象的集合。

#### 语法
```js
contentControlObject.search(searchText, searchOptions);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|searchText|string|必须。搜索文本。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|可选。用于搜索的选项。|

#### 返回
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
选择内容控件。这会导致 Word 滚动到选定内容。选择模式可以为“Select”、“Start”或“End”。

#### 语法
```js
contentControlObject.select(selectionMode);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|可选。选择模式可以为“Select”、“Start”或“End”。“Select”为默认值。|

#### 返回
void

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to select the first content control.
            contentControls.items[0].select();
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Selected the first content control.');
            });
        }
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## 属性访问示例

### 加载所有的内容控件属性
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls. 
    context.load(contentControls, 'id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control. 
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');             
        
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' + 
                        '   ----- appearance: ' + contentControls.items[0].appearance + 
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
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
在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。