# Document 对象（适用于 Word 的 JavaScript API）

Document 对象是顶层对象。Document 对象包含一个或多个节、内容控件以及包含文档内容的正文。

_适用于：Word 2016、Word for iPad、Word for Mac_

## 属性
| 属性   | 类型|说明
|:---------------|:--------|:----------|
|Saved|bool|指示是否已保存在文档中所做的更改。如果值为 true，表示文档自上次保存以来并未更改。只读。|

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|body|[Body](body.md)|获取文档的正文。正文是不包括标头、页脚、脚注、文本框等的文本。只读。|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|获取当前文档中的内容控件对象集合。这包括文档正文、标头、页脚、文本框等中的内容控件。只读。|
|Sections|[SectionCollection](sectioncollection.md)|获取文档中的 section 对象集合。只读。|

## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[getSelection()](#getselection)|[Range](range.md)|获取文档的当前选定内容。不支持多重选择。|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[save()](#save)|void|保存文档。如果文档以前未保存过，将使用 Word 的默认文件命名约定。|

## 方法详细信息

### getSelection()
获取文档的当前选定内容。不支持多重选择。

#### 语法
```js
documentObject.getSelection();
```

#### 参数
无

#### 返回
[Range](range.md)

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    var textSample = 'This is an example of the insert text method. This is a method ' + 
        'which allows users to insert text into a selection. It can insert text into a ' +
        'relative location or it can overwrite the current selection. Since the ' +
        'getSelection method returns a range object, look up the range object documentation ' +
        'for everything you can do with a selection.';
    
    // Create a range proxy object for the current selection.
    var range = context.document.getSelection();
    
    // Queue a commmand to insert text at the end of the selection.
    range.insertText(textSample, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Inserted the text at the end of the selection.');
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
    
    // Create a proxy object for the document.
    var thisDocument = context.document;
    
    // Queue a command to load content control properties.
    context.load(thisDocument, 'contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (thisDocument.contentControls.items.length !== 0) {
            for (var i = 0; i < thisDocument.contentControls.items.length; i++) {
                console.log(thisDocument.contentControls.items[i].id);
                console.log(thisDocument.contentControls.items[i].text);
                console.log(thisDocument.contentControls.items[i].tag);
            }
        } else {
            console.log('No content controls in this document.');
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

### 保存。
保存文档。如果文档以前未保存过，将使用 Word 的默认文件命名约定。

#### 语法
```js
documentObject.save();
```

#### 参数
无

#### 返回
无效

#### 示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document.
    var thisDocument = context.document;

    // Queue a commmand to load the document save state (on the saved property).
    context.load(thisDocument, 'saved');    
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        if (thisDocument.saved === false) {
            // Queue a command to save this document.
            thisDocument.save();
            
            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Saved the document');
            });
        } else {
            console.log('The document has not changed since the last save.');
        }
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## 支持详细信息

在运行时检查过程中使用[要求设置](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx)。 
