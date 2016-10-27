# <a name="body-object-(javascript-api-for-word)"></a>Body 对象（适用于 Word 的 JavaScript API）

表示文档或节的正文。

_适用于：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|style|string|获取或设置用于正文的样式。这是预安装样式或自定义样式的名称。|
|text|string|获取正文的文本。使用 insertText 方法插入文本。只读。|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|获取正文中的富文本内容控件对象集合。只读。|
|font|[Font](font.md)|获取正文的文本格式。使用此对象获取和设置字体名称、大小、颜色和其他属性。只读。|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|获取正文中的嵌入式图片对象集合。集合不包括浮动图像。只读。|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|获取正文中的段落对象集合。只读。|
|parentContentControl|[ContentControl](contentcontrol.md)|获取包含正文的内容控件。如果不存在父内容控件，返回 null。只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|清除 body 对象的内容。用户可以对已清除的内容执行撤消操作。|
|[getHtml()](#gethtml)|string|获取 body 对象的 HTML 表示形式。|
|[getOoxml()](#getooxml)|string|获取 body 对象的 OOXML (Office Open XML) 表示形式。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定位置插入分隔符。分隔符只能插入到主文档正文中，除非它是可以插入到任何 body 对象的换行符。insertLocation 值可以为“Start”或“End”。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|使用富文本内容控件封装 body 对象。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|将文档插入到正文中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|在指定位置插入 HTML。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|将图片插入到正文中的指定位置。insertLocation 值可以为“Start”或“End”。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|在指定位置插入 OOXML。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定位置插入段落。insertLocation 值可以为“Start”或“End”。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|将文本插入到正文中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|使用指定搜索选项搜索 body 对象的范围。搜索结果是 range 对象的集合。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|选择正文并在 Word UI 中进行浏览。SelectionMode 值可以为“Select”、“Start”或“End”。|

## <a name="method-details"></a>方法详细信息

### <a name="clear()"></a>clear()
清除 body 对象的内容。用户可以对已清除的内容执行撤消操作。

#### <a name="syntax"></a>语法
```js
bodyObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to clear the contents of the body.
    body.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) 外接程序示例说明如何使用 **clear** 方法清除文档内容。

### <a name="gethtml()"></a>getHtml()
获取 body 对象的 HTML 表示形式。

#### <a name="syntax"></a>语法
```js
bodyObject.getHtml();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
string

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="getooxml()"></a>getOoxml()
获取 body 对象的 OOXML (Office Open XML) 表示形式。

#### <a name="syntax"></a>语法
```js
bodyObject.getOoxml();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
string

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)
在指定位置插入分隔符。分隔符只能插入到主文档正文中，除非它是可以插入到任何 body 对象的换行符。insertLocation 值可以为“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|breakType|BreakType|必需。要添加到正文的分隔符类型。|
|insertLocation|InsertLocation|必需。此值可以为“Start”或“End”。|

#### <a name="returns"></a>返回
void

#### <a name="additional-details"></a>其他详细信息
除了换行符之外，您不能在标头、页脚、脚注、尾注、注释和文本框中插入分隔符。

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {

    // Create a proxy object for the document body.
    var body = ctx.document.body;

    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### <a name="insertcontentcontrol()"></a>insertContentControl()
使用富文本内容控件封装 body 对象。

#### <a name="syntax"></a>语法
```js
bodyObject.insertContentControl();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[ContentControl](contentcontrol.md)

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
将文档插入到正文中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64File|string|必需。要插入的 base64 编码的文件内容。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

[Silly stories](https://aka.ms/sillystorywordaddin) 外接程序示例说明如何使用 **insertFileFromBase64** 方法插入服务中的 docx 文件。

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation:InsertLocation)
在指定位置插入 HTML。insertLocation 值可以为“Replace”、“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Html|string|必需。要插入到文档中的 HTML。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
将图片插入到正文中的指定位置。insertLocation 值可以为“Start”或“End”。

#### <a name="syntax"></a>语法
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必需。将 base64 编码的图像插入正文中。|
|insertLocation|InsertLocation|必需。此值可以为“Start”或“End”。|

#### <a name="returns"></a>返回
[InlinePicture](inlinepicture.md)

### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
在指定位置插入 OOXML。insertLocation 值可以为“Replace”、“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|ooxml|string|必需。要插入的 OOXML 或 wordProcessingML。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="known-issues"></a>已知问题
此方法导致 Word Online 中的延迟时间较长，从而影响用户对外接程序的体验。我们建议仅当其他解决方案不可用时才使用此方法。 

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>其他信息
阅读 [使用 Office Open XML 创建更好的 Word 外接程序](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx) 以获取使用 OOXML 的指南。[Word-Add-in-DocumentAssembly][body.insertOoxml] 示例显示如何使用此 API 来汇编文档。

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation:InsertLocation)
在指定位置插入段落。insertLocation 值可以为“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|paragraphText|string|必需。要插入的段落文本。|
|insertLocation|InsertLocation|必需。此值可以为“Start”或“End”。|

#### <a name="returns"></a>返回
[Paragraph](paragraph.md)

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### <a name="additional-information"></a>其他信息
[Word-Add-in-DocumentAssembly][body.insertParagraph] 示例显示如何使用 insertParagraph 方法来汇编文档。

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation:InsertLocation)
将文本插入到正文中的指定位置。insertLocation 值可以为“Replace”、“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|text|string|必需。要插入的文本。|
|insertLocation|InsertLocation|必需。此值可以为“Replace”、“Start”或“End”。|

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### <a name="search(searchtext:-string,-searchoptions:-paramtypestrings.searchoptions)"></a>search(searchText: string, searchOptions:ParamTypeStrings.SearchOptions)
使用指定搜索选项搜索 body 对象的范围。搜索结果是 range 对象的集合。

#### <a name="syntax"></a>语法
```js
bodyObject.search(searchText, searchOptions);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|searchText|string|必需。搜索文本。|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|可选。用于搜索的选项。|

#### <a name="returns"></a>返回
[SearchResultCollection](searchresultcollection.md)

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length +
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
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

#### <a name="additional-information"></a>其他信息
[Word-Add-in-DocumentAssembly][body.search] 示例提供了如何搜索文档的另一个示例。

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
选择正文并在 Word UI 中进行浏览。SelectionMode 值可以为“Select”、“Start”或“End”。

#### <a name="syntax"></a>语法
```js
bodyObject.select(selectionMode);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|可选。选择模式可以为“Select”、“Start”或“End”。“Select”为默认值。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to select the document body. The Word UI will
    // move to the selected document body.
    body.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="property-access-examples"></a>属性访问示例

### <a name="get-the-text-property-on-the-body-object"></a>获取 body 对象的文本属性
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load the text in document body.
    context.load(body, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### <a name="get-the-style-and-the-font-size,-font-name,-and-font-color-properties-on-the-body-object."></a>获取 body 对象的样式和字体大小、字体名称和字体颜色属性。

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="support-details"></a>支持详细信息

在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


[body.insertOoxml]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "插入 OOXML"
[body.insertParagraph]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "插入段落"
[body.search]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "正文搜索"
