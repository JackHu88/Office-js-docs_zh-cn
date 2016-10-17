# <a name="outline-object-(javascript-api-for-onenote)"></a>边框对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示 Paragraph 对象的容器。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|字符串|获取边框对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-id)|

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|pageContent|[PageContent](pagecontent.md)|获取包含边框的 PageContent 对象。此对象定义页面上 Outline 的位置。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-pageContent)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|获取“边框”中 Paragraph 对象的集合。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-paragraphs)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[appendHtml(html: string)](#appendhtmlhtml-string)|void|将指定的 HTML 添加到“边框”的底部。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendHtml)|
|[appendImage(base64EncodedImage: string, width: double, height: double)](#appendimagebase64encodedimage-string-width-double-height-double)|[Image](image.md)|将指定的图像添加到“边框”的底部。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendImage)|
|[appendRichText(paragraphText: string)](#appendrichtextparagraphtext-string)|[RichText](richtext.md)|将指定的文本添加到“边框”的底部。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendRichText)|
|[appendTable(rowCount: number, columnCount: number, values: string[][])](#appendtablerowcount-number-columncount-number-values-string)|[Table](table.md)|将具有指定行数和列数的表格添加到边框的底部。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-appendTable)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-outline-load)|

## <a name="method-details"></a>方法详细信息


### <a name="appendhtml(html:-string)"></a>appendHtml(html: string)
将指定的 HTML 添加到“边框”的底部。

#### <a name="syntax"></a>语法
```js
outlineObject.appendHtml(html);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Html|字符串|要追加的 HTML 字符串。请查看 OneNote 外接程序 JavaScript API [支持的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline")
            {
                // First item is an outline.
                outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendHtml("<p>new paragraph</p>");

                // Run the queued commands.
                return context.sync();
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


### <a name="appendimage(base64encodedimage:-string,-width:-double,-height:-double)"></a>appendImage(base64EncodedImage: string, width: double, height: double)
将指定的图像添加到“边框”的底部。

#### <a name="syntax"></a>语法
```js
outlineObject.appendImage(base64EncodedImage, width, height);
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
将指定的文本添加到“边框”的底部。

#### <a name="syntax"></a>语法
```js
outlineObject.appendRichText(paragraphText);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|paragraphText|字符串|要追加的 HTML 字符串。|

#### <a name="returns"></a>返回
[RichText](richtext.md)

### <a name="appendtable(rowcount:-number,-columncount:-number,-values:-string[][])"></a>appendTable(rowCount: number, columnCount: number, values: string[][])
将具有指定行数和列数的表格添加到边框的底部。

#### <a name="syntax"></a>语法
```js
outlineObject.appendTable(rowCount, columnCount, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|rowCount|数字|必需。表格的行数。|
|columnCount|数字|必需。表格的列数。|
|值|string[][]|可选。可选的二维数组。如果指定数组中的对应字符串，则填充单元格。|

#### <a name="returns"></a>返回
[Table](table.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Get pageContents of the activePage. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            if (pageContents.items.length != 0 && pageContents.items[0].type == "Outline") {
                // First item is an outline.
                var outline = pageContents.items[0].outline;

                // Queue a command to append a paragraph to the outline.
                outline.appendTable(2, 2, [[1, 2],[3, 4]]);

                // Run the queued commands.
                return context.sync();
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
