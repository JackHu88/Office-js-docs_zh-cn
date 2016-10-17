# <a name="page-object-(javascript-api-for-onenote)"></a>页面对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_   


表示一个 OneNote 页面。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|clientUrl|字符串|页面的客户端 url只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-clientUrl)|
|id|字符串|获取页面的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-id)|
|pageLevel|int|获取或设置页面的缩进级别。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-pageLevel)|
|title|字符串|获取或设置页面的标题。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-title)|
|webUrl|string|页面的 Web URL。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-webUrl)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|内容|[PageContentCollection](pagecontentcollection.md)|页面上 PageContent 对象的集合。只读只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-contents)|
|inkAnalysisOrNull|[InkAnalysis](inkanalysis.md)|页面上墨迹的文本解释。如果没有墨迹分析信息，则返回 null。只读。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-inkAnalysisOrNull)|
|parentSection|[Section](section.md)|获取包含页面的分区。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-parentSection)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[addOutline(left: double, top: double, html:String)](#addoutlineleft-double-top-double-html-string)|[Outline](outline.md)|添加 Outline 至指定位置的页面。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-addOutline)|
|[copyToSection(destinationSection:Section)](#copytosectiondestinationsection-section)|[Page](page.md)|将此页复制到指定的分区中。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-copyToSection)|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|在当前分区之前或之后插入一个新的页面。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-insertPageAsSibling)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-page-load)|

## <a name="method-details"></a>方法详细信息


### <a name="addoutline(left:-double,-top:-double,-html:-string)"></a>addOutline(left: double, top: double, html:String)
添加 Outline 至指定位置的页面。

#### <a name="syntax"></a>语法
```js
pageObject.addOutline(left, top, html);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|left|double|顶部的左边位置，Outline 的左角。|
|top|double|顶部的顶层位置，Outline 的左角。|
|html|字符串|描述边框的可视化演示文稿的 HTML 字符串。请查看 OneNote 外接程序 JavaScript API [支持的 HTML](../../docs/onenote/onenote-add-ins-page-content.md#supported-html)。|

#### <a name="returns"></a>返回
[Outline](outline.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var page = context.application.getActivePage();

    // Queue a command to add an outline with given html. 
    var outline = page.addOutline(200, 200,
"<p>Images and a table below:</p> \
 <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==\"> \
 <img src=\"http://imagenes.es.sftcdn.net/es/scrn/6653000/6653659/microsoft-onenote-2013-01-535x535.png\"> \
 <table> \
   <tr> \
     <td>Jill</td> \
     <td>Smith</td> \
     <td>50</td> \
   </tr> \
   <tr> \
     <td>Eve</td> \
     <td>Jackson</td> \
     <td>94</td> \
   </tr> \
 </table>"     
        );

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
});
```


### <a name="copytosection(destinationsection:-section)"></a>copyToSection(destinationSection:Section)
将此页复制到指定的分区中。

#### <a name="syntax"></a>语法
```js
pageObject.copyToSection(destinationSection);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|destinationSection|分区|要将此页复制到的分区。|

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    
    // Gets the active notebook.
    var notebook = app.getActiveNotebook();
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load sections under the notebook.
    notebook.load('sections');
    
    var newPage;
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync()
        .then(function() {
            var section = notebook.sections.items[0];
            
            // copy page to the section.
            newPage = page.copyToSection(section);
            newPage.load('id');
            return ctx.sync();
        })
        .then(function() {
            console.log(newPage.id);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="insertpageassibling(location:-string,-title:-string)"></a>insertPageAsSibling(location: string, title: string)
在当前分区之前或之后插入一个新的页面。

#### <a name="syntax"></a>语法
```js
pageObject.insertPageAsSibling(location, title);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|location|string|相对于当前页面的新页面的位置。可能的值是：Before、After|
|职位|字符串|新页面的标题。|

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var newPage = activePage.insertPageAsSibling("After", "Next Page");

    // Queue a command to load the newPage to access its data.
    context.load(newPage);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log("page is created with title: " + newPage.title);
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
### <a name="property-access-examples"></a>属性访问示例

**contents**
```js
OneNote.run(function (context) {

    // Gets the active page.
    var activePage = context.application.getActivePage();

    // Queue a command to add a new page after the active page. 
    var pageContents = activePage.contents;

    // Queue a command to load the pageContents to access its data.
    context.load(pageContents);

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            for(var i=0; i < pageContents.items.length; i++)
            {
                var pageContent = pageContents.items[i];
                if (pageContent.type == "Outline")
                {
                    console.log("Found an outline");
                }
                else if (pageContent.type == "Image")
                {
                    console.log("Found an image");
                }
                else if (pageContent.type == "Other")
                {
                    console.log("Found a type not supported yet.");
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

**webUrl**
```js
OneNote.run(function (context) {

    var app = context.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Queue a command to load the webUrl of the page.
    page.load("webUrl");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function() {
            console.log(page.webUrl);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**inkAnalysisOrNull**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load ink words
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    
    return ctx.sync()
        .then(function() {
            if (!page.inkAnalysisOrNull.isNull)
                console.log(page.inkAnalysisOrNull.paragraphs.length);
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

