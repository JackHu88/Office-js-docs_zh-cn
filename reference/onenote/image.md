# <a name="image-object-(javascript-api-for-onenote)"></a>图像对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示 Image。Image 可以是 PageContent 对象或 Paragraph 对象的直接子级。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|description|字符串|获取或设置 Image 的说明。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-description)|
|高度|double|获取或设置 Image 布局的高度。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-height)|
|超链接|字符串|获取或设置 Image 的超链接。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-hyperlink)|
|id|string|获取“图像”对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-id)|
|width|double|获取或设置 Image 布局的宽度。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-width)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|ocrData|[ImageOcrData](imageocrdata.md)|获取由此 Image 的 OCR（光学字符识别）获取的数据，如 OCR 文本和语言。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-ocrData)|
|pageContent|[PageContent](pagecontent.md)|获取包含 Image 的 PageContent 对象。如果 Image 不是 PageContent 的直接子级，则引发。此对象定义页面上的 Image 位置。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-pageContent)|
|paragraph|[Paragraph](paragraph.md)|获取包含 Image 的 Paragraph 对象。如果 Image 不是 Paragraph 的直接子级，则引发。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-paragraph)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getBase64Image()](#getbase64image)|string|获取 Image 的 base64 编码的二进制表示形式。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-getBase64Image)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-image-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getbase64image()"></a>getBase64Image()
获取 Image 的 base64 编码的二进制表示形式。

#### <a name="syntax"></a>语法
```js
imageObject.getBase64Image();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
string

#### <a name="examples"></a>示例
```js

var image = null;
var imageString;

OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                imageString = image.getBase64Image();
                return ctx.sync();
            }
        })
        .then(function(){
            console.log(imageString);
        });
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
**ID、宽度、高度、说明和超链接**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var image = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
        })
        .then(function(){
            if (image != null)
            {
                // load every properties and relationships
                ctx.load(image);
                return ctx.sync();
            }
        })
        .then(function(){
            if (image != null)
            {                   
                console.log("image " + image.id + " width is " + image.width + " height is " + image.height);
                console.log("description: " + image.description);                   
                console.log("hyperlink: " + image.hyperlink);
            }
        });
});
```

**ocrData**
```js
var image = null;

OneNote.run(function(ctx){
    // Get the current outline.
    var outline = ctx.application.getActiveOutline();

    // Queue a command to load paragraphs and their types.
    outline.load("paragraphs")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
            if (image != null)
            {
               image.load("ocrData");
            }
            return ctx.sync();
        })
        .then(function(){
            console.log(image.ocrData);
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**paragraph**
```js
OneNote.run(function(ctx){
    // Get the current outline.         
    var outline = ctx.application.getActiveOutline();
    var searchedParagraph = null;
    
    // Queue a command to load paragraphs and their types. 
    outline.load("paragraphs/type")
    return ctx.sync().
        then(function() {
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    searchedParagraph = paragraph;
                    break;
                }
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {
                // load every properties and relationships
                searchedParagraph.image.load('paragraph');
                return ctx.sync();
            }
        })
        .then(function() {
            if (searchedParagraph != null)
            {                   
                if (searchedParagraph.id != searchedParagraph.image.paragraph.id)
                {
                    console.log("id must match");
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

