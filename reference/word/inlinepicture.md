# <a name="inlinepicture-object-(javascript-api-for-word)"></a>InlinePicture 对象（适用于 Word 的 JavaScript API）

表示嵌入式图片。

_适用于：Word 2016、Word for iPad、Word for Mac、Word Online_

## <a name="properties"></a>属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|altTextDescription|string|获取或设置表示与嵌入式图像相关的替代文本的字符串。|
|altTextTitle|string|获取或设置包含嵌入式图像的标题的字符串。|
|hyperlink|string|获取或设置与嵌入式图像相关的超链接。|
|lockAspectRatio|bool|获取或设置指示在您调整嵌入式图像大小时其是否保留原始比例的值。|

## <a name="relationships"></a>Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|高度|**float**|获取或设置描述嵌入式图像的高度的数字。此值以点为单位计量。 |
|parentContentControl|[ContentControl](contentcontrol.md)|获取包含嵌入式图像的内容控件。如果不存在父内容控件，返回 null。只读。|
|段落|[paragraph](paragraph.md)|获取包含嵌入式图像的段落。只读。
|width|**float**|获取或设置描述嵌入式图像的宽度的数字。此值以点为单位计量。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|从文档中删除该图片。|
|[getBase64ImageSrc()](#getbase64imagesrc)|object|获取数值采用嵌入式图像的 base64 编码字符串表示的对象。|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|在指定位置插入分隔符。insertLocation 值可以为“Before”或“After”。|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|使用富文本内容控件封装嵌入式图像。|
|[insertFileFromBase64(base64File: string, insertLocation:InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|将文档插入到正文中的指定位置。insertLocation 值可以为“Before”或“After”。|
|[insertHtml(html: string, insertLocation:InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|在指定位置插入 HTML。insertLocation 值可以为“Before”或“After”。|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|将图片插入到正文中的指定位置。insertLocation 值可以为“Replace”、“Before”或“After”。 |
|[insertOoxml(ooxml: string, insertLocation:InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|在指定位置插入 OOXML。insertLocation 值可以为“Before”或“After”。|
|[insertParagraph(paragraphText: string, insertLocation:InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|在指定位置插入段落。insertLocation 值可以为“Before”或“After”。|
|[insertText(text: string, insertLocation:InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|将文本插入到正文中的指定位置。insertLocation 值可以为“Before”或“After”。|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|选择图片并在 Word UI 中进行浏览。SelectionMode 值可以为“Select”、“Start”或“End”。|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息

### <a name="delete()"></a>delete()
从文档中删除该图片。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="getbase64imagesrc()"></a>getBase64ImageSrc()
获取数值采用嵌入式图像的 base64 编码字符串表示的对象。

#### <a name="syntax"></a>语法
```js
var base64 = inlinePictureObject.getBase64ImageSrc();
return context.sync().then(function () {    
    console.log("base64 string is " + base64.value);
});

```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
对象 



### <a name="insertbreak(breaktype:-breaktype,-insertlocation:-insertlocation)"></a>insertBreak(breakType: BreakType, insertLocation: InsertLocation)

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertBreak(breakType, insertLocation);
```
#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|breakType|BreakType|必需。要添加到正文的分隔符类型。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
void

### <a name="insertcontentcontrol()"></a>insertContentControl()
使用富文本内容控件封装嵌入式图像。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertContentControl();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[ContentControl](contentcontrol.md)

### <a name="insertfilefrombase64(base64file:-string,-insertlocation:-insertlocation)"></a>insertFileFromBase64(base64File: string, insertLocation:InsertLocation)
将文档插入到正文中的指定位置。insertLocation 值可以为“Before”或“After”。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertFileFromBase64(base64File, insertLocation);
```
#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64File|string|必需。Base64 编码的 docx 文件内容。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="inserthtml(html:-string,-insertlocation:-insertlocation)"></a>insertHtml(html: string, insertLocation:InsertLocation)
在指定位置插入 HTML。insertLocation 值可以为“Before”或“After”。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertHtml(html, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Html|string|必需。要插入到文档中的 HTML。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
[Range](range.md)


### <a name="insertinlinepicturefrombase64(base64encodedimage:-string,-insertlocation:-insertlocation)"></a>insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation:InsertLocation)
将图片插入到正文中的指定位置。insertLocation 值可以为“Before”或“After”。

#### <a name="syntax"></a>语法
inlinePictureObject.insertInlinePictureFromBase64(image, insertLocation);

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|base64EncodedImage|string|必需。将 base64 编码的图像插入正文中。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
[InlinePicture](inlinepicture.md)


### <a name="insertooxml(ooxml:-string,-insertlocation:-insertlocation)"></a>insertOoxml(ooxml: string, insertLocation:InsertLocation)
在指定位置插入 OOXML。insertLocation 值可以为“Before”或“After”。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertOoxml(ooxml, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|ooxml|string|必需。要插入的 OOXML。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="insertparagraph(paragraphtext:-string,-insertlocation:-insertlocation)"></a>insertParagraph(paragraphText: string, insertLocation:InsertLocation)
在指定位置插入段落。insertLocation 值可以为“Before”或“After”。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertParagraph(paragraphText, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|paragraphText|string|必需。要插入的段落文本。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
[Paragraph](paragraph.md)

### <a name="inserttext(text:-string,-insertlocation:-insertlocation)"></a>insertText(text: string, insertLocation:InsertLocation)
将文本插入到正文中的指定位置。insertLocation 值可以为“Before”或“After”。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.insertText(text, insertLocation);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|text|string|必需。要插入的文本。|
|insertLocation|InsertLocation|必需。值可以为“Before”或“After”。|

#### <a name="returns"></a>返回
[Range](range.md)

### <a name="select(selectionmode:-selectionmode)"></a>select(selectionMode: SelectionMode)
选择图片并在 Word UI 中进行浏览。SelectionMode 值可以为“Select”、“Start”或“End”。

#### <a name="syntax"></a>语法
```js
inlinePictureObject.select(selectionMode);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|可选。选择模式可以为“Select”、“Start”或“End”。“Select”为默认值。|

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

## <a name="support-details"></a>支持详细信息
在运行时检查过程中使用[要求设置](../office-add-in-requirement-sets.md)可以确保您的应用程序受到 Word 主机版本的支持。有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序要求](../../docs/overview/requirements-for-running-office-add-ins.md)。
