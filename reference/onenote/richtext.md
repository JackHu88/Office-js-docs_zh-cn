# <a name="richtext-object-javascript-api-for-onenote"></a>RichText 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示 Paragraph 中的 RichText 对象。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|string|获取 RichText 对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-id)|
|languageId|string|文本的语言 ID。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-languageId)|
|text|字符串|获取 RichText 对象的文本内容。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-text)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|获取包含 RichText 对象的 Paragraph 对象。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-paragraph)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getHtml()](#gethtml)|string|获取格式文本的 HTML|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-getHtml)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-richText-load)|

## <a name="method-details"></a>方法详细信息


### <a name="gethtml"></a>getHtml()
获取格式文本的 HTML

#### <a name="syntax"></a>语法
```js
richTextObject.getHtml();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
字符串

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

**id 和 text**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Get the first PageContent on the page, and then get its outline's paragraphs.
    var outlinePageContents = [];
    var paragraphs = [];
    var richTextParagraphs = [];
    // Queue a command to load the id and type of each page content in the outline.
    pageContents.load("id,type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            // Load all page contents of type Outline
            $.each(pageContents.items, function(index, pageContent) {
                if(pageContent.type == 'Outline')
                {
                    pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
                    outlinePageContents.push(pageContent);
                }
            });
            return context.sync();
        })
        .then(function () {
            // Load all rich text paragraphs across outlines
            $.each(outlinePageContents, function(index, outlinePageContent) {
                var outline = outlinePageContent.outline;
                paragraphs = paragraphs.concat(outline.paragraphs.items);
            });
            $.each(paragraphs, function(index, paragraph) {
                if(paragraph.type == 'RichText')
                {
                    richTextParagraphs.push(paragraph);
                    paragraph.load("id,richText/text");
                }
            });
            return context.sync();
        })
        .then(function () {
            // Display all rich text paragraphs to the console
            $.each(richTextParagraphs, function(index, richTextParagraph) {
                var richText = richTextParagraph.richText;
                console.log("Paragraph found with richtext content : " + richText.text + " and richtext id : " + richText.id);
            });
            return context.sync();
        });
});
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
