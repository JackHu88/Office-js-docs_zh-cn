# <a name="inkanalysisword-object-(javascript-api-for-onenote)"></a>InkAnalysisWord 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示墨迹笔划形成的已识别字词的墨迹分析数据。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|string|获取 InkAnalysisWord 对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-id)|
|languageId|string|此 inkAnalysisWord 中已识别语言的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-languageId)|
|wordAlternates|string|按照可能性的顺序，已在此墨迹字词中识别的字词。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-wordAlternates)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|line|[InkAnalysisLine](inkanalysisline.md)|对父级 InkAnalysisLine 的引用。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-line)|
|strokePointers|[InkStrokePointer](inkstrokepointer.md)|对已识别为此墨迹分析字词一部分的墨迹笔划的弱引用。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-strokePointers)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisWord-load)|

## <a name="method-details"></a>方法详细信息


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

**wordAlternates 和 languageId**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    page.load('inkAnalysisOrNull/paragraphs/lines/words');
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            $.each(inkParagraphs.items, function(i, inkParagraph) {
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function(j, inkLine) {
                    var inkWords = inkLine.words;
                    $.each(inkWords.items, function(k, inkWord) {
                    
                        // Log language Id of the word
                        console.log(inkWord.languageId);
                        
                        // Log every ink analyzed words.
                        $.each(inkWord.wordAlternates, function(l, word) {
                            console.log(word);                                  
                        })
                    })
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```