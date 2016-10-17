# <a name="inkanalysisline-object-(javascript-api-for-onenote)"></a>InkAnalysisLine 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_   


表示墨迹笔划形成的已识别文本行的墨迹分析数据。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|string|获取 InkAnalysisLine 对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-id)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|paragraph|[InkAnalysisParagraph](inkanalysisparagraph.md)|对父级 InkAnalysisParagraph 的引用。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-paragraph)|
|words|[InkAnalysisWordCollection](inkanalysiswordcollection.md)|获取此墨迹分析行中的墨迹分析字词。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-words)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisLine-load)|

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

**words**
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
                    // Word counts in a line.
                    console.log(inkLine.words.items.length);
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