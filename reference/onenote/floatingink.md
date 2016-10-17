# <a name="floatingink-object-(javascript-api-for-onenote)"></a>FloatingInk 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


表示一组墨迹笔划。

## <a name="properties"></a>属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|id|string|获取 FloatingInk 对象的 ID。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-id)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|inkStrokes|[InkStrokeCollection](inkstrokecollection.md)|获取 FloatingInk 对象的笔划。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-inkStrokes)|
|pageContent|[PageContent](pagecontent.md)|获取 FloatingInk 对象的 PageContent 父级。只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-pageContent)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-floatingInk-load)|

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

**id**
```js
OneNote.run(function(context) {

    // Gets the active page.
    var page = context.application.getActivePage();
    var contents = page.contents;
    
    // Load page contents and their types.
    page.load('contents/type');
    return context.sync()
        .then(function(){
        
            // Load every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    content.load('ink/id');
                }                           
            })
            return context.sync();
        })
        .then(function(){
        
            // Log ID of every ink content.
            $.each(contents.items, function(i, content) {
                if (content.type == "Ink")
                {
                    console.log(content.ink.id);
                }                           
            })              
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```
