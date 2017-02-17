# <a name="hyperlinkcollection-object-javascript-api-for-visio"></a>HyperlinkCollection 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_
>**注意：**Visio JavaScript API 暂处于预览阶段，可能会发生变更。暂不支持在生产环境中使用 Visio JavaScript API。

表示 HyperlinkCollection。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 反馈|
|:---------------|:--------|:----------|:---|
|items|[Hyperlink[]](hyperlink.md)|一组超链接对象。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-items)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|获取超链接的数量。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getCount)|
|[getItem(Key: number or string)](#getitemkey-number-or-string)|[Hyperlink](hyperlink.md)|按键（名称或 ID）获取超链接。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-getItem)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-hyperlinkCollection-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
获取超链接的数量。

#### <a name="syntax"></a>语法
```js
hyperlinkCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemkey-number-or-string"></a>getItem(Key: number or string)
按键（名称或 ID）获取超链接。

#### <a name="syntax"></a>语法
```js
hyperlinkCollectionObject.getItem(Key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string|键是要检索的超链接的名称或索引。|

#### <a name="returns"></a>返回
[Hyperlink](hyperlink.md)

### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Manager Belt";
    var shape = activePage.shapes.getItem(shapeName);
    var hyperlinks = shape.hyperlinks;
    shapeHyperlinks.load();
        ctx.sync().then(function () {
            for(var i=0; i<shapeHyperlinks.items.length;i++)
                {
                  var hyperlink = shapeHyperlinks.items[i];
                  console.log("Description:"+hyperlink.description +"Address:"+hyperlink.address +"SubAddress:  "+ hyperlink.subAddress);
                }

            });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
