# <a name="commentcollection-object-javascript-api-for-visio"></a>CommentCollection 对象 (Visio JavaScript API)

适用于：_Visio Online_

表示给定形状的 CommentCollection 对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明
|:---------------|:--------|:----------|
|items|[Comment[]](comment.md)|一组 comment 对象。只读。|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|获取 Comment 对象的数量。|
|[getItem(key: string)](#getitemkey-string)|[Comment](comment.md)|按 Comment 对象的名称获取此对象。|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
获取 Comment 对象的数量。

#### <a name="syntax"></a>语法
```js
CommentCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemkey-string"></a>getItem(key: string)
按 Comment 对象的名称获取此对象。

#### <a name="syntax"></a>语法
```js
CommentCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|
|Key|string|键是要检索的 Comment 的名称。|

#### <a name="returns"></a>返回
[Comment](comment.md)

### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
```js
 Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Position Belt.41";
    var shape = activePage.shapes.getItem(shapeName);
    var shapecomments= shape.comments;
        shapecomments.load();
        return ctx.sync().then(function () {
             for(var i=0; i<shapecomments.items.length;i++)
        {
                    var comment= shapecomments.items[i];
            console.log("comment Author: " + comment.author);
            console.log("Comment Text: " + comment.text);
            console.log("Date " + comment.date);
        }
     });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
