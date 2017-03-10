# <a name="pagecollection-object-javascript-api-for-visio"></a>PageCollection 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_

表示属于文档的 Page 对象的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明|
|:---------------|:--------|:----------|
|项|[Page[]](page.md)|页面对象的集合。只读。|

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getCount()](#getcount)|int|获取集合中的页面数量。|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Page](page.md)|按键（名称或 ID）获取页面。|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
获取集合中的页面数量。

#### <a name="syntax"></a>语法
```js
pageCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
按键（名称或 ID）获取页面。

#### <a name="syntax"></a>语法
```js
pageCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string|键是要检索的页面的名称或 ID。|

#### <a name="returns"></a>返回
[Page](page.md)

#### <a name="examples"></a>示例
```js
Visio.run(function (ctx) { 
    var pageName = 'Page-1';
    var page = ctx.document.pages.getItem(pageName);
    page.activate();
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

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
