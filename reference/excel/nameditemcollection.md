# <a name="nameditemcollection-object-(javascript-api-for-excel)"></a>NamedItemCollection 对象（适用于 Excel 的 JavaScript API）

属于工作簿的所有 nameditem 对象的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|项目|[NamedItem[]](nameditem.md)|namedItem 对象的集合。只读。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|使用其名称获取 nameditem 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="getitem(name:-string)"></a>getItem(name: string)
使用其名称获取 nameditem 对象

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|name|string|nameditem 名称。|

#### <a name="returns"></a>返回
[NamedItem](nameditem.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var nameditem = ctx.workbook.names.getItem(wSheetName);
    nameditem.load('type');
    return ctx.sync().then(function() {
            console.log(nameditem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var nameditem = ctx.workbook.names.getItemAt(0);
    nameditem.load('name');
    return ctx.sync().then(function() {
            console.log(nameditem.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="load(param:-object)"></a>load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

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

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < nameditems.items.length; i++)
        {
            console.log(nameditems.items[i].name);
            console.log(nameditems.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

获取 nameditem 的数目。

```js
Excel.run(function (ctx) { 
    var nameditems = ctx.workbook.names;
    nameditems.load('count');
    return ctx.sync().then(function() {
        console.log("nameditems: Count= " + nameditems.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

