# <a name="nameditemcollection-object-javascript-api-for-excel"></a>NamedItemCollection 对象（适用于 Excel 的 JavaScript API）

属于工作簿的所有 nameditem 对象的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|items|[NamedItem[]](nameditem.md)|namedItem 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|按名称获取命名项对象|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[NamedItem](nameditem.md)|按名称获取命名项对象。如果命名项对象不存在，则返回的对象 isNull 属性为 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getitemname-string"></a>getItem(name: string)
使用其名称获取 nameditem 对象

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|name|string|nameditem 名称。|

#### <a name="returns"></a>返回
[NamedItem](nameditem.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var sheetName = 'Sheet1';
    var nameditem = ctx.workbook.names.getItem(sheetName);
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
### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
按名称获取命名项对象。如果命名项对象不存在，则返回的对象 isNull 属性为 true。

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|name|string|nameditem 名称。|

#### <a name="returns"></a>返回
[NamedItem](nameditem.md)

### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

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


