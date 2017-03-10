# <a name="nameditemcollection-object-javascript-api-for-excel"></a>NamedItemCollection 对象 (Excel JavaScript API)

属于工作簿或工作表（具有取决于限定到的范围）的所有 NamedItem 对象的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|项|[NamedItem[]](nameditem.md)|namedItem 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(name: string, reference:Range or string, comment: string)](#addname-string-reference-range-or-string-comment-string)|[NamedItem](nameditem.md)|将新名称添加到给定范围的集合。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[addFormulaLocal(name: string, formula: string, comment: string)](#addformulalocalname-string-formula-string-comment-string)|[NamedItem](nameditem.md)|使用用户的公式区域设置，将新名称添加到给定范围的集合。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|获取集合中已命名项的数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|按名称获取命名项对象|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[NamedItem](nameditem.md)|按 NamedItem 对象的名称获取此对象。如果没有 NamedItem 对象，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addname-string-reference-range-or-string-comment-string"></a>add(name: string, reference:Range or string, comment: string)
将新名称添加到给定范围的集合。

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.add(name, reference, comment);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|name|string|已命名项的名称。|
|reference|Range 或 string|名称将引用的公式或区域。|
|comment|string|可选。与此已命名项相关联的注释。|

#### <a name="returns"></a>返回
[NamedItem](nameditem.md)

### <a name="addformulalocalname-string-formula-string-comment-string"></a>addFormulaLocal(name: string, formula: string, comment: string)
使用用户的公式区域设置，将新名称添加到给定范围的集合。

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.addFormulaLocal(name, formula, comment);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|name|string|已命名项的“名称”。|
|formula|string|名称将引用的采用用户区域设置的公式。|
|comment|string|可选。与此已命名项相关联的注释。|

#### <a name="returns"></a>返回
[NamedItem](nameditem.md)

### <a name="getcount"></a>getCount()
获取集合中已命名项的数量。

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemname-string"></a>getItem(name: string)
使用其名称获取 nameditem 对象

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.getItem(name);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
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
### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
按 NamedItem 对象的名称获取此对象。如果没有 NamedItem 对象，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
namedItemCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|name|string|nameditem 名称。|

#### <a name="returns"></a>返回
[NamedItem](nameditem.md)
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


