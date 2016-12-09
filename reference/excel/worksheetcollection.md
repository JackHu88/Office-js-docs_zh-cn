# <a name="worksheetcollection-object-javascript-api-for-excel"></a>WorksheetCollection 对象（适用于 Excel 的 JavaScript API）

表示属于工作簿的 worksheet 对象的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|items|[Worksheet[]](worksheet.md)|worksheet 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|获取工作簿中当前处于活动状态的工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|按名称或 ID 获取工作表对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Worksheet](worksheet.md)|按名称或 ID 获取工作表对象。如果工作表对象不存在，则返回的对象 isNull 属性为 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addname-string"></a>add(name: string)
向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.add(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|name|string|可选。要添加的工作表的名称。如果指定，名称应唯一。如果未指定，Excel 将确定新工作表的名称。|

#### <a name="returns"></a>返回
[Worksheet](worksheet.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getactiveworksheet"></a>getActiveWorksheet()
获取工作簿中当前处于活动状态的工作表。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Worksheet](worksheet.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemkey-string"></a>getItem(key: string)
使用其名称或 ID 获取 worksheet 对象。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Key|string|工作表的名称或 ID。|

#### <a name="returns"></a>返回
[Worksheet](worksheet.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: string)
按名称或 ID 获取工作表对象。如果工作表对象不存在，则返回的对象 isNull 属性为 true。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Key|string|工作表的名称或 ID。|

#### <a name="returns"></a>返回
[Worksheet](worksheet.md)

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
    var worksheets = ctx.workbook.worksheets;
    worksheets.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < worksheets.items.length; i++)
        {
            console.log(worksheets.items[i].name);
            console.log(worksheets.items[i].index);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
