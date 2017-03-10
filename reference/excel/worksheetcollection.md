# <a name="worksheetcollection-object-javascript-api-for-excel"></a>WorksheetCollection 对象 (Excel JavaScript API)

表示属于工作簿的 Worksheet 对象的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|项|[Worksheet[]](worksheet.md)|worksheet 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|获取工作簿中当前处于活动状态的工作表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount(visibleOnly: bool)](#getcountvisibleonly-bool)|int|获取集合中的工作表数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|按名称或 ID 获取工作表对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Worksheet](worksheet.md)|按 Worksheet 对象的名称或 ID 获取此对象。如果没有 Worksheet 对象，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addname-string"></a>add(name: string)
向工作簿添加新工作表。工作表将添加到现有工作表的末尾。如果您想要激活新添加的工作表，请对其调用 ".activate()。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.add(name);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
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


### <a name="getcountvisibleonly-bool"></a>getCount(visibleOnly: bool)
获取集合中的工作表数量。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.getCount(visibleOnly);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|visibleOnly|bool|可选。如果设置为 true，将仅返回可见的工作表。 |

#### <a name="returns"></a>返回
int

### <a name="getitemkey-string"></a>getItem(key: string)
使用其名称或 ID 获取 worksheet 对象。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|string|工作表的名称或 ID。|

#### <a name="returns"></a>返回
[Worksheet](worksheet.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
按 Worksheet 对象的名称或 ID 获取此对象。如果没有 Worksheet 对象，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
worksheetCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|string|工作表的名称或 ID。|

#### <a name="returns"></a>返回
[Worksheet](worksheet.md)
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
