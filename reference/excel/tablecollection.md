# <a name="tablecollection-object-javascript-api-for-excel"></a>TableCollection 对象 (Excel JavaScript API)

表示属于工作簿或工作表（具有取决于限定到的范围）的所有表的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|count|int|返回工作簿中的表数目。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Table[]](table.md)|table 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(address: [object, hasHeaders: bool)](#addaddress-object-hasheaders-bool)|[Table](table.md)|新建表。范围对象或源地址决定了在哪个工作表下添加表。如果无法添加表（例如，由于地址无效，或者表与另一个表重叠），则会引发错误。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|获取集合中的表数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|按名称或 ID 获取表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|按表在集合中的位置获取此对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number or string)](#getitemornullobjectkey-number-or-string)|[Table](table.md)|按名称或 ID 获取表。如果没有表，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addaddress-object-hasheaders-bool"></a>add(address: [object, hasHeaders: bool)
新建表。范围对象或源地址决定了在哪个工作表下添加表。如果无法添加表（例如，由于地址无效，或者表与另一个表重叠），则会引发错误。

#### <a name="syntax"></a>语法
```js
tableCollectionObject.add(address, hasHeaders);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|address|[object|Range 对象或表示数据源的区域的字符串地址或名称。如果地址不包含工作表名称，将使用当前活动的工作表。对于 1.1，使用 string 参数；对于 1.3，也可以接受 Range 对象。|
|hasHeaders|bool|指示导入的数据是否具有列标签的布尔值。如果源不包含标头（即，当此属性设置为 false 时），Excel 将自动生成标头，数据将向下移动一行。|

#### <a name="returns"></a>返回
[Table](table.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
    table.load('name');
    return ctx.sync().then(function() {
        console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getcount"></a>getCount()
获取集合中的表数量。

#### <a name="syntax"></a>语法
```js
tableCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
按名称或 ID 获取表。

#### <a name="syntax"></a>语法
```js
tableCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string|要检索的表的名称或 ID。|

#### <a name="returns"></a>返回
[Table](table.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
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
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitematindex-number"></a>getItemAt(index: number)
根据其在集合中的位置获取表。

#### <a name="syntax"></a>语法
```js
tableCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[Table](table.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number or string)
按名称或 ID 获取表。如果没有表，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
tableCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string|要检索的表的名称或 ID。|

#### <a name="returns"></a>返回
[Table](table.md)
### <a name="property-access-examples"></a>属性访问示例

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load();
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

获取表的数目

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```