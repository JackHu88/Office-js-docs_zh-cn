# <a name="tablecolumncollection-object-javascript-api-for-excel"></a>TableColumnCollection 对象 (Excel JavaScript API)

表示属于表的所有列的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|count|INT|返回表中的列数。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[TableColumn[]](tablecolumn.md)|tableColumn 对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(index: number, values: (boolean or string or number)[][], name: string)](#addindex-number-values-boolean-or-string-or-number-name-string)|[TableColumn](tablecolumn.md)|向表中添加新列。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|获取表中的列数。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|按名称或 ID 获取列对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|按列在集合中的位置获取此对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: number or string)](#getitemornullobjectkey-number-or-string)|[TableColumn](tablecolumn.md)|按名称或 ID 获取 column 对象。如果没有 column 对象，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addindex-number-values-boolean-or-string-or-number-name-string"></a>add(index: number, values: (boolean or string or number)[][], name: string)
向表中添加新列。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.add(index, values, name);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|index|number|可选。指定新列的相对位置。如果为 NULL 或 -1，将在末尾进行添加。索引更高的列将被移到一侧。从零开始编制索引。|
|值|(boolean or string or number)[][]|可选。未设置格式的表列值的二维数组。|
|name|string|可选。指定新列的名称。如果为 Null，将使用默认名称。|

#### <a name="returns"></a>返回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = tables.getItem("Table1").columns.add(null, values);
    column.load('name');
    return ctx.sync().then(function() {
        console.log(column.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getcount"></a>getCount()
获取表中的列数。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
按名称或 ID 获取 column 对象。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string| 列名称或 ID。|

#### <a name="returns"></a>返回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem('Table1').columns.getItem(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
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
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitematindex-number"></a>getItemAt(index: number)
根据其在集合中的位置获取列。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
    tablecolumn.load('name');
    return ctx.sync().then(function() {
            console.log(tablecolumn.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemornullobjectkey-number-or-string"></a>getItemOrNullObject(key: number or string)
按名称或 ID 获取 column 对象。如果没有 column 对象，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string| 列名称或 ID。|

#### <a name="returns"></a>返回
[TableColumn](tablecolumn.md)
### <a name="property-access-examples"></a>属性访问示例

```js
Excel.run(function (ctx) { 
    var tablecolumns = ctx.workbook.tables.getItem('Table1').columns;
    tablecolumns.load('items');
    return ctx.sync().then(function() {
        console.log("tablecolumns Count: " + tablecolumns.count);
        for (var i = 0; i < tablecolumns.items.length; i++)
        {
            console.log(tablecolumns.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```