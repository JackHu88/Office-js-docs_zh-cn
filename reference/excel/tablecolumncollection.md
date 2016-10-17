# <a name="tablecolumncollection-object-(javascript-api-for-excel)"></a>TableColumnCollection 对象（适用于 Excel 的 JavaScript API）

表示属于表的所有列的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|count|INT|返回表中的列数。只读。|
|items|[TableColumn[]](tablecolumn.md)|tableColumn 对象的集合。只读。|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[add(index: number, values: (boolean or string or number)[][])](#addindex-number-values-boolean-or-string-or-number)|[TableColumn](tablecolumn.md)|向表中添加新列。|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[TableColumn](tablecolumn.md)|按名称或 ID 获取 column 对象。|
|[getItemAt(index: number)](#getitematindex-number)|[TableColumn](tablecolumn.md)|根据其在集合中的位置获取列。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="add(index:-number,-values:-(boolean-or-string-or-number)[][])"></a>add(index: number, values: (boolean or string or number)[][])
向表中添加新列。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.add(index, values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|指定新列的相对位置。之前位于此位置的列向右移动。索引值应等于或小于最后一列的索引值，因此不能用于在表末尾附加列。从零开始编制索引。|
|values|(boolean or string or number)[][]|可选。未设置格式的表列值的二维数组。|

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


### <a name="getitem(key:-number-or-string)"></a>getItem(key: number or string)
按名称或 ID 获取 column 对象。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|Key|number or string| 列名称或 ID。|

#### <a name="returns"></a>返回
[TableColumn](tablecolumn.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
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

### <a name="getitemat(index:-number)"></a>getItemAt(index: number)
根据其在集合中的位置获取列。

#### <a name="syntax"></a>语法
```js
tableColumnCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
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
    var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
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
