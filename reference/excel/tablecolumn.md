# <a name="tablecolumn-object-(javascript-api-for-excel)"></a>TableColumn 对象（适用于 Excel 的 JavaScript API）

代表表格中的一列。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|id|int|返回标识表内的列的唯一键。只读。|
|index|int|返回表的列集合内列的索引编号。从零开始编制索引。只读。|
|name|string|返回表格列的名称。只读。|
|values|object[][]|表示指定区域的原始值。返回的数据类型可能是字符串、数字或布尔值。包含错误的单元格将返回错误的字符串。|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|筛选器|[Filter](filter.md)|检索应用于列的筛选器。只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|从表中删除列。|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|获取与列的数据体相关的 range 对象。|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|获取与列的标头行相关的 range 对象。|
|[getRange()](#getrange)|[Range](range.md)|获取与整个列相关的 range 对象。|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|获取与列的总计行相关的 range 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="delete()"></a>delete()
从表中删除列。

#### <a name="syntax"></a>语法
```js
tableColumnObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
    column.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getdatabodyrange()"></a>getDataBodyRange()
获取与列的数据体相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableColumnObject.getDataBodyRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var dataBodyRange = column.getDataBodyRange();
    dataBodyRange.load('address');
    return ctx.sync().then(function() {
        console.log(dataBodyRange.address);
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getheaderrowrange()"></a>getHeaderRowRange()
获取与列的标头行相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableColumnObject.getHeaderRowRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var headerRowRange = columns.getHeaderRowRange();
    headerRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(headerRowRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getrange()"></a>getRange()
获取与整个列相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableColumnObject.getRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var columnRange = columns.getRange();
    columnRange.load('address');
    return ctx.sync().then(function() {
        console.log(columnRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="gettotalrowrange()"></a>getTotalRowRange()
获取与列的总计行相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableColumnObject.getTotalRowRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
    var totalRowRange = columns.getTotalRowRange();
    totalRowRange.load('address');
    return ctx.sync().then(function() {
        console.log(totalRowRange.address);
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
    var tableName = 'Table1';
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
    column.load('index');
    return ctx.sync().then(function() {
        console.log(column.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
    var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
    column.values = newValues;
    column.load('values');
    return ctx.sync().then(function() {
        console.log(column.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
