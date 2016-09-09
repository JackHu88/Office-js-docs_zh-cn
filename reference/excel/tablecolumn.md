# TableColumn 对象（适用于 Excel 的 JavaScript API）

代表表格中的一列。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|id|int|返回标识表内的列的唯一键。只读。|
|index|int|返回表的列集合内列的索引编号。从零开始编制索引。只读。|
|name|string|返回表格列的名称。只读。|
|values|object[][]|表示指定区域的原始值。返回的数据类型可能是字符串、数字或布尔值。包含错误的单元格将返回错误的字符串。|

_请参阅属性访问[示例](#示例)。_

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|筛选器|[筛选](filter.md)|检索应用于列的筛选器。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|从表中删除列。|
|[getDataBodyRange()](#getdatabodyrange)|[Range 对象设置内联图片](range.md)|获取与列的数据体相关的 range 对象。|
|[getHeaderRowRange()](#getheaderrowrange)|[Range 对象设置内联图片](range.md)|获取与列的标头行相关的 range 对象。|
|[getRange()](#getrange)|[Range 对象设置内联图片](range.md)|获取与整个列相关的 range 对象。|
|[getTotalRowRange()](#gettotalrowrange)|[Range 对象设置内联图片](range.md)|获取与列的总计行相关的 range 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### delete()
从表中删除列。

#### 语法
```js
tableColumnObject.delete();
```

#### 参数
无

#### 返回
void

#### 示例

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


### getDataBodyRange()
获取与列的数据体相关的 range 对象。

#### 语法
```js
tableColumnObject.getDataBodyRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

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


### getHeaderRowRange()
获取与列的标头行相关的 range 对象。

#### 语法
```js
tableColumnObject.getHeaderRowRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

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

### getRange()
获取与整个列相关的 range 对象。

#### 语法
```js
tableColumnObject.getRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

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


### getTotalRowRange()
获取与列的总计行相关的 range 对象。

#### 语法
```js
tableColumnObject.getTotalRowRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例

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


### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void
### 属性访问示例

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
