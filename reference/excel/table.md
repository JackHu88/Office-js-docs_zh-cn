# Table 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示一个 Excel 表。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|id|int|返回用于唯一标识指定工作簿中表的值。即使表被重命名，标识符的值仍然相同。只读。|
|name|string|表的名称。|
|showHeaders|bool|指示标头行是否可见。该值可以设置为显示或删除标头行。|
|showTotals|bool|指示总计行是否可见。该值可以设置为显示或删除总计行。|
|style|string|表示表格样式的常量值。可能的值是：TableStyleLight1 through TableStyleLight21、TableStyleMedium1 through TableStyleMedium28、TableStyleStyleDark1 through TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|

_请参阅属性访问[示例](#示例)。_

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|columns|[TableColumnCollection](tablecolumncollection.md)|表示表中所有列的集合。只读。|
|rows|[TableRowCollection](tablerowcollection.md)|表示表中所有行的集合。只读。|
|排序|[TableSort](tablesort.md)|表示表的排序配置。只读。|
|工作表|[工作表](worksheet.md)|包含当前表的工作表。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[clearFilters()](#clearfilters)|void|清除当前应用于表的所有筛选器。|
|[convertToRange()](#converttorange)|[Range 对象设置内联图片](range.md)|将表转换为普通单元格区域。保留所有数据。|
|[delete()](#delete)|void|删除表。|
|[getDataBodyRange()](#getdatabodyrange)|[Range 对象设置内联图片](range.md)|获取与表的数据体相关的 range 对象。|
|[getHeaderRowRange()](#getheaderrowrange)|[Range 对象设置内联图片](range.md)|获取与表的标头行相关的 range 对象。|
|[getRange()](#getrange)|[Range 对象设置内联图片](range.md)|获取与整个表相关的 range 对象。|
|[getTotalRowRange()](#gettotalrowrange)|[Range 对象设置内联图片](range.md)|获取与表的总计行相关的 range 对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[reapplyFilters()](#reapplyfilters)|void|重新应用当前表上的所有筛选器。|

## 方法详细信息


### clearFilters()
清除当前应用于表的所有筛选器。

#### 语法
```js
tableObject.clearFilters();
```

#### 参数
无

#### 返回
void

### convertToRange()
将表转换为普通单元格区域。保留所有数据。

#### 语法
```js
tableObject.convertToRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.convertToRange();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### delete()
删除表。

#### 语法
```js
tableObject.delete();
```

#### 参数
无

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getDataBodyRange()
获取与表的数据体相关的 range 对象。

#### 语法
```js
tableObject.getDataBodyRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableDataRange = table.getDataBodyRange();
    tableDataRange.load('address')
    return ctx.sync().then(function() {
            console.log(tableDataRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getHeaderRowRange()
获取与表的标题行相关的 range 对象。

#### 语法
```js
tableObject.getHeaderRowRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableHeaderRange = table.getHeaderRowRange();
    tableHeaderRange.load('address');
    return ctx.sync().then(function() {
        console.log(tableHeaderRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange()
获取与整个表相关的 range 对象。

#### 语法
```js
tableObject.getRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItem(tableName);
    var tableRange = table.getRange();
    tableRange.load('address'); 
    return ctx.sync().then(function() {
            console.log(tableRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getTotalRowRange()
获取与表的总计行相关的 range 对象。

#### 语法
```js
tableObject.getTotalRowRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    var tableTotalsRange = table.getTotalRowRange();
    tableTotalsRange.load('address');   
    return ctx.sync().then(function() {
            console.log(tableTotalsRange.address);
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

按名称获取表。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.load('index')
    return ctx.sync().then(function() {
            console.log(table.index);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

按索引获取表。

```js
Excel.run(function (ctx) { 
    var index = 0;
    var table = ctx.workbook.tables.getItemAt(0);
    table.name('name')
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

设置表格样式。 

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.name = 'Table1-Renamed';
    table.showTotals = false;
    table.tableStyle = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.tableStyle);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
