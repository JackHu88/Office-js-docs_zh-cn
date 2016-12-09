# <a name="table-object-javascript-api-for-excel"></a>Table 对象（适用于 Excel 的 JavaScript API）

表示 Excel 表。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|highlightFirstColumn|bool|指明第一列是否包含特殊格式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|highlightLastColumn|bool|指明最后一列是否包含特殊格式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|id|int|返回用于唯一标识指定工作簿中表的值。即使表被重命名，标识符的值仍然相同。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|表的名称。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedColumns|bool|指明列是否采用镶边格式来以不同的方式突出显示奇数列与偶数列，让表更易于阅读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showBandedRows|bool|指明行是否采用镶边格式来以不同的方式突出显示奇数行与偶数行，让表更易于阅读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showFilterButton|bool|指明是否在每个列标题的顶部显示筛选器按钮。仅当 table 中包含标题行时，才允许设定此设置。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|showHeaders|bool|指示标头行是否可见。该值可以设置为显示或删除标题行。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|showTotals|bool|指示总计行是否可见。该值可以设置为显示或删除总计行。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|style|string|表示表格样式的常量值。可能的值是：TableStyleLight1 thru TableStyleLight21、TableStyleMedium1 thru TableStyleMedium28、TableStyleStyleDark1 thru TableStyleStyleDark11。还可以指定工作簿中显示的用户定义的自定义样式。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|columns|[TableColumnCollection](tablecolumncollection.md)|表示表中所有列的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rows|[TableRowCollection](tablerowcollection.md)|表示表中所有行的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|sort|[TableSort](tablesort.md)|表示表的排序。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|worksheet|[Worksheet](worksheet.md)|包含当前表格的工作表。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[clearFilters()](#clearfilters)|void|清除当前在表中应用的所有筛选器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[convertToRange()](#converttorange)|[Range](range.md)|将表转换为普通单元格区域。保留所有数据。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|删除表。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|获取与表的数据主体相关联的范围对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|获取与表的标题行相关联的范围对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|获取与整个表相关联的范围对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|获取与表的总计行相关联的范围对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapplyFilters()](#reapplyfilters)|void|重新应用当前在表上应用的所有筛选器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="clearfilters"></a>clearFilters()
清除当前应用于表的所有筛选器。

#### <a name="syntax"></a>语法
```js
tableObject.clearFilters();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="converttorange"></a>convertToRange()
将表转换为普通单元格区域。保留所有数据。

#### <a name="syntax"></a>语法
```js
tableObject.convertToRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
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

### <a name="delete"></a>delete()
删除表。

#### <a name="syntax"></a>语法
```js
tableObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
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


### <a name="getdatabodyrange"></a>getDataBodyRange()
获取与表的数据体相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableObject.getDataBodyRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
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

### <a name="getheaderrowrange"></a>getHeaderRowRange()
获取与表的标头行相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableObject.getHeaderRowRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
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


### <a name="getrange"></a>getRange()
获取与整个表相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableObject.getRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
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


### <a name="gettotalrowrange"></a>getTotalRowRange()
获取与表的总计行相关的 range 对象。

#### <a name="syntax"></a>语法
```js
tableObject.getTotalRowRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例
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

### <a name="reapplyfilters"></a>reapplyFilters()
重新应用当前在表上应用的所有筛选器。

#### <a name="syntax"></a>语法
```js
tableObject.reapplyFilters();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例

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
    table.load('id')
    return ctx.sync().then(function() {
            console.log(table.id);
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
    table.style = 'TableStyleMedium2';
    table.load('tableStyle');
    return ctx.sync().then(function() {
            console.log(table.style);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
