# Filter 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

管理表格列的筛选。

## 属性

无

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|条件|[FilterCriteria](filtercriteria.md)|给定列上当前应用的筛选器。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[apply(criteria:FilterCriteria)](#applycriteria-filtercriteria)|void|在给定列中应用给定的筛选条件。使用以下任一帮助程序方法均可实现相同的功能。|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|将“Bottom Item”筛选器应用于列，获取给定数量的元素。|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|将“Bottom Percent”筛选器应用于列，获取给定百分比的元素。|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|将“Cell Color”筛选器应用于列，获取给定颜色。|
|[applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|将“Icon”筛选器应用于列，获取给定条件的字符串。|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|将“Dynamic”筛选器应用于列。|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|将“Font Color”筛选器应用于列，获取给定颜色。|
|[applyIconFilter(icon:Icon)](#applyiconfiltericon-icon)|void|将“Icon”筛选器应用于列，获取给定图标。|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|将“Top Item”筛选器应用于列，获取给定数量的元素。|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|将“Top Percent”筛选器应用于列，获取给定百分比的元素。|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|将“Values”筛选器应用于列，获取给定值。|
|[clear()](#clear)|void|清除给定列上的筛选器。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### apply(criteria:FilterCriteria)
在给定列中应用给定的筛选条件。使用以下任一帮助程序方法均可实现相同的功能。 

#### 语法
```js
filterObject.apply(criteria);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|条件|FilterCriteria|要应用的条件。|

#### 返回
void

#### 示例
以下示例演示如何使用泛型 apply() 方法应用自定义筛选器。

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
        filterOn: Excel.FilterOn.custom,
        criterion1: ">50",
        operator: Excel.FilterOperator.and,
        criterion2: "<100"
        } 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
将“Bottom Item”筛选器应用于列，获取给定数量的元素。

#### 语法
```js
filterObject.applyBottomItemsFilter(count);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|count|编号|要显示的底部元素的数量。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
将“Bottom Percent”筛选器应用于列，获取给定百分比的元素。

#### 语法
```js
filterObject.applyBottomPercentFilter(percent);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|百分比|编号|要显示的底部元素的百分比。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
将“Cell Color”筛选器应用于列，获取给定颜色。


#### 语法
```js
filterObject.applyCellColorFilter(color);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|color|string|要显示的单元格的背景颜色。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper:FilterOperator)
将“Icon”筛选器应用于列，获取给定条件的字符串。

#### 语法
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|criteria1|string|第一个条件字符串。|
|criteria2|string|可选。第二个条件字符串。|
|运算符|FilterOperator|可选。说明这两个条件如何联接的运算符。|

#### 返回
void


#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
将“Dynamic”筛选器应用于列。

#### 语法
```js
filterObject.applyDynamicFilter(criteria);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|条件|string|要应用的动态条件。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember, BelowAverage、LastMonth, LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
将“Font Color”筛选器应用于列，获取给定颜色。

#### 语法
```js
filterObject.applyFontColorFilter(color);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|color|string|要显示的单元格的字体颜色。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon:Icon)
将“Icon”筛选器应用于列，获取给定图标。

#### 语法
```js
filterObject.applyIconFilter(icon);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|icon|图标|要显示的单元格图标。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
将“Top Item”筛选器应用于列，获取给定数量的元素。

#### 语法
```js
filterObject.applyTopItemsFilter(count);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|count|编号|要显示的顶部元素的数量。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
将“Top Percent”筛选器应用于列，获取给定百分比的元素。

#### 语法
```js
filterObject.applyTopPercentFilter(percent);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|百分比|编号|要显示的顶部元素的百分比。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
将“Values”筛选器应用于列，获取给定值。

#### 语法
```js
filterObject.applyValuesFilter(values);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|值|()[]|要显示的值的列表。|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
清除给定列上的筛选器。

#### 语法
```js
filterObject.clear();
```

#### 参数
无

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
    return ctx.sync(); 
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
