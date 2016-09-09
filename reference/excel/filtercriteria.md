# FilterCriteria 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示应用于列的筛选条件。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|color|string|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选一起使用。|
|criterion1|string|第一个条件用于筛选数据。在“自定义”筛选中用作运算符。|
|criterion2|string|第二个条件用于筛选数据。在“自定义”筛选中仅用作运算符。|
|dynamicCriteria|string|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|
|filterOn|string|筛选器使用的属性，用于确定是否应将值保持为可见时。可能的值是：    BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom |
|values|object[]|要用作“值”筛选的一部分的值集。|

## 关系
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|icon|[图标](icon.md)|用于筛选单元格的图标。与“图标”筛选一起使用。|
|运算符|[FilterOperator](filteroperator.md)|使用“自定义”筛选时，用于组合条件 1 和 2 的运算符。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


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
