# <a name="filtercriteria-object-javascript-api-for-excel"></a>FilterCriteria 对象 (Excel JavaScript API)

表示应用于列的筛选条件。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|color|string|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选一起使用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|第一个条件用于筛选数据。在“自定义”筛选中用作运算符。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|第二个条件用于筛选数据。在“自定义”筛选中仅用作运算符。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|string|filter 使用的属性，用于确定值是否应一直可见。可取值为：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|operator|string|使用“自定义”筛选器时，用于组合条件 1 和 2 的运算符。可取值为：And、Or。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|一组用于“values”筛选器的值。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|用于筛选单元格的图标。与“图标”筛选一起使用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
无

