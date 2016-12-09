# <a name="filtercriteria-object-javascript-api-for-excel"></a>FilterCriteria 对象（适用于 Excel 的 JavaScript API）

表示应用于列的筛选条件。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|color|string|用于筛选单元格的 HTML 颜色字符串。与“cellColor”和“fontColor”筛选器结合使用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion1|string|第一个条件用于筛选数据。在“custom”筛选器中用作运算符。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|criterion2|string|第二个条件用于筛选数据。在“custom”筛选器中仅用作运算符。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dynamicCriteria|string|Excel.DynamicFilterCriteria 集中的动态条件将应用于此列。与“动态”筛选一起使用。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|filterOn|string|filter 使用的属性，用于确定值是否应一直可见。可取值为：BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|operator|string|使用“custom”筛选器时，用于组合条件 1 和 2 的运算符。可取值为：And、Or。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[]|一组用于“values”筛选器的值。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|用于筛选单元格的图标。与“icon”筛选器结合使用。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


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
