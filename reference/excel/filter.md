# <a name="filter-object-javascript-api-for-excel"></a>Filter 对象（适用于 Excel 的 JavaScript API）

管理表中列的筛选。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|criteria|[FilterCriteria](filtercriteria.md)|给定列上当前应用的筛选器。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[apply(criteria:FilterCriteria)](#applycriteria-filtercriteria)|void|在给定列中应用给定的筛选条件。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|将“Bottom Item”筛选器应用于列，以获取给定数量的元素。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|将“Bottom Percent”筛选器应用于列，以获取给定比例的元素。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|将“Cell Color”筛选器应用于列，以获取给定颜色。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: string)](#applycustomfiltercriteria1-string-criteria2-string-oper-string)|void|将“Icon”筛选器应用于列，以获取给定的条件字符串。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|将“Dynamic”筛选器应用于列。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|将“Font Color”筛选器应用于列，以获取给定颜色。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyIconFilter(icon:Icon)](#applyiconfiltericon-icon)|void|将“Icon”筛选器应用于列，以获取给定图标。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|将“Top Item”筛选器应用于列，以获取给定数量的元素。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|将“Top Percent”筛选器应用于列，以获取给定比例的元素。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|将“Values”筛选器应用于列，以获取给定值。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|清除给定列上的筛选器。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="applycriteria-filtercriteria"></a>apply(criteria:FilterCriteria)
在给定列中应用给定的筛选条件。

#### <a name="syntax"></a>语法
```js
filterObject.apply(criteria);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|条件|FilterCriteria|要应用的条件。|

#### <a name="returns"></a>返回
void

### <a name="applybottomitemsfiltercount-number"></a>applyBottomItemsFilter(count: number)
将“Bottom Item”筛选器应用于列，获取给定数量的元素。

#### <a name="syntax"></a>语法
```js
filterObject.applyBottomItemsFilter(count);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|count|编号|要显示的底部元素的数量。|

#### <a name="returns"></a>返回
void

### <a name="applybottompercentfilterpercent-number"></a>applyBottomPercentFilter(percent: number)
将“Bottom Percent”筛选器应用于列，获取给定百分比的元素。

#### <a name="syntax"></a>语法
```js
filterObject.applyBottomPercentFilter(percent);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|百分比|编号|要显示的底部元素的百分比。|

#### <a name="returns"></a>返回
void

### <a name="applycellcolorfiltercolor-string"></a>applyCellColorFilter(color: string)
将“Cell Color”筛选器应用于列，获取给定颜色。

#### <a name="syntax"></a>语法
```js
filterObject.applyCellColorFilter(color);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|color|string|要显示的单元格的背景颜色。|

#### <a name="returns"></a>返回
void

### <a name="applycustomfiltercriteria1-string-criteria2-string-oper-string"></a>applyCustomFilter(criteria1: string, criteria2: string, oper: string)
将“Icon”筛选器应用于列，以获取给定的条件字符串。

#### <a name="syntax"></a>语法
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|criteria1|string|第一个条件字符串。|
|criteria2|string|可选。第二个条件字符串。|
|oper|string|可选。说明这两个条件如何联接的运算符。可取值为：And、Or|

#### <a name="returns"></a>返回
void

### <a name="applydynamicfiltercriteria-string"></a>applyDynamicFilter(criteria: string)
将“Dynamic”筛选器应用于列。

#### <a name="syntax"></a>语法
```js
filterObject.applyDynamicFilter(criteria);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|条件|string|要应用的动态条件。可能的值是：Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember, BelowAverage、LastMonth, LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday|

#### <a name="returns"></a>返回
void

### <a name="applyfontcolorfiltercolor-string"></a>applyFontColorFilter(color: string)
将“Font Color”筛选器应用于列，获取给定颜色。

#### <a name="syntax"></a>语法
```js
filterObject.applyFontColorFilter(color);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|color|string|要显示的单元格的字体颜色。|

#### <a name="returns"></a>返回
void

### <a name="applyiconfiltericon-icon"></a>applyIconFilter(icon:Icon)
将“Icon”筛选器应用于列，获取给定图标。

#### <a name="syntax"></a>语法
```js
filterObject.applyIconFilter(icon);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|icon|图标|要显示的单元格图标。|

#### <a name="returns"></a>返回
void

### <a name="applytopitemsfiltercount-number"></a>applyTopItemsFilter(count: number)
将“Top Item”筛选器应用于列，获取给定数量的元素。

#### <a name="syntax"></a>语法
```js
filterObject.applyTopItemsFilter(count);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|count|编号|要显示的顶部元素的数量。|

#### <a name="returns"></a>返回
void

### <a name="applytoppercentfilterpercent-number"></a>applyTopPercentFilter(percent: number)
将“Top Percent”筛选器应用于列，获取给定百分比的元素。

#### <a name="syntax"></a>语法
```js
filterObject.applyTopPercentFilter(percent);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|百分比|编号|要显示的顶部元素的百分比。|

#### <a name="returns"></a>返回
void

### <a name="applyvaluesfiltervalues-"></a>applyValuesFilter(values: ()[])
将“Values”筛选器应用于列，获取给定值。

#### <a name="syntax"></a>语法
```js
filterObject.applyValuesFilter(values);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|值|()[]|要显示的值的列表。|

#### <a name="returns"></a>返回
void

### <a name="clear"></a>clear()
清除给定列上的筛选器。

#### <a name="syntax"></a>语法
```js
filterObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

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
