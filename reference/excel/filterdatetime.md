# <a name="filterdatetime-object-javascript-api-for-excel"></a>FilterDatetime 对象 (Excel JavaScript API)

表示在筛选值时如何筛选日期。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|date|string|用于筛选数据的采用 ISO8601 格式的日期。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|specificity|string|用于保留数据的日期的具体程度。例如，如果日期是 2005-04-02 并且将特殊性设置为“月”，则筛选操作将保留包含 2009 年 4 月日期的所有行。可能的值是：Year、Monday、Day、Hour、Minute、Second。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法
无

