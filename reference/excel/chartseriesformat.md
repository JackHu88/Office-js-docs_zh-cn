# <a name="chartseriesformat-object-javascript-api-for-excel"></a>ChartSeriesFormat 对象（适用于 Excel 的 JavaScript API）

封装图表系列的格式属性

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|fill|[ChartFill](chartfill.md)|表示图表系列的填充格式，包括背景格式信息。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|line|[ChartLineFormat](chartlineformat.md)|表示线条格式。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

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
