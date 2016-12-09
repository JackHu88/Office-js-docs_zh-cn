# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>RangeViewCollection 对象（适用于 Excel 的 JavaScript API）

表示一组属于工作簿的工作表对象。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|items|[RangeView[]](rangeview.md)|一组 rangeView 对象。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|按索引获取 RangeView 行。从零开始编制索引。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getitematindex-number"></a>getItemAt(index: number)
按索引获取 RangeView 行。从零开始编制索引。

#### <a name="syntax"></a>语法
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|index|number|可见行的索引。|

#### <a name="returns"></a>返回
[RangeView](rangeview.md)

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
