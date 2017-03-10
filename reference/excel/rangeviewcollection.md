# <a name="rangeviewcollection-object-javascript-api-for-excel"></a>RangeViewCollection 对象 (Excel JavaScript API)

表示一组 RangeView 对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|项|[RangeView[]](rangeview.md)|一组 rangeView 对象。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|获取集合中 RangeView 对象的数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[RangeView](rangeview.md)|按索引获取 RangeView 行。从零开始编制索引。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
获取集合中 RangeView 对象的数量。

#### <a name="syntax"></a>语法
```js
rangeViewCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitematindex-number"></a>getItemAt(index: number)
按索引获取 RangeView 行。从零开始编制索引。

#### <a name="syntax"></a>语法
```js
rangeViewCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|index|number|可见行的索引。|

#### <a name="returns"></a>返回
[RangeView](rangeview.md)
