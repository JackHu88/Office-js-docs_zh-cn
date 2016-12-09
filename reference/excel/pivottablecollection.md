# <a name="pivottablecollection-object-javascript-api-for-excel"></a>PivotTableCollection 对象（适用于 Excel 的 JavaScript API）

表示属于工作簿或工作表的所有数据透视表的集合。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|items|[PivotTable[]](pivottable.md)|一组数据透视表对象。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|按名称获取数据透视表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(name: string)](#getitemornullname-string)|[PivotTable](pivottable.md)|按名称获取数据透视表。如果数据透视表对象不存在，则返回的对象 isNull 属性为 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|刷新集合中的所有数据透视表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getitemname-string"></a>getItem(name: string)
按名称获取数据透视表。

#### <a name="syntax"></a>语法
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|name|string|要检索的数据透视表的名称。|

#### <a name="returns"></a>返回
[PivotTable](pivottable.md)

### <a name="getitemornullname-string"></a>getItemOrNull(name: string)
按名称获取数据透视表。如果数据透视表对象不存在，则返回的对象 isNull 属性为 true。

#### <a name="syntax"></a>语法
```js
pivotTableCollectionObject.getItemOrNull(name);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|name|string|要检索的数据透视表的名称。|

#### <a name="returns"></a>返回
[PivotTable](pivottable.md)

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

### <a name="refreshall"></a>refreshAll()
刷新集合中的所有数据透视表。

#### <a name="syntax"></a>语法
```js
pivotTableCollectionObject.refreshAll();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void
