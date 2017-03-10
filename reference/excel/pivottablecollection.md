# <a name="pivottablecollection-object-javascript-api-for-excel"></a>PivotTableCollection 对象 (Excel JavaScript API)

表示属于工作簿或工作表的所有 PivotTable 对象的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|项|[PivotTable[]](pivottable.md)|一组 PivotTable 对象。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|获取集合中的数据透视表的数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[PivotTable](pivottable.md)|按名称获取数据透视表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[PivotTable](pivottable.md)|按 PivotTable 对象的名称获取此对象。如果没有 PivotTable 对象，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[refreshAll()](#refreshall)|void|刷新集合中的所有数据透视表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
获取集合中的数据透视表的数量。

#### <a name="syntax"></a>语法
```js
pivotTableCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemname-string"></a>getItem(name: string)
按名称获取数据透视表。

#### <a name="syntax"></a>语法
```js
pivotTableCollectionObject.getItem(name);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|name|string|要检索的数据透视表的名称。|

#### <a name="returns"></a>返回
[PivotTable](pivottable.md)

### <a name="getitemornullobjectname-string"></a>getItemOrNullObject(name: string)
按 PivotTable 对象的名称获取此对象。如果没有 PivotTable 对象，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
pivotTableCollectionObject.getItemOrNullObject(name);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|name|string|要检索的数据透视表的名称。|

#### <a name="returns"></a>返回
[PivotTable](pivottable.md)

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
