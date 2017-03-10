# <a name="pivottable-object-javascript-api-for-excel"></a>PivotTable 对象 (Excel JavaScript API)

表示 Excel 数据透视表。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|名称|string|数据透视表的名称。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|工作表|[Worksheet](worksheet.md)|包含当前 PivotTable 对象的工作表。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[refresh()](#refresh)|void|刷新数据透视表。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="refresh"></a>refresh()
刷新数据透视表。

#### <a name="syntax"></a>语法
```js
pivotTableObject.refresh();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void
