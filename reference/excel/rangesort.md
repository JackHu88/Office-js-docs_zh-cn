# <a name="rangesort-object-javascript-api-for-excel"></a>RangeSort 对象（适用于 Excel 的 JavaScript API）

管理对范围对象的排序操作。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|void|执行排序操作。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string"></a>apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)
执行排序操作。

#### <a name="syntax"></a>语法
```js
rangeSortObject.apply(fields, matchCase, hasHeaders, orientation, method);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|域|SortField[]|要用作排序依据的条件列表。|
|matchCase|bool|可选。是否让大小写对字符串排序产生影响。|
|hasHeaders|bool|可选。该区域是否有标头。|
|orientation|string|可选。该操作是对行还是列排序。可能的值是：Rows、Columns|
|方法|string|可选。用于中文字符的排序方法。可能的值是：PinYin、StrokeCount|

#### <a name="returns"></a>返回
void
