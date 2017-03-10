# <a name="rangeview-object-javascript-api-for-excel"></a>RangeView 对象 (Excel JavaScript API)

RangeView 表示父区域的一组可见单元格。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|cellAddresses|object[][]|表示 RangeView 的单元格地址。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|columnCount|int|返回可见列数。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulas|object[][]|表示采用 A1 表示法的公式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasLocal|object[][]|表示采用 A1 样式表示法的公式，使用用户的语言和数字格式区域设置。例如，英语中的公式 "=SUM(A1, 1.5)" 在德语中将变为 "=SUMME(A1; 1,5)"。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|formulasR1C1|object[][]|表示采用 R1C1 表示法的公式。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|index|int|返回表示 RangeView 的索引的值。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|获取一个值，该值指定|表示 Excel 中指定单元格的数字格式代码。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|返回可见行数。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|text|对象[][]|指定区域的文本值。文本值与单元格宽度无关。在 Excel UI 中替代 # 符号不会影响 API 返回的文本值。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|valueTypes|string|表示每个单元格的数据类型。只读。可能的值是：Unknown、Empty、String、Integer、Double、Boolean、Error。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|values|object[][]|表示指定的 RangeView 的原始值。返回的数据可能是字符串、数字，也可能是布尔值。包含错误的单元格将返回错误字符串。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|rows|[RangeViewCollection](rangeviewcollection.md)|表示一组与 range 相关联的 RangeView。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getRange()](#getrange)|[Range](range.md)|获取与当前 RangeView 相关联的父范围。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getrange"></a>getRange()
获取与当前 RangeView 相关联的父范围。

#### <a name="syntax"></a>语法
```js
rangeViewObject.getRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)
