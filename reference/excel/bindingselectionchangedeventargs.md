# <a name="bindingselectionchangedeventargs-object-javascript-api-for-excel"></a>BindingSelectionChangedEventArgs 对象 (Excel JavaScript API)

提供有关引发了 SelectionChanged 事件的绑定的信息。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|columnCount|int|获取选择的列数。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|获取选择的行数。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startColumn|int|获取选定区域第一列的索引（从零开始编制索引）。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startRow|int|获取选定区域第一行的索引（从零开始编制索引）。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|binding|[Binding](binding.md)|获取表示引发了 SelectionChanged 事件的绑定的 Binding 对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
无

