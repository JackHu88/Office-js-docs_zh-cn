# <a name="sortfield-object-javascript-api-for-excel"></a>SortField 对象 (Excel JavaScript API)

表示排序操作中的条件。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|ascending|bool|表示是否执行升序排序。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|表示按字体或单元格颜色进行排序时，条件的目标颜色。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|dataOption|string|表示此字段的其他排序选项。可能的值是：Normal、TextAsNumber。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|Key|int|表示条件所在的列（或行，具体取决于排序方向）。表示与第一列（或行）的偏移量。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|sortOn|string|表示此条件的排序类型。可能的值是：Value、CellColor、FontColor、Icon。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|icon|[Icon](icon.md)|表示对单元格图标进行排序时，条件的目标图标。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法
无

