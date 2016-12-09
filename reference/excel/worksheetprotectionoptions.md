# <a name="worksheetprotectionoptions-object-javascript-api-for-excel"></a>WorksheetProtectionOptions 对象（适用于 Excel 的 JavaScript API）

表示工作表保护选项。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|allowAutoFilter|bool|表示允许使用自动筛选功能的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteColumns|bool|表示允许删除列的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowDeleteRows|bool|表示允许删除行的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatCells|bool|表示允许格式化单元格的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatColumns|bool|表示允许格式化列的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowFormatRows|bool|表示允许格式化行的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertColumns|bool|表示允许插入列的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertHyperlinks|bool|表示允许插入超链接的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowInsertRows|bool|表示允许插入行的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowPivotTables|bool|表示允许使用数据透视表功能的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|allowSort|bool|表示允许使用排序功能的工作表保护选项。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


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
