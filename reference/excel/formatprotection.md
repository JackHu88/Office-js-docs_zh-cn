# <a name="formatprotection-object-(javascript-api-for-excel)"></a>FormatProtection 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示对范围对象的格式保护。

## <a name="properties"></a>属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|formulaHidden|bool|表示 Excel 是否隐藏区域中的单元格公式。指示整个区域不具有统一公式隐藏设置的空值。|
|locked|bool|指示 Excel 是否将对象中的单元格锁定。指示整个区域不具有统一锁定设置的空值。|

_请参阅属性访问 [示例](#property-access-examples)_。

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="load(param:-object)"></a>load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
