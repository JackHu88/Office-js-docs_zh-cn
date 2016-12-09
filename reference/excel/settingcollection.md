# <a name="settingcollection-object-javascript-api-for-excel"></a>SettingCollection 对象（适用于 Excel 的 JavaScript API）

表示一组属于工作簿的工作表对象。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|items|[Setting[]](setting.md)|一组设置对象。只读。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|按键获取设置项。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Setting](setting.md)|按键获取设置项。如果设置对象不存在，则返回的对象 isNull 属性为 true。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[set(key: string, value: string)](#setkey-string-value-string)|[Setting](setting.md)|设置指定的设置对象，或将其添加到工作簿中。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="getitemkey-string"></a>getItem(key: string)
按键获取设置项。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Key|string|设置的键。|

#### <a name="returns"></a>返回
[Setting](setting.md)

### <a name="getitemornullkey-string"></a>getItemOrNull(key: string)
按键获取设置项。如果设置对象不存在，则返回的对象 isNull 属性为 true。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.getItemOrNull(key);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Key|string|设置的键。|

#### <a name="returns"></a>返回
[Setting](setting.md)

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

### <a name="setkey-string-value-string"></a>set(key: string, value: string)
设置指定的设置对象，或将其添加到工作簿中。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.set(key, value);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Key|string|新设置的键。|
|value|string|新设置的值。|

#### <a name="returns"></a>返回
[Setting](setting.md)
