# <a name="settingcollection-object-javascript-api-for-excel"></a>SettingCollection 对象 (Excel JavaScript API)

表示属于工作簿的 Worksheet 对象的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|项|[Setting[]](setting.md)|一组设置对象。只读。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(key: string, value: (any)[])](#addkey-string-value-any)|[Setting](setting.md)|设置指定的 Setting 对象，或将其添加到工作簿中。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|获取集合中的 Setting 对象的数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(key: string)](#getitemkey-string)|[Setting](setting.md)|按键获取 Setting 项。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(key: string)](#getitemornullobjectkey-string)|[Setting](setting.md)|按键获取 Setting 项。如果没有 Setting 项，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addkey-string-value-any"></a>add(key: string, value: (any)[])
设置指定的 Setting 对象，或将其添加到工作簿中。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.add(key, value);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|string|新设置的键。|
|value|(any)[]|新设置的值。|

#### <a name="returns"></a>返回
[Setting](setting.md)

### <a name="getcount"></a>getCount()
获取集合中的 Setting 对象的数量。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemkey-string"></a>getItem(key: string)
按键获取设置项。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|string|设置的键。|

#### <a name="returns"></a>返回
[Setting](setting.md)

### <a name="getitemornullobjectkey-string"></a>getItemOrNullObject(key: string)
按键获取 Setting 项。如果没有 Setting 项，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
settingCollectionObject.getItemOrNullObject(key);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|Key|string|设置的键。|

#### <a name="returns"></a>返回
[Setting](setting.md)
