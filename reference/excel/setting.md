# <a name="setting-object-javascript-api-for-excel"></a>Setting 对象 (Excel JavaScript API)

Setting 表示文档保留设置的键值对。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|Key|string|返回表示设置对象的 ID 的键。只读。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|表示为此设置存储的值。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|删除 Setting 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="delete"></a>delete()
删除设置对象。

#### <a name="syntax"></a>语法
```js
settingObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void
