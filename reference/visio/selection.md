# <a name="selection-object-javascript-api-for-visio"></a>Selection 对象 (Visio JavaScript API)

适用于：_Visio Online_

表示页面中的 Selection 对象。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
| 关系 | 类型    |说明|
|:---------------|:--------|:----------|
|shapes|[ShapeCollection](shapecollection.md)|获取 Selection 对象的形状，只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
