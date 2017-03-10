# <a name="boundingbox-object-javascript-api-for-visio"></a>BoundingBox 对象 (Visio JavaScript API)

适用于：_Visio Online_

表示形状的 BoundingBox 对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明|
|:---------------|:--------|:----------|
|height|int|形状边界框的上下边缘之间的距离，不包括与形状相关联的任何数据图形。|
|width|int|形状边界框的左右边缘之间的距离，不包括与形状相关联的任何数据图形。|
|x|int|指定边界框的 X 坐标的整数。|
|y|int|指定边界框的 Y 坐标的整数。|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


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
