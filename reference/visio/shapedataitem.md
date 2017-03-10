# <a name="shapedataitem-object-javascript-api-for-visio"></a>ShapeDataItem 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_

表示 ShapeDataItem 对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明|
|:---------------|:--------|:----------|
|format|string|指定形状数据项格式的字符串。只读。|
|formattedValue|string|指定形状数据项格式值的字符串。只读。|
|label|string|指定形状数据项标签的字符串。只读。|
|value|string|指定形状数据项值的字符串。只读。|

_请参阅属性访问 [示例。](#property-access-examples)_

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
### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
        var shapeDataItem = shape.shapeDataItems.getItem(0);
    shapeDataItem.load();
        return ctx.sync().then(function() {
                console.log(shapeDataItem.label);
                console.log(shapeDataItem.value);
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
