# <a name="shapecollection-object-javascript-api-for-visio"></a>ShapeCollection 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_

>**注意：**目前 Visio JavaScript API 不适用于预览版或生产环境。

表示 ShapeCollection。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:---|
|items|[Shape[]](shape.md)|一组形状对象。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-items)|

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:---|
|[getCount()](#getcount)|int|获取集合中的形状数量。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getCount)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Shape](shape.md)|按键（名称或索引）获取形状。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-getItem)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeCollection-load)|

## <a name="method-details"></a>方法详细信息


### <a name="getcount"></a>getCount()
获取集合中的形状数量。

#### <a name="syntax"></a>语法
```js
shapeCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

#### <a name="examples"></a>示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var numShapesActivePage = activePage.shapes.getCount();
    return ctx.sync().then(function () {
        console.log("Shapes Count: " + numShapesActivePage.value);
    });

}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="getitemkey-number-or-string"></a>getItem(key: number or string)
按键（名称或索引）获取形状。

#### <a name="syntax"></a>语法
```js
shapeCollectionObject.getItem(key);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Key|number or string|键是要检索的形状的名称或索引。|

#### <a name="returns"></a>返回
[Shape](shape.md)

### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
