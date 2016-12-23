# <a name="shape-object-javascript-api-for-visio"></a>Shape 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_
>**注意：**Visio JavaScript API 暂处于预览阶段，可能会发生变更。暂不支持在生产环境中使用 Visio JavaScript API。

表示 Shape 类。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:---|
|id|int|形状的标识符。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-id)|
|name|string|形状的名称。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-name)|
|select|bool|如果选择形状，则返回 true。用户可以设置为 true，从而明确选择形状。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-select)|
|text|string|形状的文本。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-text)|

_请参阅属性访问 [示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 反馈|
|:---------------|:--------|:----------|:---|
|hyperlinks|[HyperlinkCollection](hyperlinkcollection.md)|返回一组形状对象超链接。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-hyperlinks)|
|shapeDataItems|[ShapeDataItemCollection](shapedataitemcollection.md)|返回形状的数据部分。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-shapeDataItems)|
|subShapes|[ShapeCollection](shapecollection.md)|获取 SubShape 集合。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-subShapes)|
|view|[ShapeView](shapeview.md)|返回形状的视图。只读。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-view)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:---|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shape-load)|

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
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shapeName = "Sample Name";
    var shape = activePage.shapes.getItem(shapeName);
    shape.load();
    return ctx.sync().then(function () {
        console.log(shape.name );
        console.log(shape.id );
        console.log(shape.Text );
        console.log(shape.Select );
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    shape.view.highlight = { color: "#E7E7E7", width: 100 };
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```