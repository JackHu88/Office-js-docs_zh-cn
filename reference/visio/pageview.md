# <a name="pageview-object-javascript-api-for-visio"></a>PageView 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_
>**注意：**Visio JavaScript API 暂处于预览阶段，可能会发生变更。暂不支持在生产环境中使用 Visio JavaScript API。

表示 PageView 类。

## <a name="properties"></a>属性

| 属性 | 类型 |说明| 反馈|
|:---------------|:--------|:----------|:---|
|zoom|int|获取并设置页面的缩放级别。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-zoom)|

## <a name="relationships"></a>关系

无

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:---|
|[centerViewportOnShape(ShapeId: number)](#centerviewportonshapeshapeid-number)|void|平移 Visio 绘图，将指定的形状放置在视图中心位置。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-centerViewportOnShape)|
|[fitToWindow()](#fittowindow)|void|让页面填满当前窗口。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-fitToWindow)|
|[isShapeInViewport(Shape:Shape)](#isshapeinviewportshape-shape)|bool|检查形状是否在页面的视区内。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-isShapeInViewport)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageView-load)|

## <a name="method-details"></a>方法详细信息


### <a name="centerviewportonshapeshapeid-number"></a>centerViewportOnShape(ShapeId: number)
平移 Visio 绘图，将指定的形状放置在视图中心位置。

#### <a name="syntax"></a>语法
```js
pageViewObject.centerViewportOnShape(ShapeId);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|ShapeId|number|在中心位置显示的形状的 ID。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    activePage.view.centerViewportOnShape(shape.Id);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="fittowindow"></a>fitToWindow()
让页面填满当前窗口。

#### <a name="syntax"></a>语法
```js
pageViewObject.fitToWindow();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="isshapeinviewportshape-shape"></a>isShapeInViewport(Shape:Shape)
检查形状是否在页面的视区内。

#### <a name="syntax"></a>语法
```js
pageViewObject.isShapeInViewport(Shape);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Shape|Shape|要检查的形状。|

#### <a name="returns"></a>返回
bool

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

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|Position|Position|指定页面在视图中的新位置的位置对象。|

#### <a name="returns"></a>返回
void
### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    activePage.view.zoom = 300;
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

