# <a name="shapeview-object-javascript-api-for-visio"></a>ShapeView 对象（适用于 Visio 的 JavaScript API）

适用于：_Visio Online_
>**注意：**目前 Visio JavaScript API 不适用于预览版或生产环境。

表示 ShapeView 类。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
无

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:---|
|[addOverlay(OverlayType:OverlayType, Content: string, HorizontalAlignment:HorizontalAlignment, VerticalAlignment:VerticalAlignment, Width: number, Height: number)](#addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number)|int|在形状之上添加覆盖。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-addOverlay)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-load)|
|[removeOverlay(OverlayId: number)](#removeoverlayoverlayid-number)|void|删除形状上的特定覆盖或所有覆盖。|[转到反馈页](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-shapeView-removeOverlay)|

## <a name="method-details"></a>方法详细信息


### <a name="addoverlayoverlaytype-overlaytype-content-string-horizontalalignment-horizontalalignment-verticalalignment-verticalalignment-width-number-height-number"></a>addOverlay(OverlayType:OverlayType, Content: string, HorizontalAlignment:HorizontalAlignment, VerticalAlignment:VerticalAlignment, Width: number, Height: number)
在形状之上添加覆盖。

#### <a name="syntax"></a>语法
```js
shapeViewObject.addOverlay(OverlayType, Content, HorizontalAlignment, VerticalAlignment, Width, Height);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|OverlayType|OverlayType|覆盖类型：文本、图像。|
|Content|string|覆盖的内容。|
|HorizontalAlignment|HorizontalAlignment|覆盖的水平对齐方式：左对齐、居中对齐、右对齐|
|VerticalAlignment|VerticalAlignment|覆盖的垂直对齐方式：顶端对齐、中间对齐、底端对齐|
|Width|number|覆盖宽度。|
|Height|number|覆盖高度。|

#### <a name="returns"></a>返回
int

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

### <a name="removeoverlayoverlayid-number"></a>removeOverlay(OverlayId: number)
删除形状上的特定覆盖或所有覆盖。

#### <a name="syntax"></a>语法
```js
shapeViewObject.removeOverlay(OverlayId);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|OverlayId|number|覆盖 ID。删除形状上特定 ID 的覆盖。|

#### <a name="returns"></a>返回
void

### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
    var activePage = ctx.document.getActivePage();
    var shape = activePage.shapes.getItem(0);
    var overlayId=shape.view.addOverlay(1, "Visio Online", 2, 2, 50, 50);
    return ctx.sync();
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
    shape.view.removeOverlay(1);
    return ctx.sync();
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
