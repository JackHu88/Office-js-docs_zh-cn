# <a name="shapemouseentereventargs-object-javascript-api-for-visio"></a>ShapeMouseEnterEventArgs 对象 (Visio JavaScript API)

适用于：_Visio Online_

提供有关引发了 MouseEnter 事件的形状的信息。

## <a name="properties"></a>属性

| 属性       | 类型    |说明
|:---------------|:--------|:----------|
|shapeName|string|获取引发了 MouseEnter 事件的 Shape 对象的名称。|
|pageName|string|获取页面名称，其中包含引发了 MouseEnter 事件的 Shape 对象。|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无

## <a name="methods"></a>方法
无

### <a name="property-access-examples"></a>属性访问示例
```js
Visio.run(function (ctx) { 
  var document1= ctx.document;
               var page = document1.getActivePage();
    eventResult2 = document1.onMouseEnter.add(
            function (args){            
                         console.log(Date.now()+":OnMouseEnter Event"+JSON.stringify(args));
            });
    return ctx.sync().then(function () {
           console.log("Success");
        });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```