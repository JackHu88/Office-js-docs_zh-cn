# <a name="selectionchangedeventargs-object-javascript-api-for-visio"></a>SelectionChangedEventArgs 对象 (Visio JavaScript API)

适用于：_Visio Online_

提供有关引发了 SelectionChanged 事件的形状集合的信息。

## <a name="properties"></a>属性

| 属性       | 类型    |说明
|:---------------|:--------|:----------|
|shapeNames|string[]|获取引发了 SelectionChanged 事件的形状名称数组。|
|pageName|string|获取页面名称，其中包含引发了 SelectionChanged 事件的 ShapeCollection 对象。|

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
             eventResult1 = document1.onSelectionChanged.add(
        function (args){
                   console.log("Selected Shape Name: "+args.shapeNames[0]);
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
