# <a name="datarefreshcompleteeventargs-object-javascript-api-for-visio"></a>DataRefreshCompleteEventArgs 对象 (Visio JavaScript API)

适用于：_Visio Online_

提供有关引发了 DataRefreshComplete 事件的文档的信息。

## <a name="properties"></a>属性

| 属性       | 类型    |说明
|:---------------|:--------|:----------|
|success|bool|获取 DataRefreshComplete 事件的 successfailure。|
|document|[Document](document.md)|获取有关引发了 DataRefreshComplete 事件的 Document 对象。|

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
         eventResult1 = document1.onDataRefreshComplete.add(
    function (args){
           console.log("Data Refresh Result: "+args.success);
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
