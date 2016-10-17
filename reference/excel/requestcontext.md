# <a name="requestcontext-object-(javascript-api-for-excel)"></a>RequestContext 对象（适用于 Excel 的 JavaScript API）

RequestContext 对象可加快对 Excel 应用程序的请求。由于 Office 外接程序和 Excel 应用程序在两个不同的进程中运行，需要请求上下文来获得对 Excel 及外接程序中相关对象（如工作表、表等）的访问权限。 

## <a name="properties"></a>属性
无

## <a name="methods"></a>方法

| 方法         | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |使用参数指定的属性和选项填充在 JavaScript 层中创建的代理对象。|

## <a name="api-specification"></a>API 规范

### <a name="load(object:-object,-option:-object)"></a>load(object: object, option: object)
使用参数指定的属性和选项填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
requestContextObject.load(object, loadOption);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:----------------|:--------|:----------|
|object|object|可选。指定要加载的对象的名称。|
|选项|[loadOption](loadoption.md)|可选。指定加载选项，例如选择、展开、跳过和置顶。请参阅 LoadOption 对象了解详细信息。|

#### <a name="returns"></a>返回
void

##### <a name="examples"></a>示例

下面的示例从一个区域加载属性值，然后将其复制到另一个区域。

```js
Excel.run(function (ctx) { 
    var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
    ctx.load(range, "values");
    return ctx.sync().then(function() {
        var myvalues=range.values;
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = myvalues;
        console.log(range.values);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```
