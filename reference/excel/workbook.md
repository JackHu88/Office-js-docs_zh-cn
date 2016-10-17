# <a name="workbook-object-(javascript-api-for-excel)"></a>Workbook 对象（适用于 Excel 的 JavaScript API）

workbook 是包含相关 workbook 对象的顶级对象，例如 worksheet、table、range 等。

## <a name="properties"></a>属性

无

## <a name="relationships"></a>Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|application|[Application](application.md)|表示包含此工作簿的 Excel 应用程序实例。只读。|
|bindings|[BindingCollection](bindingcollection.md)|表示属于工作簿的绑定的集合。只读。|
|functions|[Functions](functions.md)|表示包含此工作簿的 Excel 应用程序实例。只读。|
|names|[NamedItemCollection](nameditemcollection.md)|表示工作簿范围内的已命名项目（称为区域和常量）的集合。只读。|
|tables|[TableCollection](tablecollection.md)|表示与工作簿关联的表的集合。只读。|
|Worksheets|[WorksheetCollection](worksheetcollection.md)|表示与工作簿关联的工作表的集合。只读。|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|从工作簿中获取当前选定的区域。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## <a name="method-details"></a>方法详细信息


### <a name="getselectedrange()"></a>getSelectedRange()
从工作簿中获取当前选定的区域。

#### <a name="syntax"></a>语法
```js
workbookObject.getSelectedRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### <a name="load(param:-object)"></a>load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void
