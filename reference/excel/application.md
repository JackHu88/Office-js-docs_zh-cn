# <a name="application-object-javascript-api-for-excel"></a>Application 对象 (Excel JavaScript API)

表示用于管理工作簿的 Excel 应用程序。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|calculationMode|string|返回工作簿中使用的计算模式。只读。可取值为：`Automatic`：Excel 控制重新计算；`AutomaticExceptTables`：Excel 控制重新计算，但忽略表中的更改；`Manual`：在用户请求时执行计算。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|无效|重新计算 Excel 中当前打开的所有工作簿。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|


## <a name="method-details"></a>方法详细信息


### <a name="calculatecalculationtype-string"></a>calculate(calculationType: string)
重新计算 Excel 中当前打开的所有工作簿。

#### <a name="syntax"></a>语法
```js
applicationObject.calculate(calculationType);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|calculationType|string|指定要使用的计算类型。可取值为：`Recalculate`：重新计算 Excel 已标记为脏的所有单元格（即易失或已更改的数据的从属单元格），以及以编程方式标记为脏的单元格。`Full`：这会将所有单元格都标记为脏，然后重新计算所有单元格。`FullRebuild`：这会强制重建整个计算链，将所有单元格都标记为脏，然后重新计算所有单元格。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) {
    ctx.workbook.application.calculate('Full');
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
Excel.run(function (ctx) {
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

