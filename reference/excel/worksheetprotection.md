# <a name="worksheetprotection-object-javascript-api-for-excel"></a>WorksheetProtection 对象 (Excel JavaScript API)

表示对 Worksheet 对象的保护。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|protected|bool|指明 worksheet 是否受保护。只读。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|选项|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|工作表保护选项。只读。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoptions)|void|保护 worksheet。如果 worksheet 处于受保护状态，则无法执行此方法。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[Unprotect](#unprotect)|void|解除对工作表的保护。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="protectoptions-worksheetprotectionoptions"></a>protect(options:WorksheetProtectionOptions)
保护 worksheet。如果 worksheet 处于受保护状态，则无法执行此方法。

#### <a name="syntax"></a>语法
```js
worksheetProtectionObject.protect(options);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|options|WorksheetProtectionOptions|可选。工作表保护选项。|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    var range = sheet.getRange("A1:B3").format.protection.locked = false;
    sheet.protection.protect({allowInsertRows:true});
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

```
### <a name="unprotect"></a>unprotect()
解除对工作表的保护。

#### <a name="syntax"></a>语法
```js
worksheetProtectionObject.unprotect();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void
