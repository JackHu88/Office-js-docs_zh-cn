# WorksheetProtection 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示对工作表对象的保护。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|受保护|bool|表示该工作表是否受保护。只读。|

## 关系
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|选项|[WorksheetProtectionOptions](worksheetprotectionoptions.md)|工作表保护选项。只读。|

## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用工作表的保护详细信息填充代理对象。|
|[protect(options:WorksheetProtectionOptions)](#protectoptions-worksheetprotectionoption)|void|保护工作表。如果工作表处于受保护状态，则会引发它。|
|[unprotect()](#unprotect)|void|解除工作表保护|

## 方法详细信息


### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

#### 示例
本示例加载活动工作表保护的详细信息。
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection status: " + worksheet.protection.protected);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### protect(options:WorksheetProtectionOptions)
通过可选的保护策略保护工作表。如果工作表处于受保护状态，则引发异常。 

当指定选项时，可在启用或禁用之间切换各个策略。如果不指定策略，则默认情况下为启用。 

#### 语法
```js
worksheetProtectionObject.protect(options);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|选项|WorksheetProtectionOptions|可选。工作表保护选项。|


#### 返回
void

#### 示例
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
### unprotect()
解除工作表保护。 

#### 语法
```js
worksheetProtectionObject.unprotect();
```

#### 参数
无

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
	var sheet = ctx.workbook.worksheets.getItem("Sheet1");	
	sheet.protection.unprotect();
	return ctx.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
