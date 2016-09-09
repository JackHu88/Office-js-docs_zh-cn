# WorksheetProtectionOptions 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

代表工作表保护中的选项。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|allowAutoFilter|bool|表示允许使用自动筛选功能的工作表保护选项。|
|allowDeleteColumns|bool|表示允许删除列的工作表保护选项。|
|allowDeleteRows|bool|表示允许删除行的工作表保护选项。|
|allowFormatCells|bool|表示允许格式化单元格的工作表保护选项。|
|allowFormatColumns|bool|表示允许格式化列的工作表保护选项。|
|allowFormatRows|bool|表示允许格式化行的工作表保护选项。|
|allowInsertColumns|bool|表示允许插入列的工作表保护选项。|
|allowInsertHyperlinks|bool|表示允许插入超链接的工作表保护选项。|
|allowInsertRows|bool|表示允许插入行的工作表保护选项。|
|allowPivotTables|bool|表示允许使用数据透视表功能的工作表保护选项。|
|allowSort|bool|表示允许使用排序功能的工作表保护选项。|

_请参阅属性访问[示例](#示例)。_

## Relationships
无


## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

#### 示例
本示例加载活动工作表的保护选项。
```js
Excel.run(function (ctx) {
    var worksheet = ctx.workbook.worksheets.getActiveWorksheet();
    worksheet.protection.load();            
    return ctx.sync()
        .then(function () {
            console.log("Active worksheet's protection options: " + worksheet.protection.options);
        });
})
.catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
