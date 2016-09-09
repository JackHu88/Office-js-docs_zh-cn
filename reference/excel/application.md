# Application 对象（适用于 Excel 的 JavaScript API）

表示用于管理工作簿的 Excel 应用程序。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|calculationMode|string|返回工作簿中使用的计算模式。只读。可能的值是：`Automatic` Excel 控制重新计算，`AutomaticExceptTables` Excel 控制重新计算，但忽略表中的更改，`Manual` 在用户请求时执行计算。|

_请参阅属性访问[示例](#示例)。_

## Relationships
无


## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|重新计算 Excel 中当前打开的所有工作簿。|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### calculate(calculationType: string)
重新计算 Excel 中当前打开的所有工作簿。

#### 语法
```js
applicationObject.calculate(calculationType);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|calculationType|string|指定要使用的计算类型。可能的值是：`Recalculate` 默认选项，通过计算工作簿中的所有公式执行正常计算，`Full` 强制执行数据的完整计算，`FullRebuild` 强制执行数据的完整计算并重新构建依存关系。|

#### 返回
void

#### 示例
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


### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者接受 [loadOption](loadoption.md) 对象。|

#### 返回
void
### 属性访问示例
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

