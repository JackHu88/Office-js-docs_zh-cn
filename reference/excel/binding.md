# Binding 对象（适用于 Excel 的 JavaScript API）

表示工作簿中定义的 Office.js 绑定。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|id|string|表示绑定标识符。只读。|
|type|string|返回绑定的类型。只读。可能的值是：Range、Table、Text。|

_请参阅属性访问[示例](#示例)。_

## Relationships
无


## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range 对象设置内联图片](range.md)|返回绑定表示的区域。如果绑定类型不正确，将引发错误。|
|[getTable()](#gettable)|[Table](table.md)|返回绑定表示的表。如果绑定类型不正确，将引发错误。|
|[getText()](#gettext)|string|返回绑定表示的文本。如果绑定类型不正确，将引发错误。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### getRange()
返回绑定表示的区域。如果绑定类型不正确，将引发错误。

#### 语法
```js
bindingObject.getRange();
```

#### 参数
无

#### 返回
[Range 对象设置内联图片](range.md)

#### 示例
以下示例使用绑定对象获取相关区域。

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getTable()
返回绑定表示的表。如果绑定类型不正确，将引发错误。

#### 语法
```js
bindingObject.getTable();
```

#### 参数
无

#### 返回
[Table](table.md)

#### 示例
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
    table.load('name');
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getText()
返回绑定表示的文本。如果绑定类型不正确，将引发错误。

#### 语法
```js
bindingObject.getText();
```

#### 参数
无

#### 返回
string

#### 示例

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    ctx.load('text');
    return ctx.sync().then(function() {
        console.log(text);
    });
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
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
    return ctx.sync().then(function() {
        console.log(binding.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
