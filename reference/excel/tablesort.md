# TableSort 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

管理对 Table 对象的排序操作。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|matchCase|bool|表示最后一次对表进行排序时大小写是否有影响。只读。|
|方法|string|表示最后一次对表排序所使用的中文字符排序方法。只读。可能的值是：PinYin、StrokeCount。|

## Relationships
| 关系 | 类型   |说明|
|:---------------|:--------|:----------|
|域|[SortField](sortfield.md)|表示最后一次对表排序所使用的当前条件。只读。|

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[apply(fields:SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|执行排序操作。|
|[clear()](#clear)|void|清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|
|[reapply()](#reapply)|void|对表重新应用当前的排序参数。|

## 方法详细信息


### apply(fields:SortField[], matchCase: bool, method: string)
执行排序操作。

#### 语法
```js
tableSortObject.apply(fields, matchCase, method);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|域|SortField[]|要用作排序依据的条件列表。|
|matchCase|bool|可选。是否让大小写对字符串排序产生影响。|
|方法|string|可选。用于中文字符的排序方法。可能的值是：PinYin、StrokeCount|

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。

#### 语法
```js
tableSortObject.clear();
```

#### 参数
无

#### 返回
void

#### 示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### Syntax
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

### reapply()
对表重新应用当前的排序参数。

#### 语法
```js
tableSortObject.reapply();
```

#### 参数
无

#### 返回
void

####示例
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.reapply();   
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});