# <a name="tablesort-object-javascript-api-for-excel"></a>TableSort 对象（适用于 Excel 的 JavaScript API）

管理对表对象的排序操作。

## <a name="properties"></a>属性

| 属性     | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|matchCase|bool|表示最后一次对表进行排序时大小写是否有影响。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|method|string|表示最后一次对表排序所使用的中文字符排序方法。只读。可能的值是：PinYin、StrokeCount。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="relationships"></a>关系
| 关系 | 类型   |说明| 要求集|
|:---------------|:--------|:----------|:----|
|fields|[SortField](sortfield.md)|表示最后一次对表排序所使用的当前条件。只读。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[apply(fields:SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|执行排序操作。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[clear()](#clear)|void|清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|无效|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[reapply()](#reapply)|void|对表重新应用当前的排序参数。|[1.2](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="applyfields-sortfield-matchcase-bool-method-string"></a>apply(fields:SortField[], matchCase: bool, method: string)
执行排序操作。

#### <a name="syntax"></a>语法
```js
tableSortObject.apply(fields, matchCase, method);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|域|SortField[]|要用作排序依据的条件列表。|
|matchCase|bool|可选。是否让大小写对字符串排序产生影响。|
|方法|string|可选。用于中文字符的排序方法。可能的值是：PinYin、StrokeCount|

#### <a name="returns"></a>返回
void

#### <a name="examples"></a>示例
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

### <a name="clear"></a>clear()
清除表上的当前排序。尽管这不能修改表的排序，但它会清除标题按钮的状态。

#### <a name="syntax"></a>语法
```js
tableSortObject.clear();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="loadparam-object"></a>load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### <a name="syntax"></a>语法
```js
object.load(param);
```

#### <a name="parameters"></a>参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|:---|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### <a name="returns"></a>返回
void

### <a name="reapply"></a>reapply()
对表重新应用当前的排序参数。

#### <a name="syntax"></a>语法
```js
tableSortObject.reapply();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void
