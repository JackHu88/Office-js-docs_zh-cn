# <a name="nameditem-object-javascript-api-for-excel"></a>NamedItem 对象 (Excel JavaScript API)

表示单元格区域或值的定义名称。名称可以为基元的已命名对象（如以下类型中所示）、range 对象或对区域的引用。此对象可用于获取与名称相关的 range 对象。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|comment|string|表示与此名称相关联的注释。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|对象的名称。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|scope|string|指明是否将 name 限定到工作簿或特定工作表。只读。可取值为：Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|type|string|指明 name 公式返回的值的类型。只读。可能的值是：String、Integer、Double、Boolean、Range。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|表示 name 公式计算出的值。对于已命名区域，将返回区域地址。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|visible|bool|指定对象是否可见。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
| 关系 | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|工作表|[Worksheet](worksheet.md)|返回已命名项限定到的工作表。如果项改为限定到工作簿，将引发错误。只读。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|worksheetOrNullObject|[Worksheet](worksheet.md)|返回已命名项限定到的工作表。如果项改为限定到工作簿，将返回 NULL 对象。只读。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|删除给定的名称。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange()](#getrange)|[Range](range.md)|返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将引发错误。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getRangeOrNullObject()](#getrangeornullobject)|[Range](range.md)|返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="delete"></a>delete()
删除给定的名称。

#### <a name="syntax"></a>语法
```js
namedItemObject.delete();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
void

### <a name="getrange"></a>getRange()
返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将引发错误。

#### <a name="syntax"></a>语法
```js
namedItemObject.getRange();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)

#### <a name="examples"></a>示例

返回与此名称相关的范围对象。如果名称的类型不是 `Range`，则为 `null`。备注:此 API 当前仅支持工作簿范围的项目。**

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var range = names.getItem('MyRange').getRange();
    range.load('address');
    return ctx.sync().then(function() {
            console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### <a name="getrangeornullobject"></a>getRangeOrNullObject()
返回与名称相关联的 Range 对象。如果已命名项的类型不是 Range，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
namedItemObject.getRangeOrNullObject();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
[Range](range.md)
### <a name="property-access-examples"></a>属性访问示例

```js
Excel.run(function (ctx) { 
    var names = ctx.workbook.names;
    var namedItem = names.getItem('MyRange');
    namedItem.load('type');
    return ctx.sync().then(function() {
            console.log(namedItem.type);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
