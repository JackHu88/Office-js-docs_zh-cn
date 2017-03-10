# <a name="bindingcollection-object-javascript-api-for-excel"></a>BindingCollection 对象 (Excel JavaScript API)

表示属于工作簿的所有 Binding 对象的集合。

## <a name="properties"></a>属性

| 属性       | 类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|count|INT|返回集合中绑定的数量。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|items|[Binding[]](binding.md)|绑定对象的集合。只读。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_请参阅属性访问[示例。](#property-access-examples)_

## <a name="relationships"></a>关系
无


## <a name="methods"></a>方法

| 方法           | 返回类型    |说明| 要求集|
|:---------------|:--------|:----------|:----|
|[add(range:Range 或 string, bindingType: string, id: string)](#addrange-range-or-string-bindingtype-string-id-string)|[Binding](binding.md)|将新的绑定添加到特定范围。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromNamedItem(name: string, bindingType: string, id: string)](#addfromnameditemname-string-bindingtype-string-id-string)|[Binding](binding.md)|根据工作簿中的命名项添加新的绑定。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[addFromSelection(bindingType: string, id: string)](#addfromselectionbindingtype-string-id-string)|[Binding](binding.md)|根据当前选择的内容添加新的绑定。|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|获取集合中的绑定数量。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|按 ID 获取绑定对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|按绑定在项数组中的位置获取此对象。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(id: string)](#getitemornullobjectid-string)|[Binding](binding.md)|按 ID 获取 Binding 对象。如果没有 Binding 对象，将返回 NULL 对象。|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="method-details"></a>方法详细信息


### <a name="addrange-range-or-string-bindingtype-string-id-string"></a>add(range:Range 或 string, bindingType: string, id: string)
将新的 binding 对象添加到特定区域。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.add(range, bindingType, id);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|range|Range or string|要将绑定绑定到的范围。可以是 Excel 范围对象，也可以是字符串。如果是字符串，必须包含完整地址，包括工作表名称|
|bindingType|string|绑定的类型。可取值为：Range、Table、Text|
|id|string|绑定的名称。|

#### <a name="returns"></a>返回
[Binding](binding.md)

### <a name="addfromnameditemname-string-bindingtype-string-id-string"></a>addFromNamedItem(name: string, bindingType: string, id: string)
根据工作簿中的命名项添加新的绑定。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.addFromNamedItem(name, bindingType, id);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|name|string|从中创建绑定的名称。|
|bindingType|string|绑定的类型。可取值为：Range、Table、Text|
|id|string|绑定的名称。|

#### <a name="returns"></a>返回
[Binding](binding.md)

### <a name="addfromselectionbindingtype-string-id-string"></a>addFromSelection(bindingType: string, id: string)
根据当前选择的内容添加新的绑定。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.addFromSelection(bindingType, id);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|bindingType|string|绑定的类型。可取值为：Range、Table、Text|
|id|string|绑定的名称。|

#### <a name="returns"></a>返回
[Binding](binding.md)

### <a name="getcount"></a>getCount()
获取集合中的绑定数量。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.getCount();
```

#### <a name="parameters"></a>参数
无

#### <a name="returns"></a>返回
int

### <a name="getitemid-string"></a>getItem(id: string)
按 ID 获取绑定对象。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.getItem(id);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|id|string|要检索的绑定对象的 ID。|

#### <a name="returns"></a>返回
[Binding](binding.md)

#### <a name="examples"></a>示例

创建表绑定以监视表中的数据更改。数据更改时，表的背景颜色将变为橙色。

```js
function addEventHandler() {
    //Create Table1
Excel.run(function (ctx) { 
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    return ctx.sync().then(function() {
             console.log("My Diet Data Inserted!");
    })
    .catch(function (error) {
             console.log(JSON.stringify(error));
    });
});
    //Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
        console.log("Action failed with error: " + asyncResult.error.message);
    }
    else {
        // If succeeded, then add event handler to the table binding.
        Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
    }
});
}
    
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
    // highlight the table in orange to indicate data has been changed.
    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
    return ctx.sync().then(function() {
            console.log("The value in this table got changed!");
    })
    .catch(function (error) {
            console.log(JSON.stringify(error));
    });
});
}

```



#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
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


### <a name="getitematindex-number"></a>getItemAt(index: number)
根据其在项目数组中的位置获取绑定对象。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.getItemAt(index);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### <a name="returns"></a>返回
[Binding](binding.md)

#### <a name="examples"></a>示例
```js
Excel.run(function (ctx) { 
    var lastPosition = ctx.workbook.bindings.count - 1;
    var binding = ctx.workbook.bindings.getItemAt(lastPosition);
    binding.load('type')
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


### <a name="getitemornullobjectid-string"></a>getItemOrNullObject(id: string)
按 ID 获取 Binding 对象。如果没有 Binding 对象，将返回 NULL 对象。

#### <a name="syntax"></a>语法
```js
bindingCollectionObject.getItemOrNullObject(id);
```

#### <a name="parameters"></a>参数
| 参数       | 类型    |说明|
|:---------------|:--------|:----------|:---|
|id|string|要检索的绑定对象的 ID。|

#### <a name="returns"></a>返回
[Binding](binding.md)
### <a name="property-access-examples"></a>属性访问示例

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('items');
    return ctx.sync().then(function() {
        for (var i = 0; i < bindings.items.length; i++)
        {
            console.log(bindings.items[i].id);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
获取绑定的数目

```js
Excel.run(function (ctx) { 
    var bindings = ctx.workbook.bindings;
    bindings.load('count');
    return ctx.sync().then(function() {
        console.log("Bindings: Count= " + bindings.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
