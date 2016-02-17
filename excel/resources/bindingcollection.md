# BindingCollection 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

表示属于工作簿的所有绑定对象的集合。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|count|int|返回集合中绑定的数量。只读。|
|Items|[Binding[]](binding.md)|绑定对象的集合。只读。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
无


## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[Binding](binding.md)|按 ID 获取绑定对象。|
|[getItemAt(index: number)](#getitematindex-number)|[Binding](binding.md)|根据其在项目数组中的位置获取绑定对象。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息

### getItem(id: string)
按 ID 获取绑定对象。

#### 语法
```js
bindingCollectionObject.getItem(id);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|id|string|要检索的绑定对象的 ID。|

#### 返回
[Binding](binding.md)

#### 示例

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
		// If successful, add the event handler to the table binding.
		Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
	}
});
}
	
// When data in the table is changed, this event is triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
	// Highlight the table in orange to indicate data changed.
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



#### 示例
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

### getItemAt(index: number)
根据其在项目数组中的位置获取绑定对象。

#### 语法
```js
bindingCollectionObject.getItemAt(index);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### 返回
[Binding](binding.md)

#### 示例
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

### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者接受 [loadOption](loadoption.md) 对象。|

#### 返回
void

### 属性访问示例

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < bindings.items.length; i++)
		{
			console.log(bindings.items[i].id);
			console.log(bindings.items[i].index);
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

