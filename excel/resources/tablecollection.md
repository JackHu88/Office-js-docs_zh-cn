# TableCollection 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

表示属于工作簿的所有表的集合。

## 属性

| 属性   | 类型|说明
|:---------------|:--------|:----------|
|count|int|返回工作簿中的表数目。只读。|
|Items|[Table[]](table.md)|table 对象的集合。只读。|

_请参阅属性访问[示例](#property-access-examples)。_

## Relationships
无


## 方法

| 方法   | 返回类型|说明|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Table](table.md)|创建一个新表。区域源地址确定将在其下添加表的工作表。如果无法添加表（例如，由于地址无效，或者表与另一个表重叠），将抛出错误。|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Table](table.md)|按名称或 ID 获取表。|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|根据其在集合中的位置获取表。|
|[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息

### add(address: string, hasHeaders: bool)
创建一个新表。区域源地址确定将在其下添加表的工作表。如果无法添加表（例如，由于地址无效，或者表与另一个表重叠），将抛出错误。

#### 语法
```js
tableCollectionObject.add(address, hasHeaders);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|address|string|表示数据源的 range 对象的地址或名称。如果该地址不包含工作表名称，则使用当前活动的工作表。|
|hasHeaders|bool|指示导入的数据是否具有列标签的布尔值。如果源不包含标头（例如，当此属性设置为 false 时），Excel 将自动生成标头，数据将向下移动一行。|

#### 返回
[Table](table.md)

#### 示例

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
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
### getItem(key: number or string)
按名称或 ID 获取表。

#### 语法
```js
tableCollectionObject.getItem(key);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|Key|number or string|要检索的表的名称或 ID。|

#### 返回
[Table](table.md)

#### 示例

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var table = ctx.workbook.tables.getItem(tableName);
	return ctx.sync().then(function() {
			console.log(table.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### 示例

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
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

### getItemAt(index: number)
根据其在集合中的位置获取表。

#### 语法
```js
tableCollectionObject.getItemAt(index);
```

#### 参数
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### 返回
[Table](table.md)

#### 示例

```js
Excel.run(function (ctx) { 
	var table = ctx.workbook.tables.getItemAt(0);
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
### 属性访问示例

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('items');
	return ctx.sync().then(function() {
		console.log("tables Count: " + tables.count);
		for (var i = 0; i < tables.items.length; i++)
		{
			console.log(tables.items[i].name);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

获取表的数目

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	tables.load('count');
	return ctx.sync().then(function() {
		console.log(tables.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
