﻿# TableCellCollection 对象（适用于 OneNote 的 JavaScript API）

_适用于：OneNote Online_  


包含 TableCell 对象的集合。

## 属性

| 属性     | 类型   |说明|反馈|
|:---------------|:--------|:----------|:-------|
|count|int|返回集合中的 tablecells 数。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-count)|
|项目|[TableCell[]](tablecell.md)|tableCell 对象的集合。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-items)|

_查看属性访问 [示例](#示例)。_

## Relationships
无


## 方法

| 方法           | 返回类型    |说明| 反馈|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableCell](tablecell.md)|按其在集合中的 ID 或索引获取 table cell 对象。 只读。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableCell](tablecell.md)|根据其在集合中的位置获取 tablecell。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|[转到](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableCellCollection-load)|

## 方法详细信息


### getItem(index: number or string)
按其在集合中的 ID 或索引获取 table cell 对象。 只读。

#### 语法
```js
tableCellCollectionObject.getItem(index);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|数字或字符串|用于标识 table cell 对象的索引位置的数字。|

#### 返回
[TableCell](tablecell.md)

### getItemAt(index: number)
根据其在集合中的位置获取 tablecell。

#### 语法
```js
tableCellCollectionObject.getItemAt(index);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|index|number|要检索的对象的索引值。从零开始编制索引。|

#### 返回
[TableCell](tablecell.md)

### load(param: object)
使用参数指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|param|object|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void