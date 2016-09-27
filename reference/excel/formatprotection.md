﻿# FormatProtection 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示对范围对象的格式保护。

## 属性

| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|formulaHidden|bool|表示 Excel 是否隐藏区域中的单元格公式。指示整个区域不具有统一公式隐藏设置的空值。|
|locked|bool|指示 Excel 是否锁定对象中的单元格。指示整个区域不具有统一锁定设置的空值。|

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