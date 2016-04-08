# Icon 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示单元格图标。

## 属性

| 属性	  | 类型	|说明
|:---------------|:--------|:----------||index|int|表示在给定集合中的图标索引。||set|string|表示该图标所属的集合。可能的值是：Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|_查看属性访问[示例。](#property-access-examples)_

## 关系
无


## 方法

| 方法		  | 返回类型	|说明||:---------------|:--------|:----------||[load(param: object)](#loadparam-object)|void|使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。|

## 方法详细信息


### load(param: object)
使用参数中指定的属性和对象值填充在 JavaScript 层中创建的代理对象。

#### 语法
```js
object.load(param);
```

#### 参数
| 参数	  | 类型	|说明||:---------------|:--------|:----------||参数|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者，提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

