# ChartAxes 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Office 2016_

表示图表坐标轴。

## 属性

无

## Relationships
| 关系 | 类型|说明|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|表示图表中的类别轴。只读。|
|seriesAxis|[ChartAxis](chartaxis.md)|表示三维图表的系列轴。只读。|
|valueAxis|[ChartAxis](chartaxis.md)|表示坐标轴中的数值轴。只读。|

## 方法

| 方法   | 返回类型|说明|
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
| 参数   | 类型|说明|
|:---------------|:--------|:----------|
|param|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

