# SortField 对象（适用于 Excel 的 JavaScript API）

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

表示排序操作中的条件。

## 属性

| 属性	  | 类型	|说明
|:---------------|:--------|:----------||ascending|bool|表示是否以升序方式排序。||color|字符串|表示对字体或单元格颜色进行排序时，条件的目标颜色。||dataOption|字符串|表示此字段的其他排序选项。可能的值是：Normal、TextAsNumber。||key|int|表示条件所在的列（或行，具体取决于排序方向）。表示距离第一列（或行）的偏移量。||sortOn|字符串|表示此条件的排序类型。可能的值是：Value、CellColor、FontColor、Icon。|_请参阅属性访问[示例。](#property-access-examples)_

## 关系
|关系|类型	|说明||:---------------|:--------|:----------||icon|[图标](icon.md)|表示对单元格图标进行排序时，条件的目标图标。|

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
| 参数	  | 类型	|说明||:---------------|:--------|:----------||参数|对象|可选。接受参数和关系名称作为分隔字符串或数组。或者提供 [loadOption](loadoption.md) 对象。|

#### 返回
void

