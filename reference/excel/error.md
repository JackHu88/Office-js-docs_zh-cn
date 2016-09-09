# OfficeExtension.Error 对象（适用于 Excel 的 JavaScript API）

表示使用 Excel JavaScript API 时出现的错误。

_适用于：Excel 2016、Excel Online、Excel for iOS、Office 2016_

## 属性
| 属性     | 类型   |说明
|:---------------|:--------|:----------|
|code|string|获取一个指示错误类型的值。值可以是“AccessDenied”、“ActivityLimitReached”、“BadPassword”、“GeneralException”、“InsertDeleteConflict”、“InvalidArgument”、“InvalidBinding”、“InvalidOperation”、“InvalidReference”、“InvalidSelection”、“ItemAlreadyExists”、“ItemNotFound”、“NotImplemented”或“UnsupportedOperation”。 |
|debugInfo|string|获取指示出错时所发生的问题的一个值。此值仅在开发/调试过程中使用。  |
|邮件 |string| 获取与错误代码对应的本地化的人工读取字符串。|
|name |string| 获取一个始终为“OfficeExtension.Error”的值。 |
|traceMessages |string[]| 获取值数组，这些值与通过 context.trace(); 设置的检测消息对应 |

## 方法

| 方法           | 返回类型    |说明|
|:---------------|:--------|:----------|
|[toString()](#tostring)|string|以下面的格式返回错误代码和消息值：“{0}: {1}”、代码、消息。|

## 方法详细信息

### toString()
以下面的格式返回错误代码和消息值：“{0}: {1}”、代码、消息。

#### 语法
```js
error.toString()
```

#### 参数
无。

#### 返回
字符串
