
# BindingDataChangedEventArgs 对象
提供有关引发 [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md) 事件的绑定的信息。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**BindingEvents 中的最后更改**|1.1|

```js
Office.EventType.BindingDataChanged
```


## 成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|[绑定](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|获取表示引发 [DataChanged](../../reference/shared/binding.md) 事件的绑定的 **Binding** 对象。|
|[类型](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|获取标识被引发事件的类型的 [EventType](../../reference/shared/eventtype-enumeration.md) 枚举值。|

## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|添加了对 Access 相关应用程序中此事件的支持。|
|1.0|引入|
