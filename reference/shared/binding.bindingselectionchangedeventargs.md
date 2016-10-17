
# <a name="bindingselectionchangedeventargs-object"></a>BindingSelectionChangedEventArgs 对象
提供有关引发 [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) 事件的绑定的信息。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**TableBindings 中的最后更改 **|1.1|

```
Office.EventType.BindingSelectionChanged
```


## <a name="members"></a>成员


**属性**


|**名称**|**Description**|
|:-----|:-----|
|[binding](../../reference/shared/binding.bindingselectionchangedevent.binding.md)|获取表示引发 [SelectionChanged](../../reference/shared/binding.md) 事件的绑定的 **Binding** 对象。|
|[columnCount](../../reference/shared/binding.bindingselectionchangedevent.columncount.md)|获取选择的列数。|
|[rowCount](../../reference/shared/binding.bindingselectionchangedevent.rowcount.md)|获取选择的行数。|
|[startRow](../../reference/shared/binding.bindingselectionchangedevent.startrow.md)|获取选择的第一行的索引（基于零）。|
|[startColumn](../../reference/shared/binding.bindingselectionchangedevent.startcolumn.md)|获取所选内容第一列的索引（从零开始）。|
|[type](../../reference/shared/binding.bindingselectionchangedevent.type.md)|获取标识被引发事件的类型的 [EventType](../../reference/shared/eventtype-enumeration.md) 枚举值。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|添加了对 Access 相关应用程序中表绑定的支持。|
|1.0|引入|
