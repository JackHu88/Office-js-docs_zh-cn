
# <a name="binding-object"></a>Binding 对象
表示对一部分文档的绑定的抽象类。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集中可用](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBinding, TableBinding, TextBinding|
|**包含 TableBindings 最后一次更改的版本**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

## <a name="members"></a>成员


**对象**


|**名称**|**说明**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|表示两个维度的行和列的绑定。|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|表示两个维度的行和列的绑定，标题可选。|
|[TextBinding](../../reference/shared/binding.textbinding.md)|表示文档中的绑定文本选择。|

**属性**


|**名称**|**说明**|
|:-----|:-----|
|[document](../../reference/shared/binding.document.md)|获取与绑定关联的  **Document** 对象。|
|[id](../../reference/shared/binding.id.md)|获取对象的标识符。|
|[type](../../reference/shared/binding.type.md)|获取绑定的类型。|

**方法**


|**名称**|**说明**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|将处理程序添加到指定事件类型的绑定。|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|返回绑定中包含的数据。|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|从绑定移除指定事件类型的指定处理程序。|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|将数据写入指定的绑定对象表示的文档的绑定部分。|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|设置或更新绑定表中指定项目和数据的格式。|

**事件**


|**名称**|**说明**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|绑定内的数据更改时发生。|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|绑定内的选择更改时发生。|

## <a name="remarks"></a>备注

**Binding** 对象揭示所有绑定拥有的功能，不管其类型。

不能直接调用 **Binding** 对象。它是表示每种类型的绑定的对象的抽象父类：[MatrixBinding](../../reference/shared/binding.matrixbinding.md)、[TableBinding](../../reference/shared/binding.tablebinding.md) 或 [TextBinding](../../reference/shared/binding.textbinding.md)。这三个对象都从 **Binding** 对象继承 **getDataAsync** 和 **setDataAsync** 方法，该对象可让您与绑定中的数据交互。它们还继承 **id** 和 **type** 属性，以便查询这些属性值。此外，**MatrixBinding** 和 **TableBinding** 对象揭示特定于矩阵和表的功能的其他方法，如对行和列计数。


## <a name="support-details"></a>支持详细信息


各 Office 主机应用程序对  **Binding** 对象的每个 API 成员的支持不同。请参阅每个成员主题的"支持详细信息"部分了解主机支持信息。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||
|:-----|:-----|
|**在要求集中可用**|MatrixBinding, TableBinding, TextBinding|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|
