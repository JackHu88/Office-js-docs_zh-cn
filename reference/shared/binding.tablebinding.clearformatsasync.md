
# <a name="tablebinding.clearformatsasync-method"></a>TableBinding.clearFormatsAsync 方法
清除绑定表中的格式。

|||
|:-----|:-----|
|**主机：**|Excel|
|在**要求集[中可用](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|不在集合中|
|**在其中添加**|1.1|

```js
bindingObj.clearFormatsAsync([,options] , callback);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**Description**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

当执行您传递给 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **goToByIdAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined** 因为在清除格式时没有要检索的数据或对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果将用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>注解

请参阅[操作说明：设置 Excel 相关外接程序中表的格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)，以便了解详细信息。


## <a name="example"></a>示例

以下示例显示如何清除 ID 为“myBinding”的绑定表的格式。


```js
Office.select(bindings#myBinding).clearFormatsAsync();
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|不在集合中|
|**最低权限级别**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 的支持。|
|1.0|引入|
