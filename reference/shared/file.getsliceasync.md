
# <a name="file.getsliceasync-method"></a>File.getSliceAsync 方法
返回指定的切片。

|||
|:-----|:-----|
|**主机：**|PowerPoint 和 Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|文件|
|**添加内容的版本**|1.0|

```js
File.getSliceAsync(sliceIndex, callback);
```


## <a name="parameters"></a>参数


_sliceIndex_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**数字**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;指定要检索的切片从零开始的索引。必需。<br/><br/>
    
_callback_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**对象**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;返回回调时调用的函数，其唯一的参数的类型为 [AsyncResult](../../reference/shared/asyncresult.md)。可选。
    

## <a name="callback-value"></a>回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **getSliceAsync** 方法的回调函数中，你可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问 [Slice](../../reference/shared/slice.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|文件|
|**最低权限级别**|[ReadDocument（需要使用 ReadAllDocument 来获得 Office OpenXML）](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 PowerPoint 和 Word 的支持。|
|1.0|引入|
