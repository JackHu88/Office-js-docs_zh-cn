
# <a name="settings.saveasync-method"></a>Settings.saveAsync 方法
将设置属性包的内存副本保留到文档中。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint 和 Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Settings|
|**包含最后一次更改的版本**|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## <a name="parameters"></a>参数



_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**对象**

&nbsp;&nbsp;&nbsp;&nbsp;返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult**。可选。

    



## <a name="callback-value"></a>回调值

在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **saveAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|始终返回 **undefined**，这是因为没有要检索的对象或数据。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果用户定义的一个 **object** 或值作为 _asyncContext_ 参数传递，则对其进行访问。|

## <a name="remarks"></a>备注

外接程序之前保存的所有设置都会在外接程序初始化时加载。因此，在会话的生存期内，你可以只使用 [set](../../reference/shared/settings.set.md) 和 [get](../../reference/shared/settings.get.md) 方法来处理设置属性包的内存中副本。如果你希望保留设置以便在下次使用外接程序时使用，请使用 **saveAsync** 方法。


 >**注意**：**saveAsync** 方法会将内存中设置属性包保留到文档文件中；但文档文件本身的更改仅在用户（或 **AutoRecover** 设置）将文档保存到文件系统时才会得以保存。

[refreshAsync](../../reference/shared/settings.refreshasync.md) 方法仅在合著方案（只有 Word 中才支持合著方案）中非常有用，在合著方案中相同外接程序的其他实例可能更改设置，并且应使这些更改可供所有实例使用。


## <a name="example"></a>示例




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|Settings|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关内容外接程序中自定义设置的支持。|
|1.0|引入|
