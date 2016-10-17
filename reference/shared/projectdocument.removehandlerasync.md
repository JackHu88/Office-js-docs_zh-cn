

# <a name="projectdocument.removehandlerasync-method"></a>ProjectDocument.removeHandlerAsync 方法
异步删除 [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) 对象中任务选择更改事件的事件处理程序。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**添加内容的版本**|1.0|

```js
Office.context.document.removeHandlerAsync(eventType[, options][, callback]);
```


## <a name="parameters"></a>参数
|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
|_eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|要移除的事件类型，类型为 [EventType](../../reference/shared/eventtype-enumeration.md) 常数或其对应的文本值。必需。<br/><br/>下表展示了 [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md) 对象的有效 eventType 自变量。<br/><br/><table><tr><th>枚举</th><th>文本值</th></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179836.aspx">Office.EventType.ResourceSelectionChanged</a></td><td>resourceSelectionChanged</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179816.aspx">Office.EventType.TaskSelectionChanged</a></td><td>taskSelectionChanged</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179839.aspx">Office.EventType.ViewSelectionChanged</a></td><td>viewSelectionChanged</td></tr></table>||
|_options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
|_asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
|_callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||


## <a name="callback-value"></a>回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **removeHandlerAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


|**名称**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选 _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[error](../../reference/shared/asyncresult.error.md)|关于错误的信息（如果 **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的 **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|**removeHandlerAsync** 始终返回 **undefined**。|

## <a name="example"></a>示例

以下代码示例使用 [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) 为 [ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md) 事件添加事件处理程序，然后添加 **removeHandlerAsync** 以移除处理程序。

在资源视图中选择资源时，处理程序将显示资源 GUID。移除处理程序后，则不显示 GUID。

示例假定您的应用程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义以下页面控件。




```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
            $('#remove-handler').click(removeEventHandler);
        });
    };

    // Remove the event handler.
    function removeEventHandler() {
        Office.context.document.removeHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            {handler:getResourceGuid,
            asyncContext:'The handler is removed.'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#remove-handler').attr('disabled', 'disabled');
                    $('#message').html(result.asyncContext);
                }
            }
        );
    }

    // Get the GUID of the currently selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html('Resource GUID: ' + result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**|Selection|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录

|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## <a name="see-also"></a>另请参阅



#### <a name="other-resources"></a>其他资源


[addHandlerAsync 方法](../../reference/shared/projectdocument.addhandlerasync.md)
[EventType 枚举](../../reference/shared/eventtype-enumeration.md)
[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)

