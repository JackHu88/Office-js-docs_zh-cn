

# <a name="projectdocument.getselectedviewasync-method"></a>ProjectDocument.getSelectedViewAsync 方法
异步获取文档中活动视图的类型和名称。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**添加内容的版本**|1.0|

```js
Office.context.document.getSelectedViewAsync([options,] [callback]);
```


## <a name="parameters"></a>参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|指定以下任一 [可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**array**、**boolean**、**null**、**number**、**object**、**string** 或 **undefined**|在 **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为 **AsyncResult** 。||

## <a name="callback-value"></a>回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getSelectedViewAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


****


|**名称**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选 _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[error](../../reference/shared/asyncresult.error.md)|关于错误的信息（如果 **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的 **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|包含以下属性：<br/><br/><div>* **viewName** 属性是视图的名称，作为 [ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md) 常量。<br/>* **viewType** 属性是视图的类型，作为 [ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md) 常量的整数值。</div>|

## <a name="example"></a>示例

以下代码示例调用添加 [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) 事件处理程序，它调用 **getSelectedViewAsync** 以获取文档中活动视图的名称和类型。

示例假定您的应用程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义以下页面控件。




```HTML
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
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the active view's name and type.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
                    $('#message').html(output);
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
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## <a name="see-also"></a>另请参阅



#### <a name="other-resources"></a>其他资源


[ProjectViewTypes 枚举](../../reference/shared/projectviewtypes-enumeration.md)
[AsyncResult 对象](../../reference/shared/asyncresult.md)
[ViewSelectionChanged 事件](../../reference/shared/projectdocument.viewselectionchanged.event.md)
[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
