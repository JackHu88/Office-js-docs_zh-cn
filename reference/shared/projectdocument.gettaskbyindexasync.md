

# <a name="projectdocument.gettaskbyindexasync-method"></a>ProjectDocument.getTaskByIndexAsync 方法
异步获取任务集合中具有指定索引的任务的 GUID。

**重要说明：**此 API 仅可在 Windows 桌面上的 Project 2016 中运行。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**添加内容的版本**|1.1|

```js
Office.context.document.getTaskByIndexAsync(taskIndex[, options][, callback]);
```


## <a name="parameters"></a>参数

_taskIndex_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**数字**

&nbsp;&nbsp;&nbsp;&nbsp;项目的任务集合中的任务索引。必需。

    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;以下是[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)：


&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;类型：**数组、布尔值、null、数字、对象、字符串**或**未定义**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[AsyncResult](../../reference/shared/asyncresult.md) 对象中未经改动的返回的任何类型的用户定义项。可选。</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;例如，可以使用 `{asyncContext: 'Some text'}` 或 `{asyncContext: <object>}` 格式传递 _asyncContext_ 参数。

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**函数**

&nbsp;&nbsp;&nbsp;&nbsp;返回方法调用时调用的函数，其唯一的参数的类型为 [AsyncResult](../../reference/shared/asyncresult.md)。可选。


## <a name="callback-value"></a>回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getTaskByIndexAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


|**名称**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选 _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[error](../../reference/shared/asyncresult.error.md)|关于错误的信息（如果 **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的 **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|**string** 形式的任务的 GUID。|

## <a name="remarks"></a>备注

要获取项目的任务集合的最大索引，请使用 [getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) 方法。0 索引请求表示项目摘要任务。


## <a name="example"></a>示例

下面的代码示例调用 [getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md) 以获取项目的任务集合中的最大索引，然后调用 **getTaskByIndexAsync**获取每个任务的 GUID。

示例假定您的外接程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义。




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var taskGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the maximum task index, and then get the task GUIDs.
    function getTaskInfo() {
        getMaxTaskIndex().then(
            function (data) {
                getTaskGuids(data);
            }
        );
    }

    // Get the maximum index of the tasks for the current project.
    function getMaxTaskIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxTaskIndexAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get each task GUID, and then display the GUIDs in the add-in.
    function getTaskGuids(maxTaskIndex) {
        var defer = $.Deferred();
        for (var i = 0; i <= maxTaskIndex; i++) {
            getTaskGuid(i);
        }
        return defer.promise();
        function getTaskGuid(index) {
            Office.context.document.getTaskByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        taskGuids.push(result.value);
                        if (index == maxTaskIndex) {
                            defer.resolve();
                            $('#message').html(taskGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
    }
    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**在要求集中可用**||
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录

|**版本**|**更改内容**|
|:-----|:-----|
|1.1|引入|

## <a name="see-also"></a>另请参阅



#### <a name="other-resources"></a>其他资源


[getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)
[AsyncResult 对象](../../reference/shared/asyncresult.md)
[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
