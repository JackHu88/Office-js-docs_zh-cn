

# ProjectDocument.getWSSUrlAsync 方法
异步获取同步的 SharePoint 任务列表的 URL。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**在其中添加**|1.0|

```js
Office.context.document.getWSSUrlAsync([options,] [callback]);
```


## 参数



|**名称**|**类型**|**说明**|**支持说明**|
|:-----|:-----|:-----|:-----|
| _选项_|**object**|指定以下任一[可选参数](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)。||
| _asyncContext_|**array**、 **boolean**、 **null**、 **number**、 **object** 、 **string** 或 **undefined**|在  **AsyncResult** 对象中未经改动的返回的任何类型的用户定义项。||
| _callback_|**object**|返回回调时调用的函数，其唯一的参数的类型为  **AsyncResult** 。||

## 回调值

在 _callback_ 函数执行后，它会收到你可以从回调函数的参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

对于 **getWSSUrlAsync** 方法，返回的 [AsyncResult](../../reference/shared/asyncresult.md) 对象包含下列属性。


|**姓名**|**说明**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|在可选  _asyncContext_ 参数中传递的数据（如果使用了参数）。|
|[错误](../../reference/shared/asyncresult.error.md)|关于错误的信息（ 如果  **status** 属性为 **failed**）|
|[status](../../reference/shared/asyncresult.status.md)|异步调用的  **succeeded** 或 **failed** 状态。|
|[value](../../reference/shared/asyncresult.value.md)|包含以下属性：<br/><br/><ul><li><b>listName</b> 属性是同步的 SharePoint 任务列表的名称。</li><li><b>serverUrl</b> 属性是同步的 SharePoint 任务列表的 URL。</li></ul>|

## 注解

如果活动项目未与 SharePoint 任务列表同步，则  **serverUrl** 和 **listName** 值为空。


## 示例

以下代码示例调用  **getWSSUrlAsync** 以获取同步 SharePoint 任务列表的名称和 URL。

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
            getSharePointTaskListUrl();
        });
    };

    // Get the URL of the the synchronized SharePoint task list.
    function getSharePointTaskListUrl() {
        Office.context.document.getWSSUrlAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var output = String.format(
                        'List name: {0}<br />List URL: {1}',
                        result.value.listName, result.value.serverUrl);
                    $('#message').html(output);
                }
                else {
                    onError(result.error);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**||
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## 另请参阅



#### 其他资源


[AsyncResult 对象](../../reference/shared/asyncresult.md)
(#asyncresult-对象)[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
