
# ProjectDocument.TaskSelectionChanged 事件
活动项目中的任务选择更改时发生。

|||
|:-----|:-----|
|**主机：**|Project|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Selection|
|**在其中添加**|1.0|

```js
Office.EventType.TaskSelectionChanged
```


## 注解

 **TaskSelectionChanged** 是一个 [EventType](../../reference/shared/eventtype-enumeration.md) 枚举常量，该常量可用于在 [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) 和 [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) 方法中添加或移除事件的处理程序。


## 示例

以下代码示例添加  **TaskSelectionChanged** 事件的处理程序。当文档中的任务选择变更时，它将获取所选任务的 GUID。

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
                Office.EventType.TaskSelectionChanged,
                getTaskGuid);
            getTaskGuid();
        });
    };

    // Get the GUID of the selected task and display it in the add-in.
    function getTaskGuid() {
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html(result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```

有关展示了如何在 Project 外接程序中使用 **TaskSelectionChanged** 事件处理程序的示例，请参阅[使用文本编辑器为 Project 2013 创建首个任务窗格外接程序](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此事件。空的单元格表示相应的 Office 主机应用程序不支持此事件。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||||
|:-----|:-----|:-----|
||Office for Windows Desktop|Office Online（在浏览器中）|
|**Project**|Y||

|||
|:-----|:-----|
|**在要求集中可用**|Selection|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.0|<ul><li>引入</li></ul>|

## 另请参阅



#### 其他资源


[使用文本编辑器创建 Project 2013 的第一个任务窗格外接程序](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
(#使用文本编辑器创建-project-2013-的第一个任务窗格外接程序)[EventType 枚举](../../reference/shared/eventtype-enumeration.md)
(#eventtype-枚举)[ProjectDocument.addHandlerAsync 方法](../../reference/shared/projectdocument.addhandlerasync.md)
(#projectdocument.addhandlerasync-方法)[ProjectDocument.removeHandlerAsync 方法](../../reference/shared/projectdocument.removehandlerasync.md)
(#projectdocument.removehandlerasync-方法)[ProjectDocument 对象](../../reference/shared/projectdocument.projectdocument.md)
