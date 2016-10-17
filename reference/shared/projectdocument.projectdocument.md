

# <a name="projectdocument-object"></a>ProjectDocument 对象
表示与 Office 外接程序交互的项目文档（活动项目）的抽象类。

|||
|:-----|:-----|
|**主机：**|Project|
|**添加内容的版本**|1.0|

```js
Office.context.document
```


## <a name="members"></a>成员


**方法**


|**名称**|**说明**|
|:-----|:-----|
|[addHandlerAsync 方法](../../reference/shared/projectdocument.addhandlerasync.md)|在 **ProjectDocument** 对象中为事件异步添加事件处理程序。|
|[getMaxResourceIndexAsync 方法](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|异步获取当前项目中的资源集合的最大索引。|
|[getMaxTaskIndexAsync 方法](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|异步获取当前项目中的任务集合的最大索引。|
|[getProjectFieldAsync 方法](../../reference/shared/projectdocument.getprojectfieldasync.md)|异步获取活动项目中指定字段的值。|
|[getResourceByIndexAsync 方法](../../reference/shared/projectdocument.getresourcebyindexasync.md)|异步获取资源集合中具有指定索引的资源的 GUID。|
|[getResourceFieldAsync 方法](../../reference/shared/projectdocument.getresourcefieldasync.md)|异步获取指定资源的指定字段的值。|
|[getSelectedDataAsync 方法](../../reference/shared/projectdocument.getselecteddataasync.md)|异步获取甘特图中一个或多个单元格的当前选择中包含的数据。|
|[getSelectedResourceAsync 方法](../../reference/shared/projectdocument.getselectedresourceasync.md)|异步获取选定资源的 GUID。|
|[getSelectedTaskAsync 方法](../../reference/shared/projectdocument.getselectedtaskasync.md)|异步获取所选任务的 GUID。|
|[getSelectedViewAsync 方法](../../reference/shared/projectdocument.getselectedviewasync.md)|异步获取活动视图的视图类型和名称。|
|[getTaskAsync 方法](../../reference/shared/projectdocument.gettaskasync.md)|异步获取同步的 SharePoint 任务列表中的任务名称、分配给任务的资源和任务的 ID。|
|[getTaskByIndexAsync 方法](../../reference/shared/projectdocument.gettaskbyindexasync.md)|异步获取任务集合中具有指定索引的任务的 GUID。|
|[getTaskFieldAsync 方法](../../reference/shared/projectdocument.gettaskfieldasync.md)|异步获取指定任务的指定字段的值。|
|[getWSSUrlAsync 方法](../../reference/shared/projectdocument.getwssurlasync.md)|异步获取同步的 SharePoint 任务列表的 URL。|
|[removeHandlerAsync 方法](../../reference/shared/projectdocument.removehandlerasync.md)|在 **ProjectDocument** 对象中为事件异步移除事件处理程序。|
|[setResourceFieldAsync 方法](../../reference/shared/projectdocument.setresourcefieldasync.md)|异步设置指定资源的指定字段的值。|
|[setTaskFieldAsync 方法](../../reference/shared/projectdocument.settaskfieldasync.md)|异步设置指定任务的指定字段的值。|

**事件**


|**名称**|**说明**|
|:-----|:-----|
|[ResourceSelectionChanged 事件](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|当活动项目中的资源选择发生更改时发生。|
|[TaskSelectionChanged 事件](../../reference/shared/projectdocument.taskselectionchanged.event.md)|活动项目中的任务选择更改时发生。|
|[ViewSelectionChanged 事件](../../reference/shared/projectdocument.viewselectionchanged.event.md)|当活动项目中的活动视图发生更改时发生。|

## <a name="remarks"></a>备注

请勿直接调用或实例化脚本中的  **ProjectDocument** 对象。


## <a name="example"></a>示例

以下示例实例化外接程序，然后获取 Project 文档上下文中可用的 [Document](../../reference/shared/document.md) 对象的属性。Project 文档是已打开且活动的项目，要访问 **ProjectDocument** 对象的成员，请使用 **Office.context.document** 对象，如 **ProjectDocument** 方法和事件的代码示例中所示。

示例假定您的外接程序具有对 jQuery 库的引用，且以下页面控件在页面正文的内容中定义以下页面控件。




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## <a name="see-also"></a>另请参阅



#### <a name="other-resources"></a>其他资源


[Project 的任务窗格外接程序](../../docs/project/project-add-ins.md)
[Document 对象](../../reference/shared/document.md)

