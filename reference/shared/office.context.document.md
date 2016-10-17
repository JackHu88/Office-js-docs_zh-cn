
# <a name="context.document-property"></a>Context.document 属性
获取表示与外接程序交互的文档的对象。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```js
var _document = Office.context.document;
```


## <a name="return-value"></a>返回值

[Document](../../reference/shared/document.md) 对象。


## <a name="remarks"></a>备注

您的外接程序可使用  **document** 属性访问 API 以与文档、工作簿、演示文稿、项目和（Access Web 应用程序中的）数据库中的内容交互。


## <a name="example"></a>示例




```js
// Extension initialization code.
var _document;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Initialize instance variables to access API objects.
    _document = Office.context.document;
    });
}

```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对  **Office.context.document** 访问 Access 相关内容外接程序中数据库的支持。|
|1.0|引入|
