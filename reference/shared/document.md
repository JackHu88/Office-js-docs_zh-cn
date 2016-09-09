
# Document 对象
表示与外接程序交互的文档的抽象类。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint、Project、Word|
|**在其中添加**|1.0|
|**包含最后一次更改的版本**|1.1|

```
Office.context.document
```


## 成员


**属性**


|**名称**|**说明**|**支持说明**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|获取提供对文档中定义的绑定的访问的对象。|添加了对 Access 相关内容应用程序的支持。|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|获取文档中表示自定义 XML 部件的对象。||
|[mode](../../reference/shared/document.mode.md)|获取文档所处的模式。|添加了对 Access 相关内容应用程序的支持。|
|[设置](../../reference/shared/document.settings.md)|获取用于表示当前文档的内容或任务窗格应用程序的已保存自定义设置的对象。|添加了对 Access 相关内容应用程序的支持。|
|[url](../../reference/shared/document.url.md)|获取主机应用程序当前打开的文档的 URL。|添加了对 Access 相关内容应用程序的支持。|

**方法**


|**名称**|**说明**|**支持说明**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|添加  **Document** 对象事件的事件处理程序。||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|返回演示文稿的当前视图。|在 1.1 中，现已开始支持 [PowerPoint 的外接程序](../../docs/powerpoint/powerpoint-add-ins.md)。|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|以高达 4194304 字节 (4 MB) 的切片形式返回整个文档文件。|在 1.1 中，现已开始支持在 PowerPoint 和 Word 的外接程序中将文件作为 PDF 获取。|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|获取当前文档的文件属性。在此版本中，只能获取文档的 URL。|在 1.1 中，现已开始支持在 Excel、Word 和 PowerPoint 的外接程序中获取文档的 URL。|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|读取包含在文档的当前选择中的数据。|在 1.1 中，现已开始支持在 PowerPoint 的外接程序中获取选定范围幻灯片的 ID、标题和索引。|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|转到文档中指定的对象或位置。|在 1.1 中，现已开始支持在 Excel 和 PowerPoint 的外接程序中进行文档内导航。|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|移除  **Document** 对象事件的事件处理程序。||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|将数据写入文档中的当前选择。|在 1.1 中，现已开始支持[在 Excel 的外接程序中写入数据时设置选定表的格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。|

**事件**


|**名称**|**说明**|**支持说明**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|用户更改文档的当前视图时出现。|在 1.1 中，现已开始支持 PowerPoint 的外接程序。||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|文档中的选择更改时发生。|||

## 备注

请勿在脚本中直接实例化  **Document** 对象。若要调用 **Document** 对象的成员以便与当前文档或工作表交互，请使用脚本中的 `Office.context.document`。


## 示例

如下示例使用  **Document** 对象的 **getSelectedDataAsync** 方法以作为文本检索用户的当前选择，然后将其显示在应用程序的页面中。


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 要求


对 **Document** 对象的各个 API 成员的支持因 Office 主机应用程序而异。请参阅各个成员主题的“支持详细信息”部分，了解主机支持信息。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||
|:-----|:-----|
|**在其中添加**|1.0|
|**包含最后一次更改的版本**|1.1|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|
