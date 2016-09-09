
# 枚举

您可以通过使用枚举值的完全限定枚举名称 ( `Office.CoercionType.Text`) 或其相应的文本值 ( `"text"`) 来指定该枚举值。例如，以下方法调用使用了枚举名称：


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


以下是使用枚举文本值的相同调用：




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"},
   function (result) {
      if (result.status === "success")
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });
```


## 引用



|**Name**|**定义**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|指定文档活动视图的状态，例如，用户是否可以编辑文档。|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|指定异步调用的结果。|
|[AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|指定电子邮件或会议请求的附件类型。Outlook 2013 不支持此枚举。|
|[BindingType](bindingtype-enumeration.md)|指定应返回的绑定对象的类型。|
|[BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|指定约会或邮件正文的文本类型。|
|[CoercionType](coerciontype-enumeration.md)|指定如何强制由调用方法返回或设置的数据。|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|指定节点类型。|
|[DocumentMode](documentmode-enumeration.md)|指定关联应用程序中的文档为只读，还是读写。 |
|[EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|指定实体的类型。|
|[EventType](eventtype-enumeration.md)|指定引发的事件的类型。|
|[FileType](filetype-enumeration.md)|指定返回文档的格式。|
|[GoToType](gototype-enumeration.md)|指定要导航到的位置或对象类型。|
|[FilterType](filtertype-enumeration.md)|指定检索数据时是否应用从宿主应用程序筛选。|
|[InitializationReason](initializationreason-enumeration.md)|指定是刚刚插入外接程序，还是文档中已包含。|
|[ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|指定项的类型。|
|[notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|为约会或邮件指定通知邮件。|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|指定可供 [getProjectFieldAsync](projectdocument.getprojectfieldasync.md) 方法用作参数的项目字段。|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|指定可供 [getResourceFieldAsync](projectdocument.gettaskfieldasync.md) 方法用作参数的资源字段。|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|指定可供 [getTaskFieldAsync](projectdocument.gettaskfieldasync.md) 方法用作参数的任务字段。|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|指定 [getSelectedViewAsync](projectdocument.getselectedviewasync.md) 方法可以识别的视图的类型。|
|[RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|指定约会收件人的类型。|
|[ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|指定对会议邀请的回复。|
|[SelectionMode](selectionmode-enumeration.md)|指定是否选择（突出显示）要导航到的位置（使用 [Document.goToByIdAsync](document.gotobyidasync.md) 方法时）。|
|[SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|指定由调用方法返回的数据源。|
|[Table](table-enumeration.md)|指定_表格式方法_的 [cellFormat](../../docs/excel/format-tables-in-add-ins-for-excel.md) 参数中 `cells:` 属性的枚举值。|
|[ValueFormat](valueformat-enumeration.md)|指定由调用方法返回的值（如数字和日期）返回时应用了其格式设置。|

## 支持详细信息


对每个枚举的支持在跨 Office 主机应用程序之间各不相同。请参阅每个枚举主题的"支持的详细信息"部分以了解主机支持信息。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|
