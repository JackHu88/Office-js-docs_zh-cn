
# 在 Outlook 的撰写窗体中获取和设置项目数据
了解如何在撰写方案中获取或设置 Outlook 外接程序中项目的不同属性，包括收件人、主题、正文和约会地点和时间。




## 获取和设置撰写加载项的项目属性


在撰写窗体中，您可以如同在阅读窗体中一样，获取在同一类型的项目上公开的大部分属性（如参与者、收件人、主题和正文），还可以获取仅与撰写窗体（而非阅读窗体）相关的一些其他属性（正文、密件抄送）。 

对于大多数属性，由于 Outlook 外接程序和用户可能会同时修改用户界面中的同一个属性，获取和设置属性的方法将为异步。表 1 列出了项目级别属性以及用于在撰写窗体中获取和设置属性的相应异步方法。[item.itemType](../../reference/outlook/Office.context.mailbox.item.md) 和 [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md) 属性是例外，因为用户无法修改。您可以使用与在阅读窗体中相同的编程方式，在撰写窗体中直接从父对象获取这些属性。

您无法访问适用于 Office 的 JavaScript API 中的项目属性，但可以使用 Exchange Web 服务 (EWS) 访问项目级别属性。如果具有  **ReadWriteMailbox** 权限，您可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法访问 EWS 操作，即 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 和 [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)，以获取并设置用户邮箱中一个或多个项目的更多属性。 **makeEwsRequestAsync** 在撰写窗体和阅读窗体中均可用。有关 **ReadWriteMailbox** 权限以及如何通过 Office 外接程序平台访问 EWS 的详细信息，请参阅 [指定 Outlook 外接程序对用户邮箱的访问权限](../outlook/understanding-outlook-add-in-permissions.md)和 [从 Outlook 外接程序调用 web 服务](../outlook/web-services.md)。


**表 1. 在撰写窗体中获取或设置项目属性的异步方法**


|**属性**|**属性类型**|**获取的异步方法**|**设置的异步方法**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[收件人](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[正文](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../../reference/outlook/Office.context.mailbox.item.md)|[时间](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[位置](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[开始](../../reference/outlook/Office.context.mailbox.item.md)|时间|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[主题](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[至](../../reference/outlook/Office.context.mailbox.item.md)|收件人|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## 其他资源



- [创建适用于撰写窗体的 Outlook 外接程序](../outlook/compose-scenario.md)
    
- [了解 Outlook 外接程序权限](../outlook/understanding-outlook-add-in-permissions.md)
    
- [从 Outlook 外接程序调用 web 服务](../outlook/web-services.md)
    
- [在阅读或撰写窗体中获取并设置 Outlook 项目数据](../outlook/item-data.md)
    


