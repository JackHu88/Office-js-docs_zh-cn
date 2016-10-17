
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>在阅读或撰写窗体中获取并设置 Outlook 项目数据

从 Office 外接程序清单架构的版本 1.1 开始，Outlook 可以在用户查看或撰写项目时激活外接程序。根据外接程序是在阅读窗体中激活还是在撰写窗体中激活，项目为应用程序提供的属性也有所不同。例如，仅针对已发送项目（随后在阅读窗体中查看项目）定义 [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) 和 [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) 属性，但在（撰写窗体中）创建项目时不定义这两个属性。另一个示例是： [bcc](../../reference/outlook/Office.context.mailbox.item.md) 属性仅在（撰写窗体中）创作邮件时具有意义，并且用户在阅读窗体中无法访问此属性。

表 1 显示适用于 Office 的 JavaScript API 中可用于邮件外接程序的每个阅读和撰写模式的项目级属性。通常，阅读窗体中可用的属性是只读的，撰写窗体中可用的属性是可读取/写入的， [itemId](../../reference/outlook/Office.context.mailbox.item.md)* 和 [conversationId](../../reference/outlook/Office.context.mailbox.item.md)* 属性除外，尽管这两个属性也是只读的。对于撰写窗体中的其余项目级属性，由于外接程序和用户可以同时读取或写入同一属性，在撰写模式下获取或设置这些属性的方法都是异步的，因此这些属性在撰写窗体中和阅读窗体中返回的对象类型也有所不同。有关在撰写模式下使用异步方法获取或设置项目级属性的详细信息，请参阅 [在 Outlook 的撰写窗体中获取和设置项目数据](../outlook/get-and-set-item-data-in-a-compose-form.md)。


**表 1.撰写和阅读窗体中可用的项目属性**


|**项目类型**|**属性**|**阅读窗体中的属性类型**|**撰写窗体中的属性类型**|
|:-----|:-----|:-----|:-----|
|约会和邮件|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 对象|属性不可用|
|约会和邮件|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 对象|属性不可用|
|约会和邮件|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|字符串|属性不可用|
|约会和邮件|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|字符串|属性不可用|
|约会和邮件|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|[ItemType](../../reference/outlook/Office.MailboxEnums.md) 枚举中的字符串|属性不可用|
|约会和邮件|[attachments](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|属性不可用|
|约会和邮件|[body](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body](../../reference/outlook/Body.md)|
|约会|[end](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 对象|[Time](../../reference/outlook/Time.md)|
|约会|[location](../../reference/outlook/Office.context.mailbox.item.md)|字符串|[Location](../../reference/outlook/Location.md)|
|约会和邮件|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|字符串|属性不可用|
|约会|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[Recipients](../../reference/outlook/Recipients.md)|
|约会|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|属性不可用|
|约会|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|收件人|
|约会|[resources](../../reference/outlook/Office.context.mailbox.item.md)|字符串|属性不可用|
|约会|[start](../../reference/outlook/Office.context.mailbox.item.md)|JavaScript **Date** 对象|时间|
|约会和邮件|[subject](../../reference/outlook/Office.context.mailbox.item.md)|字符串|[Subject](../../reference/outlook/Subject.md)|
|邮件|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|属性不可用|收件人|
|邮件|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|收件人|
|邮件|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|字符串|字符串（只读）|
|邮件|[from](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|属性不可用|
|邮件|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|整数|属性不可用|
|邮件|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|属性不可用|
|邮件|[to](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|收件人|

## <a name="using-exchange-server-callback-tokens-from-a-read-add-in"></a>从阅读加载项使用 Exchange Server 回调令牌


如果您的 Outlook 外接程序将要在阅读窗体中进行激活，您可以获取 Exchange 回调令牌。可以在服务器端代码中使用此令牌，可以通过 Exchange Web 服务 (EWS) 获取对完整项目的访问权限。通过在外接程序清单中指定  **ReadItem** 权限，您可以使用 [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) 方法获取 Exchange 回调令牌，使用 [mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) 属性获取用户邮箱的 EWS 终结点的 URL，并使用 [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) 方法获取选定项目的 EWS ID。然后可以将回调令牌、EWS 终结点 URL 和 EWS 项目 ID 传递给服务器端代码，以访问 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 操作，从而获取项目的更多属性。


## <a name="accessing-ews-from-a-read-or-compose-add-in"></a>从读取或撰写外接程序访问 EWS


您还可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法直接从外接程序访问 Exchange Web 服务 (EWS) 操作 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 和 [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)。您可以使用这两个操作获取并设置指定项目的多个属性。无论外接程序已在阅读还是撰写窗体中激活，只要在外接程序清单中指定了  **ReadWriteMailbox** 权限，Outlook 外接程序就可以使用此方法。有关使用 **makeEwsRequestAsync** 访问 EWS 操作的详细信息，请参阅 [从 Outlook 外接程序调用 web 服务](../outlook/web-services.md)。


## <a name="additional-resources"></a>其他资源



- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
- [在 Outlook 的撰写窗体中获取和设置项目数据](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [从 Outlook 外接程序调用 Web 服务](../outlook/web-services.md)
    


