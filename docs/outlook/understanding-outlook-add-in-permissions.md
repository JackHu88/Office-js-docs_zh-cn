
# <a name="understanding-outlook-add-in-permissions"></a>了解 Outlook 外接程序权限

Outlook 外接程序在其清单中指定所需的权限级别。可用级别为“**Restricted**”、“**ReadItem**”、“**ReadWriteItem**或“**ReadWriteMailbox**”。这些权限级别具有累积性：“**Restricted**”是最低的级别，并且每个更高级别包括所有较低级别的权限。“**ReadWriteMailbox**”包含所有受支持的权限。

在从 Office 商店安装邮件外接程序之前，您可以查看该邮件外接程序所需的权限。您还可以在 Exchange 管理员中心中查看已安装外接程序所需的权限。


## <a name="restricted-permission"></a>“Restricted”权限


“**Restricted**”权限是最基本级别的权限。在清单中的“**权限**”元素中指定“[Restricted](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx)”以请求此权限。如果外接程序不请求其清单中的将特定权限，在默认情况下，Outlook 会将此权限分配给邮件外接程序。


### <a name="can-do"></a>可以执行的操作


- [仅获取项目主题或正文的特定实体](../outlook/match-strings-in-an-item-as-well-known-entities.md)（电话号码、地址、URL）。
    
- 指定[项目激活规则](../outlook/manifests/activation-rules.md#itemis-rule)，此类规则需要阅读或撰写窗体中的当前项目为特定的项目类型，或与选定项目中支持的已知实体（电话号码、地址、URL）的任何较小子集匹配的 [ItemHasKnownEntity rule](../outlook/match-strings-in-an-item-as-well-known-entities.md) 规则。
    
- 访问 **不** 与用户或项目具体信息相关的任何属性和方法。（请参阅下一部分，获取与用户或项目具体信息相关的属性和方法列表。）
    

### <a name="can't-do"></a>不能执行的操作


- 在联系人、电子邮件地址、会议建议或任务建议实体上使用 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 规则。
    
- 使用 [ItemHasAttachment](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx) 或 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 规则。
    
- 访问以下列表中与用户或项目相关的成员。尝试访问此列表中的成员将返回  **null** 并导致出现一条错误消息，此消息指出 Outlook 要求邮件加载项具有提升的权限。
    
      - [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.attachments](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.bcc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.body](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.cc](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.from](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatches](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.getRegExMatchesByName](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.organizer](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.resources](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.sender](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [item.to](../../reference/outlook/Office.context.mailbox.item.md)
    
  - [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)
    
  - [mailbox.userProfile](../../reference/outlook/Office.context.mailbox.userProfile.md)
    
  - [Body](../../reference/outlook/Body.md) 及其所有子成员
    
  - [Location](../../reference/outlook/Location.md) 及其所有子成员
    
  - [Recipients](../../reference/outlook/Recipients.md) 及其所有子成员
    
  - [Subject](../../reference/outlook/Subject.md) 及其所有子成员
    
  - [Time](../../reference/outlook/Time.md) 及其所有子成员
    

## <a name="readitem-permission"></a>“ReadItem”权限


“**ReadItem**”权限是权限模型中的下一级别的权限。在清单中的“**权限**”元素中指定“**ReadItem**”以请求此权限。


### <a name="can-do"></a>可以执行的操作


- 在读取或 [撰写窗体](../outlook/item-data.md)[中读取当前项目的所有属性](../outlook/get-and-set-item-data-in-a-compose-form.md)，例如阅读窗体中的 [item.to](../../reference/outlook/Office.context.mailbox.item.md) 和撰写窗体中的 [item.to.getAsync](../../reference/outlook/Recipients.md)。
    
- [获取回调令牌以获取项目附件](../outlook/get-attachments-of-an-outlook-item.md)或整个项目。
    
- 
  [编写加载项在该项目上设置的自定义属性](http://msdn.microsoft.com/library/30217d63-7615-4f3f-8618-c91e4e60cd43%28Office.15%29.aspx)。
    
- 从该项目的主题或正文中[获取所有现有已知实体](../outlook/match-strings-in-an-item-as-well-known-entities.md)，而不仅仅是一个子集。
    
- 使用 [ItemHasKnownEntity](../outlook/manifests/activation-rules.md#itemhasknownentity-rule) 规则中所有的 [已知实体](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx)，或者 [ItemHasRegularExpressionMatch](../outlook/manifests/activation-rules.md#itemhasregularexpressionmatch-rule) 规则中的 [正则表达式](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)。以下示例遵循架构 v1.1。这说明，如果在选定邮件的主题或正文中找到一个或多个已知实体，则以下规则将激活加载项：
    

```XML
<Permissions>ReadItem</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="MeetingSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="TaskSuggestion" />
        <Rule xsi:type="ItemHasKnownEntity" 
            EntityType="EmailAddress" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
</Rule>
```


### <a name="can't-do"></a>不能执行的操作

访问 **mailbox.makeEWSRequestAsync** 或者以下撰写方法：


- [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.bcc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.bcc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.body.prependAsync](../../reference/outlook/Body.md)
    
- [item.body.setAsync](../../reference/outlook/Body.md)
    
- [item.body.setSelectedDataAsync](../../reference/outlook/Body.md)
    
- [item.cc.addAsync](../../reference/outlook/Recipients.md)
    
- [item.cc.setAsync](../../reference/outlook/Recipients.md)
    
- [item.end.setAsync](../../reference/outlook/Time.md)
    
- [item.location.setAsync](../../reference/outlook/Location.md)
    
- [item.optionalAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.optionalAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)
    
- [item.requiredAttendees.addAsync](../../reference/outlook/Recipients.md)
    
- [item.requiredAttendees.setAsync](../../reference/outlook/Recipients.md)
    
- [item.start.setAsync](../../reference/outlook/Time.md)
    
- [item.subject.setAsync](../../reference/outlook/Subject.md)
    
- [item.to.addAsync](../../reference/outlook/Recipients.md)
    
- [item.to.setAsync](../../reference/outlook/Recipients.md)
    

## <a name="readwriteitem-permission"></a>"ReadWriteItem"权限


可以在清单中的  **Permissions** 元素中指定 **ReadWriteItem** 以请求此权限。在使用撰写方法（例如， **Message.to.addAsync** 或 **Message.to.setAsync**）的撰写窗体中激活的邮件加载项必须使用至少这个等级的权限。


### <a name="can-do"></a>可以执行的操作


- [读取和写入正在 Outlook 中查阅或撰写的项目的所有项目级别属性](../outlook/item-data.md)。
    
- [添加或移除该项目的附件](../outlook/add-and-remove-attachments-to-an-item-in-a-compose-form.md)。
    
- 使用适用于邮件加载项的 Office JavaScript API 的所有其他成员（ **Mailbox.makeEWSRequestAsync** 除外）。
    

### <a name="can't-do"></a>不能执行的操作

使用 **Mailbox.makeEWSRequestAsync**。


## <a name="readwritemailbox-permission"></a>“ReadWriteMailbox”权限


“**ReadWriteMailbox**”权限是最高级别的权限。在清单中的“**权限**”元素中指定“**ReadWriteMailbox**”以请求此权限。

除了“**ReadWriteItem**权限所支持的操作，通过使用 **Mailbox.makeEWSRequestAsync**，你还可以访问支持的 Exchange Web Services (EWS) 操作，以执行以下操作：


- 读取和写入用户邮箱中任何项目的所有属性。
    
- 创建、读取和写入该邮箱中的任何文件夹或项目。
    
- 从该邮箱发送项目
    
通过 **mailbox.makeEWSRequestAsync**，可以访问以下 EWS 操作：


- 
  [CopyItem](http://msdn.microsoft.com/en-us/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)
    
- 
  [CreateFolder](http://msdn.microsoft.com/en-us/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)
    
- 
  [CreateItem](http://msdn.microsoft.com/en-us/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)
    
- 
  [FindConversation](http://msdn.microsoft.com/en-us/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)
    
- 
  [FindFolder](http://msdn.microsoft.com/en-us/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)
    
- 
  [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)
    
- 
  [GetConversationItems](http://msdn.microsoft.com/en-us/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)
    
- 
  [GetFolder](http://msdn.microsoft.com/en-us/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)
    
- 
  [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)
    
- 
  [MarkAsJunk](http://msdn.microsoft.com/en-us/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)
    
- 
  [MoveItem](http://msdn.microsoft.com/en-us/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)
    
- 
  [SendItem](http://msdn.microsoft.com/en-us/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)
    
- 
  [UpdateFolder](http://msdn.microsoft.com/en-us/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)
    
- 
  [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)
    
尝试使用不受支持的操作将导致出现错误响应。


## <a name="additional-resources"></a>其他资源



- [Outlook 外接程序的隐私、权限和安全性](../outlook/../../docs/develop/privacy-and-security.md)
    
- [将 Outlook 项目中的字符串作为已知实体进行匹配](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
