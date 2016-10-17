
# <a name="create-outlook-add-ins-for-read-forms"></a>创建适用于阅读窗体的 Outlook 外接程序

阅读外接程序是在 Outlook 中的阅读窗格或阅读检查器中激活的 Outlook 外接程序。与撰写外接程序（用户创建邮件或约会时激活的 Outlook 外接程序）不同，阅读外接程序在以下用户方案中可用：


- 查看电子邮件、会议请求、会议响应或会议取消。*
    
- 查看用户参与的会议项目。
    
- 查看用户作为组织者的会议项目（仅限 Outlook 2013 和 Exchange 2013 的 RTM 版本）。
    
     >**注意**  从 Office 2013 SP1 版本开始，如果用户查看由用户组织的会议项目，则只有撰写外接程序才能够激活并可用。该方案中不再提供读取外接程序。
* Outlook 不激活针对特定邮件类型的外接程序阅读窗体，包括另一封邮件附件的项目、Outlook"草稿"或"垃圾邮件"文件夹中的项目，或以其他方式加密或保护的项目。

在每个阅读方案中，满足激活条件时 Outlook 便会激活外接程序，用户可以选择和打开阅读窗格或阅读检查器中外接程序栏上的已激活外接程序。图 1 显示用户阅读包含地理地址的邮件时激活和打开的“**必应地图**”外接程序。


**图 1.对于所选的包含地址的 Outlook 邮件，显示必应地图外接程序的外接程序窗格处于活跃状态**

![Outlook 中的必应地图邮件应用程序](../../images/off15appsdk_BingMapMailAppScreenshot.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a>阅读模式下可用的外接程序的类型


阅读外接程序可以为下列类型的任意组合。


- [适用于 Outlook 的外接程序命令](../outlook/add-in-commands-for-outlook.md)
    
- [上下文 Outlook 外接程序](../outlook/contextual-outlook-add-ins.md)
    
- [自定义窗格 Outlook 外接程序](../outlook/custom-pane-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a>阅读外接程序可用的 API 功能


有关适用于 Office 的 JavaScript API 为阅读窗体中的 Outlook 外接程序提供的功能列表，请参阅 [每个版本的邮件应用程序](http://msdn.microsoft.com/library/f34e2f44-8c9d-4e90-b1d7-3f29506adb92%28Office.15%29.aspx)中的表 1 和表 2。 

另请参阅：


- 对于激活阅读窗体中的外接程序：请参阅 [在清单中指定激活规则](../outlook/manifests/activation-rules.md#specify-activation-rules-in-a-manifest) 中的表 1。
    
- [使用正则表达式激活规则显示 Outlook 外接程序](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [将 Outlook 项目中的字符串作为已知实体进行匹配](../outlook/match-strings-in-an-item-as-well-known-entities.md)
    
- [从 Outlook 项目中提取实体字符串](../outlook/extract-entity-strings-from-an-item.md)
    
- [从服务器获取 Outlook 项目的附件](../outlook/get-attachments-of-an-outlook-item.md)
    

## <a name="additional-resources"></a>其他资源



- [适用于 Office 365 的 Outlook 外接程序入门](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
