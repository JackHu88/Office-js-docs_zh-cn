
# <a name="outlook-addins"></a>Outlook 外接程序

Outlook 外接程序是由第三方使用基于 Web 技术的新平台构建到 Outlook 中的集成。Outlook 外接程序具有三个关键方面：


- 相同的外接程序和业务逻辑可在适用于 Windows 和 Mac 的桌面 Outlook、Web（Office 365 和 Outlook.com）和移动应用上使用。
    
-  Outlook 外接程序包括一个清单，其中介绍了如何将外接程序集成到 Outlook（例如，按钮或任务窗格）中，以及构成外接程序 UI 和业务逻辑的 JavaScript/HTML 代码。
    
- 最终用户或管理员可以从 Office 应用商店获取 Outlook 外接程序，也可以进行旁加载。
    
Outlook 外接程序与 COM 或 VSTO 外接程序（特定于在 Windows 上运行的 Outlook 的较早集成项）不同。与 COM 外接程序不同的是，Outlook 外接程序不具有任何实际安装到用户设备或 Outlook 客户端的代码。对于 Outlook 外接程序，Outlook 读取清单并挂钩在 UI 中指定的控件，然后加载 JavaScript 和 HTML。所有这些操作均在沙盒的浏览器的上下文中执行。

支持邮件外接程序的 Outlook 项目包括电子邮件、会议请求、响应和取消及约会。每个邮件外接程序均定义其可用的上下文，包括项目类型以及用户是在阅读还是撰写项目。


## <a name="extension-points"></a>扩展点


扩展点是外接程序与 Outlook 集成的方式。以下是执行此操作的方法：


- 外接程序可以声明出现在所有邮件和约会的命令界面中的按钮。有关详细信息，请参阅 [用于 Outlook 的外接程序命令](../outlook/add-in-commands-for-outlook.md)。
    
    **功能区上具有命令按钮的外接程序**

    ![外接程序命令无 UI 形状](../../images/41e46a9c-19ec-4ccc-98e6-a227283623d1.png)

- 外接程序可以在邮件和约会中中断与正则表达式匹配项或检测实体的链接。 有关详细信息，请参阅 [上下文 Outlook 外接程序](../outlook/contextual-outlook-add-ins.md)。
    
    **用于突出显示的实体（地址）的上下文相关外接程序**

    ![在卡片中显示上下文相关应用程序](../../images/59bcabc2-7cb0-4b9b-bb9f-06089dca9c31.png)


## <a name="mailbox-items-available-to-addins"></a>外接程序可用的邮箱项目


在撰写或阅读时，Outlook 外接程序对邮件或约会可用，但对其他项目类型不可用。如果撰写或阅读窗体中的当前邮件项目为以下项之一，则 Outlook 不会激活邮件外接程序：


- 使用信息权限管理 (IRM) 进行保护，采用 S/MIME 格式或使用其他保护方式进行加密。由于数字签名依赖于这些机制之一，数字签名邮件就是一个示例。
    
- 在“垃圾邮件”文件夹中。
    
- 具有邮件类别 IPM.Report.* 的送达报告或通知，包括送达和未送达报告 (NDR)，以及已读、未读和延迟通知。
    
- 属于其他邮件的附件的 .msg 文件。
    
- 从文件系统打开的 .msg 文件。
    
通常，Outlook 可以为“已发送邮件”文件夹中的项目在阅读窗体中激活外接程序，基于已知实体字符串匹配激活的外接程序除外。有关其背后的具体原因，请参阅[将 Outlook 项目中的字符串作为已知实体进行匹配](../outlook/match-strings-in-an-item-as-well-known-entities.md)中的“支持已知实体”。


## <a name="supported-hosts"></a>支持的主机


在 Outlook 2013 和更高版本、Outlook 2016 for Mac、Exchange 2013 内部环境中的 Outlook Web App、Office 365 和 Outlook.com 中的 Outlook Web App 中均支持 Outlook 外接程序。不是所有最新功能都会同时在所有客户端中受到支持。请参阅各个主题和 API 参考，以查看它们在哪些主机中不受支持。


## <a name="get-started-building-outlook-addins"></a>开始构建 Outlook 外接程序


若要开始构建 Outlook 外接程序，请参阅 [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)。


## <a name="additional-resources"></a>其他资源


- [Outlook 外接程序体系结构和功能概述](../outlook/overview.md)
- [开发 Office 外接程序的最佳做法](../../docs/overview/add-in-development-best-practices.md)
- [Office 外接程序的设计准则](../../docs/design/add-in-design.md)
- [许可 Office 和 SharePoint 外接程序](http://msdn.microsoft.com/library/3e0e8ff6-66d6-44ff-b0c2-59108ebd9181%28Office.15%29.aspx)
- [发布 Office 外接程序](../publish/publish.md)
- [将 Office 与 SharePoint 外接程序和 Office 365 Web 应用提交到 Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)

