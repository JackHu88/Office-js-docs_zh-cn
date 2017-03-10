
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a>将任务窗格和内容外接程序发布到 SharePoint 目录

>**重要说明！**SharePoint 上的外接程序目录不支持在[外接程序清单](../overview/add-in-manifests.md)的 VersionOverrides 节点中实现的外接程序功能，如外接程序命令。 

>如果面向云或混合环境，我们建议通过 [管理中心预览](https://support.office.com/en-ie/article/Deploy-Office-Add-ins-in-the-Office-365-new-Admin-Center-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE)使用**集中部署**来发布外接程序。

外接程序目录是 SharePoint Web 应用程序或 SharePoint Online 租户中的专用网站集合，它托管 Office 和 SharePoint 外接程序的文档库。管理员可以将 Office 外接程序清单文件上载到外接程序目录以供组织使用。管理员将外接程序目录注册为受信任的目录时，用户可从 Office 客户端应用程序中的插入 UI 中插入外接程序。

SharePoint 目录不支持 Office 2016 for Mac。若要向 Mac 客户端部署 Office 外接程序，必须将其提交到 [Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)。   

## <a name="to-set-up-an-add-in-catalog-on-sharepoint"></a>在 SharePoint 上设置外接程序目录

1. 浏览到“**管理中心网站**”（“**开始**” > “**所有程序**” > “**Microsoft SharePoint 2013 产品**” > “**SharePoint 2013 管理中心**”）。
    
2. 在左侧的任务窗格中，选择“**外接程序**”。
    
3. 在“**外接程序**”页面的“**外接程序管理**”下，选择“**管理外接程序目录**”。
    
4. 在“**管理外接程序目录**”页上，确保在“**Web 应用程序选择器**”中选择了正确的 Web 应用程序。
    
5. 选择“**查看网站设置**”。
    
6. 在“**网站设置**”页上选择“**网站集管理员**”以指定网站集管理员，然后选择“**确定**”。
    
7. 若要向用户授予网站权限，请选择“**网站权限**”，然后选择“**授予权限**”。
    
8. 在“**共享‘应用程序目录网站’**”对话框中，指定一个或多个网站用户，为他们设置相应的权限，选择性地设置其他选项，然后选择“**共享**”。
    
9. 若要向 Office 外接程序外接程序目录添加外接程序，请选择“**Office 外接程序**”。

## <a name="to-set-up-an-add-in-catalog-on-office-365"></a>在 Office 365 上设置外接程序目录

1. 在“Office 365 管理中心”页上，选择“**管理**”，然后选择“**SharePoint**”。
    
2. 在左侧的任务窗格中，选择“**外接程序**”。
    
3. 在“**外接程序**”页上，选择“**外接程序目录**”。
    
4. 在“**外接程序目录网站**”页上，选择“**确定**”以接受默认选项，并新建外接程序目录网站。
    
5. 在“**创建外接程序目录网站集**”页上，指定外接程序目录网站的标题。
    
6. 指定网站地址。
    
7. 将“**存储配额**”设置为可能的最低值（当前为 110）。你将仅在该网站集上安装外接程序包，它们非常小。
    
8. 将“**服务器资源配额**”设置为 0（零）。（服务器资源配额与限制性能不佳的沙盒解决方案有关，但你不会在外接程序目录网站上安装任何沙盒解决方案。）
    
9. 选择“**确定**”。
    
若要将外接程序添加到外接程序目录网站，请浏览至已创建的网站。在左侧导航窗格中，选择“**Office 外接程序**”，然后选择“**新外接程序**”以上传 Office 外接程序清单文件。    

## <a name="publish-to-an-add-in-catalog"></a>发布到外接程序目录


1. 浏览至外接程序目录：

    1- 打开 SharePoint 管理中心主页。
    
    2- 选择“**外接程序**”。
    
    3- 选择“**管理外接程序目录**”。
    
    4- 选择提供的链接，然后选择左侧导航栏上的“**Office 外接程序**”。
    
2. 选择“**单击以添加新项目**”链接。
    
3. 选择“**浏览**”，然后指定要上传的[清单](../../docs/overview/add-in-manifests.md)。
    
    此目录中的内容和任务窗格外接程序现在可从“**Office 外接程序**”对话框提供。若要访问这些外接程序，请在“**插入**”选项卡上选择“**我的外接程序**”，然后选择“**我的组织**”。
    
将外接程序清单上载到 Office 外接程序 目录后，用户可以通过执行下列操作来访问外接程序：


1. 在 Office 应用程序中，转到“**文件**” > “**选项**” > “**信任中心**” > “**信任中心设置**” > “**受信任的外接程序目录**”。
    
2. 指定外接程序目录的 _父级 SharePoint 网站集_ 的 URL。例如，如果 Office 外接程序 目录的 URL 是：
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    仅指定父网站集的 URL：
    
    `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. 关闭并重新打开 Office 应用程序。外接程序目录会出现在“**Office 外接程序**”对话框中。
    
或者，管理员可以通过使用组策略来指定 SharePoint 上的 Office 外接程序目录。有关详细信息，请参阅 TechNet 提供的 [Office 外接程序概述](https://technet.microsoft.com/en-us/library/jj219429.aspx)中的"使用组策略管理用户安装和使用 Office 相关外接程序的方式"一节。

