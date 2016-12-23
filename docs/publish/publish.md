
# <a name="deploy-and-publish-your-office-add-in"></a>部署和发布 Office 外接程序

可以使用几种方法之一来部署 Office 外接程序，以用于对用户进行测试或分发： 

- [旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) - 用作测试在 Windows、Office Online、iPad 或 Mac 上运行的外接程序的开发过程的一部分。
- [Office 365 管理中心预览](https://support.office.com/en-ie/article/Deploy-Office-Add-Ins-in-Office-365-737e8c86-be63-44d7-bf02-492fa7cd9c3f?ui=en-US&rs=en-IE&ad=IE) - 用于向云或混合部署的组织中的用户分发你的外接程序。
- [Office 应用商店] - 用于向用户公开分发你的外接程序。
- [SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) - 用作测试外接程序或在本地环境中向组织中的用户分发外接程序的开发过程的一部分。

具体什么选项可用取决于你定位的 Office 主机以及所创建的外接程序的类型。

>**注意：**生成外接程序时，如果计划将外接程序[发布](../publish/publish.md)到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

## <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel 和 PowerPoint 外接程序的部署选项

| 扩展点            | 旁加载 | Office 365 管理中心预览 |Office 应用商店  | SharePoint 目录  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| 内容         | X           | X                  | X                               | X            |
| 任务窗格       | X           | X                  | X                               | X            |
| 命令         | X           | X                  | X                                |              |

> **注意：**SharePoint 目录不支持 Office 2016 for Mac。若要向 Mac 客户端部署 Office 外接程序，必须将其提交到 [Office 应用商店]。    

## <a name="deployment-options-for-outlook-add-ins"></a>Outlook 外接程序的部署选项

| 扩展点     | 旁加载 | Exchange 服务器 | Office 应用商店 |
|:---------|:-----------:|:---------------:|:------------:|
| 邮件应用 | X           | X               | X            |
| 命令  | X           | X               | X            |

若要确保你的外接程序能够覆盖更多的最终用户，请确保它能够跨平台正常运行。Office 外接程序支持 Windows、Mac、Web、iOS 和 Android。有关每个平台所支持功能的概述，请参阅 [Office 外接程序主机和平台可用性]。   

有关许可 Office 应用商店外接程序的信息，请参阅[许可外接程序](https://msdn.microsoft.com/EN-US/library/office/jj163257.aspx)。

有关最终用户如何获取、插入和运行外接程序的信息，请参阅 [开始使用你的 Office 外接程序](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。

## <a name="additional-resources"></a>其他资源

- [Office 外接程序主机和平台可用性]
- [部署和安装 Outlook 外接程序以进行测试](../outlook/testing-and-tips.md) 
- [将外接程序和 Web 应用提交到 Office 应用商店] [Office 应用商店]
- [Office 外接程序的设计准则](../design/add-in-design)
- [创建了有效的 Office 应用商店外接程序](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [解决 Office 外接程序中的用户错误](../testing/testing-and-troubleshooting.md)

[Office 应用商店]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office 外接程序主机和平台可用性]: http://dev.office.com/add-in-availability
