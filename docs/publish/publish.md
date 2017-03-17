
# <a name="deploy-and-publish-your-office-add-in"></a>部署和发布 Office 外接程序

可以使用几种方法之一来部署 Office 外接程序，以用于对用户进行测试或分发：

|**方法**|**Use...**|
|:---------|:------------|
|[旁加载](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|作为测试在 Windows、Office Online、iPad 或 Mac 上运行的外接程序的开发过程的一部分。|
|[Office 365 管理中心（预览）](#office-365-admin-center-preview)|在云或混合部署中，用于向组织中的用户分发外接程序。|
|[Office 应用商店]|用于向用户公开分发外接程序。|
|[SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|在本地环境中，用于向组织中的用户分发外接程序。|
|[Exchange 服务器](#outlook-add-in-deployment)|在本地或联机环境中，用于向用户分发 Outlook 外接程序。|

可用的选项具体取决于你定位的 Office 主机以及所创建的外接程序的类型。

>**注意：**如果计划将外接程序发布到 Office 应用商店，请务必遵循 [Office 应用商店验证策略](https://msdn.microsoft.com/en-us/library/jj220035.aspx)。例如，外接程序必须适用于支持你定义的方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 外接程序主机和可用性](https://dev.office.com/add-in-availability)页）。

## <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Word、Excel 和 PowerPoint 外接程序的部署选项

| 扩展点            | 旁加载 | Office 365 管理中心（预览） |Office 应用商店| SharePoint 目录*  |
|:----------------|:-----------:|:------------------:|:-------------------------------:|:------------:|
| 内容         | X           | X                  | X                               | X|
| 任务窗格       | X           | X                  | X                               | X|
| 命令           | X           | X                  | X                               |  |

&#42; SharePoint 目录不支持 Office 2016 for Mac。

## <a name="deployment-options-for-outlook-add-ins"></a>Outlook 外接程序的部署选项

| 扩展点     | 旁加载 | Exchange 服务器 | Office 应用商店 |
|:---------|:-----------:|:---------------:|:------------:|
| 邮件应用 | X           | X               | X            |
| 命令  | X           | X               | X            |


有关最终用户如何获取、插入和运行外接程序的信息，请参阅[开始使用你的 Office 外接程序](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE)。

## <a name="office-365-admin-center-preview-deployment"></a>Office 365 管理中心（预览）部署

通过 Office 365 管理中心，管理员可以轻松地将 Word、Excel 和 PowerPoint 外接程序部署到组织内的用户或组。通过管理中心部署的外接程序可立即供 Office 应用程序中的用户使用，而无需进行客户端配置。可以通过管理中心部署内部外接程序以及 ISV 提供的外接程序。

管理中心当前支持以下方案：

- 个人、组或组织新的和更新的外接程序的集中部署。
- 支持多个平台，其中包括 Windows 和 Office Online，即将推出对 Mac 的支持。
- 到英语语言租户和全球范围租户的部署。
- 云托管的外接程序部署。
- 在启动 Office 应用程序时自动安装。
- 在防火墙内托管的外接程序 URL。
- Office 应用商店外接程序的部署（即将推出）。

<!--
The admin center also includes a pre-deployment validation checking service.
-->

外接程序部署方案中的未来投入重点为 Office 365 管理中心。如果组织满足先决条件，我们建议使用管理中心将外接程序部署到组织。

### <a name="prerequisites-for-admin-center-deployment"></a>管理中心部署的先决条件 

如果组织满足以下条件，则可以通过管理中心部署外接程序：

- 用户运行 Office 2016 内部版本 7070 或更高版本。
- 用户使用其工作或学校帐户登录 Office 2016。
- 组织使用 Azure Active Directory (Azure AD) 标识服务。

管理中心不支持以下内容：

- 面向 Office 2013 中 Word、Excel 或 PowerPoint 的外接程序。
- 本地目录服务。
- SharePoint 外接程序部署。
- 到 Office Online Server 的外接程序部署。
- COM/VSTO 外接程序的部署。

若要部署 SharePoint 外接程序或面向 Office 2013 的外接程序，请使用 [SharePoint 外接程序目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

>**重要说明！**SharePoint 外接程序目录不支持在外接程序清单的 [VersionOverrides](../../reference/manifest/versionoverrides.md) 节点中实现的外接程序功能，如[外接程序命令](../design/add-in-commands.md)。 

若要部署 COM/VSTO 外接程序，请使用 ClickOnce 或 Windows Installer。有关详细信息，请参阅[部署 Office 解决方案](https://msdn.microsoft.com/en-us/library/bb386179.aspx)。

## <a name="sharepoint-catalog-deployment"></a>SharePoint 目录部署

SharePoint 外接程序目录是可以创建以托管 Word、Excel 和 PowerPoint 外接程序的特殊网站集。因为 SharePoint 目录不支持在清单的 VersionOverrides 节点中实现的新外接程序功能（包括外接程序命令），如果可能，建议通过管理中心（预览）来使用集中部署。默认情况下，在任务窗格中打开通过 SharePoint 目录部署的外接程序。

如果要在本地环境中部署外接程序，请使用 SharePoint 目录。有关详细信息，请参阅[将任务窗格和内容外接程序发布到 SharePoint 目录](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。

> **注意：**SharePoint 目录不支持 Office 2016 for Mac。若要向 Mac 客户端部署 Office 外接程序，必须将其提交到 [Office 应用商店]。 

## <a name="outlook-add-in-deployment"></a>Outlook 外接程序部署

对于不使用 Azure AD 标识服务的本地和联机环境，可以通过 Exchange 服务器部署 Outlook 外接程序。 

Outlook 外接程序部署需要以下内容：

- Office 365、Exchange Online 或 Exchange Server 2013 或更高版本
- Outlook 2013 或更高版本

若要将外接程序分配给租户，请使用 Exchange 管理中心从文件或 URL直接上载清单，或从 Office 应用商店添加外接程序。若要将外接程序分配给单个用户，则必须使用 Exchange PowerShell。有关详细信息，请参阅 TechNet 上的[安装或删除组织的 Outlook 外接程序](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx)。


## <a name="additional-resources"></a>其他资源

- [部署和安装 Outlook 外接程序以进行测试](../outlook/testing-and-tips.md) 
- [将外接程序和 Web 应用提交到 Office 应用商店] [Office 应用商店]
- [Office 外接程序的设计准则](../design/add-in-design)
- [创建有效的 Office 应用商店外接程序](https://msdn.microsoft.com/en-us/library/jj635874.aspx)
- [解决 Office 外接程序中的用户错误](../testing/testing-and-troubleshooting.md)

[Office 应用商店]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
