
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>部署和安装 Outlook 外接程序以进行测试


作为开发 Outlook 外接程序的一个环节，您可能会发现自己在反复部署和安装外接程序以进行测试，这会涉及以下步骤：


1. 创建描述外接程序的清单文件。
    
2. 将外接程序 UI 文件部署到 Web 服务器。
    
3. 在邮箱中安装外接程序。
    
4. 测试外接程序，对 UI 或清单文件进行适当的更改，重复步骤 2 和 3 来测试所做更改。
    

## <a name="creating-a-manifest-file-for-the-add-in"></a>创建外接程序的清单文件

每个外接程序都通过一个 XML 清单进行描述，该文档为服务器提供有关外接程序的信息，为用户提供外接程序的描述性信息，并标识外接程序 UI HTML 文件的位置。您可以在本地文件夹或服务器上存储该清单，只要所测试的邮箱的 Exchange 服务器能够访问这个位置即可。我们假定您在本地文件夹中存储清单。有关如何创建清单文件的信息，请参阅 [Outlook 外接程序清单](../outlook/manifests/manifests.md)。 


## <a name="deploying-an-add-in-to-a-web-server"></a>将外接程序部署到 Web 服务器

可以使用 HTML 和 JavaScript 创建外接程序 UI。生成的源文件存储在承载该外接程序的 Exchange 服务器可以访问的 Web 服务器上。该源文件由外接程序清单文件中指定的 **DesktopSettings** 元素、 [TabletSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx) 元素和/或 [PhoneSettings](http://msdn.microsoft.com/en-us/library/5c89cc7c-7ae0-49c9-fdd5-4c52118228f6%28Office.15%29.aspx) 元素中的 [SourceLocation](http://msdn.microsoft.com/en-us/library/13e4eae3-8e8c-fd55-a1c2-3297b485f327%28Office.15%29.aspx) 子元素标识。

在最初部署外接程序的 UI 文件后，可以通过将 Web 服务器上存储的 HTML 文件替换为 HTML 文件的新版本以更新外接程序 UI 和行为。


## <a name="installing-the-add-in"></a>安装外接程序


准备好外接程序清单文件并将外接程序 UI 部署到可以访问的 Web 服务器后，可以使用 Outlook 富客户端、Outlook Web App 或 适用于设备的 OWA 或通过运行远程 Windows PowerShell cmdlet 为 Exchange 服务器上的邮箱安装外接程序。


### <a name="installing-an-add-in-in-an-outlook-rich-client"></a>在 Outlook 富客户端中安装外接程序

如果你的邮箱位于 Exchange Online、Exchange 2013 或更高版本，则可安装外接程序。在 Outlook for Windows 中，你可以通过 Office Fluent Backstage 视图安装外接程序。依次选择“**文件**”和“**管理外接程序**”。可以登录 Exchange 管理中心。登录后，继续进行下一节中第 4 步的安装过程。

在 Outlook for Mac 中，选择外接程序栏最右侧的“**管理外接程序**”，然后登录到 Exchange 管理中心。继续进行下一节中的第 4 步。


### <a name="installing-an-add-in-by-using-outlook-web-app-or-outlookcom"></a>使用 Outlook Web App 或 Outlook.com 安装外接程序

若要使用 Outlook Web App (OWA) 安装 Outlook 外接程序，请按照下列步骤操作：


1. 浏览到组织的 OWA URL 或 Outlook.com 并登录。
    
2. 选择右上角的齿轮图标并选择“**管理外接程序**”。
    
3. 选择加号 ( **+**) 添加新的外接程序。
    
4. 假定你将清单存储在本地文件夹中，则从下拉列表中，选择“**从文件添加**”。
    
5. 浏览到清单的文件路径，然后选择“**安装**”。
    
6. 选择窗口右上角的用户名，然后选择“**我的邮件**”切换到你的电子邮件以测试外接程序。
    

如果你没有 Exchange Server 最低的“我的自定义应用”角色，则只能从 Office 应用商店安装外接程序。为了测试外接程序或通过指定外接程序清单的 URL 或文件名安装外接程序，应请求你的 Exchange 管理员提供所需的权限。

Exchange 管理员可运行以下 PowerShell cmdlet 向单个用户分配必要的权限。在此示例中，wendyri 是用户的电子邮件别名。

```New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"```

如有必要，管理员可运行以下 cmdlet 向多个用户分配类似的必要权限：

```$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}```

有关我的自定义应用角色的详细信息，请参阅[我的自定义应用角色](http://technet.microsoft.com/en-us/library/aa0321b3-2ec0-4694-875b-7a93d3d99089%28EXCHG.150%29.aspx)。 

使用 Office 365 或 Visual Studio 开发外接程序会向你分配组织管理员角色，这便允许你按 EAC 中的文件或 URL 或者按 Powershell cmdlet 安装外接程序。


### <a name="installing-an-add-in-by-using-remote-powershell"></a>使用远程 PowerShell 安装外接程序

在 Exchange 服务器上创建远程 Windows PowerShell 会话后，可以使用  **New-App** cmdlet 及以下 PowerShell 命令安装 Outlook 外接程序。


```
New-App -URL:"http://<fully-qualified URL">
```

完全限定的 URL 是您为加载项准备的加载项清单文件的位置。

可以使用下列附加 PowerShell cmdlet 来管理邮箱的加载项：


-  **Get-App** —  列出为邮箱启用的外接程序。
    
-  **Set-App** — 启用或禁用邮箱上的外接程序。
    
-  **Remove-App** — 从 Exchange 服务器删除以前安装的外接程序。
    

## <a name="additional-resources"></a>其他资源



- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
- [解决 Office 外接程序中的用户错误](../testing/testing-and-troubleshooting.md)
    
