
# 更新适用于 Office 的 JavaScript API 的版本和清单架构文件



本文介绍了如何将 Office 外接程序项目中的 JavaScript 文件（Office.js 和特定于应用程序的 .js 文件）和外接程序清单验证文件更新到版本 1.1。

## 使用最新的项目文件

如果您使用 Visual Studio 来开发您的外接程序，以使用适用于 Office 的 JavaScript API 的 [最新 API 成员](../../reference/what's-changed-in-the-javascript-api-for-office.md)和 [外接程序清单 v1.1 功能](../../docs/overview/add-in-manifests.md)（根据 offappmanifest-1.1.xsd 进行了验证），则您需要下载并安装 [Visual Studio 2015 和最新的 Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs)。

如果您使用文本编辑器或 Visual Studio 以外的 IDE 开发您的 外接程序，则您需要针对在 外接程序 的清单中引用的 Office.js 和架构版本，将引用更新到 CDN。

若要运行使用新的和已更新的 Office.js API 和加载项清单功能开发的加载项，您的客户必须运行 Office 2013 SP1 或更高版本的本地产品，并在适用的情况下运行 SharePoint Server 2013 SP1 和相关的服务器产品、Exchange Server 2013 Service Pack 1 (SP1) 或相当于联机托管的产品：Office 365、SharePoint Online 和 Exchange Online。

若要下载 Office、SharePoint 和 Exchange SP1 产品，请参阅以下：


- [Microsoft Office 2013 和相关桌面产品的所有 Service Pack 1 (SP1) 更新的列表](http://support.microsoft.com/kb/2850036)
    
- [Microsoft SharePoint Server 2013 和相关服务器产品的所有 Service Pack 1 (SP1) 更新的列表](http://support.microsoft.com/kb/2850035)
    
- [Exchange Server 2013 Service Pack 1 的说明](http://support.microsoft.com/kb/2926248)
    

## 更新使用 Visual Studio 创建的 Office 外接程序项目以使用适用于 Office 的 JavaScript API 的最新库和 1.1 版本的外接程序清单架构


对于在适用于 Office 的 JavaScript API v1.1 和外接程序清单架构发布之前创建的项目，你可以使用“**NuGet 程序包管理器**”更新项目文件，然后更新外接程序的 HTML 页以进行引用。 

请注意，更新过程对于 _每个项目_ 执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，您需要重复更新过程。




### 将项目中适用于 Office 的 JavaScript API 库文件更新到最新版本


1. 在 Visual Studio 2015 中，打开或创建新的“**Office 外接程序**”项目。
    
      - 在左侧窗格中，选择“**更新**”并完成程序包更新过程。
    
  - 转到步骤 6。
    
2. 依次选择“**工具**” > “**NuGet 包管理器**” > “**管理解决方案的 Nuget 包**”。
    
3. 在“**NuGet 程序包管理器**”中，为“**程序包源**”选择“**nuget.org**”并为“**筛选器**”选择“**可用升级**”。 并选择 Microsoft.Office.js。
    
4. 在左侧窗格中，选择“**更新**”并完成程序包更新过程。
    
5. 在你的外接程序的 HTML 页的 **head** 标记中，注释掉或删除任何现有的 office.js 脚本引用。例如：`<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`，现在引用已更新的适用于 Office 的 JavaScript API 库，方法如下（将版本值更改为“1”）。 

   >**注意**在以下 CDN URL 中，在 office.js 前面的 /1/ 指定使用 Office.js 版本 1 中的最新增量版本。
    
```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


### 更新项目中的清单文件以使用架构版本 1.1


- 在项目的外接程序清单 (_projectname_ Manifest.xml) 文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 1.1（除 **xmlns** 属性以外的属性保持不变）。
    
```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```


>
  **注意**  在将外接程序清单架构的版本更新为 1.1 之后，你将需要删除 **Capabilities** 和 **Capability** 元素，并将其替换为 [Hosts 和 Host 元素](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx)或 [Requirements 和 Requirement 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

## 更新使用文本编辑器或其他 IDE 创建的 Office 外接程序项目，以使用适用于 Office 的 JavaScript API 的最新库和 1.1 版本的加载项清单架构


对于在发布适用于 Office 的 JavaScript API v1.1 和加载项清单架构之前创建的项目，您需要将加载项的 HTML 页更新到 v1.1 的 CDN 引用库中，将您的加载项清单文件更新为使用架构 v1.1。 

更新过程对_每个项目_分别执行，这意味着对于想要在其中使用 Office.js v1.1 的每个加载项项目以及加载项清单架构，你需要重复更新过程。

你不需要适用于 Office 的 JavaScript API 文件（Office.js 和特定于应用程序的.js 文件）的本地副本来开发 Office 加载项（在运行时引用 Office.js 的 CDN 会下载必要的文件），但如果你想要库文件的本地副本，你可以使用 [NuGet 命令行实用程序](http://docs.nuget.org/consume/installing-nuget)和 `Install-Package Microsoft.Office.js` 命令来下载它们。

 > **注意** 若要获取有关 v1.1 外接程序清单的 XSD（XML 架构定义）副本，请参阅 [Office 外接程序清单的架构参考 (v1.1)](../overview/add-in-manifests.md)) 中列出的内容。


### 将项目中适用于 Office 的 JavaScript API 库文件更新为使用最新版本


1. 在您的文本编辑器或 IDE 中打开您的加载项的 HTML 页。
    
2. 在你的外接程序的 HTML 页的 **head** 标记中，注释掉或删除任何现有的 office.js 脚本引用。例如：`<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`，现在引用已更新的适用于 Office 的 JavaScript API 库，方法如下（将版本值更改为“1”）。
    
```
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


    The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.
    

### 更新项目中的清单文件以使用架构版本 1.1


- 在项目的外接程序清单 ( _projectname_ Manifest.xml) 文件中，更新 **OfficeApp** 元素的 **xmlns** 属性，将版本值更改为 `1.1`（除  **xmlns** 属性以外的属性保持不变）。
    
```XML
<OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```

>
  **注意**在将外接程序清单架构的版本更新为 1.1 之后，你将需要删除 **Capabilities** 和 **Capability** 元素，并将其替换为 [Hosts 和 Host 元素](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx)或 [Requirements 和 Requirement 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md)。
    

## 其他资源



- [指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [了解 适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [适用于 Office 的 JavaScript API](../../reference/javascript-api-for-office.md)
    
- [Office 外接程序清单的架构参考 (v1.1)](../overview/add-in-manifests.md)
    
