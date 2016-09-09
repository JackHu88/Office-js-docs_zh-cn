
# 使用 Napa 或 Visual Studio 打包外接程序以准备发布

您的 Office 外接程序包含一个您将用于发布外接程序的 XML 文件。您将需要单独发布项目的 Web 应用程序文件。

## 使用 Napa 打包您创建的 Office 外接程序



1. 在 Napa 中，在页面的一侧，选择“**发布**”按钮（![发布按钮](../../images/Apps_NAPA_Publish.png)）。
    
2. 在“**发布设置**”对话框中，选择“**下一步**”。
    
3. 提供托管外接程序内容文件的网站 URL（例如，项目的默认 HTML 和 JavaScript 文件），然后选择“**发布**”。
    
4. 在“**发布成功**”对话框中，选择“**发布位置**”链接。
    
    将显示文档库，其中包含外接程序的 XML 清单文件和 Web 内容文件。 
    
接下来，手动将（样式表、JavaScript 文件以及 HTML 文件）的 Web 内容文件复制到 Web 服务器，该服务器托管了在“**发布设置**”对话框中提供的网站。

现在，您可以将您的 XML 清单上载到适当位置以 [发布外接程序](../publish/publish.md)。 


## 部署 Web 项目并使用 Visual Studio 2015 打包您的外接程序



### 部署 Web 项目


1. 在“**解决方案资源管理器**”中，打开外接程序项目的快捷菜单，然后选择“**发布**”。
    
    将显示“**发布外接程序**”页。
    
2. 在“**当前配置文件**”下拉列表中，选择一个配置文件或选择“**新建…**”以创建一个新配置文件。
    
     >**注意**  发布配置文件指定你要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。

    如果你选择“**新建...**”，将会显示“**创建发布配置文件**”向导。 可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。
    
    有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅 [创建发布配置文件](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile)。
    
3. 在“**发布外接程序**”页中，选择“**部署 Web 项目**”链接。
    
    将显示“**发布 Web**”对话框。 有关使用此向导的详细信息，请参阅 [操作说明：在 Visual Studio 中使用点击发布来部署 Web 项目](http://msdn.microsoft.com/en-us/library/dd465337.aspx)。
    

### 打包你的外接程序


1. 在“**发布外接程序**”页中，选择“**打包外接程序**”链接。
    
    将显示“**发布 Office 和 SharePoint 外接程序**”向导。
    
2. 在“**你的网站托管在何处?**”下拉列表中，选择或输入托管外接程序内容文件的网站 URL，然后选择“**完成**”。
    
    你必须指定以 HTTPS 前缀开头的地址来完成此向导。 一般情况下，为你的网站使用 HTTPS 终结点是最好的方法，但如果你不打算将外接程序发布到 Office 应用商店则不需要这样做。 在创建包后，你可以在记事本中打开此清单，并使用 HTTP 前缀替换你的网站的 HTTPS 前缀。 有关详细信息，请参阅 [为什么我的外接程序必须采用 SSL 保护?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7)。 
    
     >**注意**  Azure 网站自动提供 HTTPS 终结点。

    Visual Studio 生成发布外接程序所需的文件，然后打开发布输出文件夹。 
    
如果计划将外接程序提交到 Office 应用商店，可以选择“**执行验证检查**”链接以确定将阻止外接程序被接受的问题。 应先解决所有问题，再将外接程序提交到应用商店。

现在，您可以将您的 XML 清单上载到适当位置以 [发布外接程序](../publish/publish.md)。您可以在  `OfficeAppManifests` 文件夹的 `app.publish` 中找到 XML 清单。例如：

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## 其他资源



- [发布 Office 外接程序](../publish/publish.md)
    
- [将 Office 与 SharePoint 外接程序和 Office 365  Web 应用提交到 Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
