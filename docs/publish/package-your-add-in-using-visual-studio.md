
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>使用 Visual Studio 打包外接程序以准备发布

您的 Office 外接程序包含一个您将用于发布外接程序的 XML 文件。您将需要单独发布项目的 Web 应用程序文件。


## <a name="deploy-your-web-project-and-package-your-add-in-by-using-visual-studio-2015"></a>部署 Web 项目并使用 Visual Studio 2015 打包外接程序



### <a name="to-deploy-your-web-project"></a>部署 Web 项目


1. 在“**解决方案资源管理器**”中，打开外接程序项目的快捷菜单，然后选择“**发布**”。
    
    将显示“**发布外接程序**”页。
    
2. 在“**当前配置文件**”下拉列表中，选择一个配置文件或选择“**新建…**”以创建一个新配置文件。
    
     >**注意**  发布配置文件指定你要部署到的服务器、登录服务器所需的凭据、要部署的数据库和其他部署选项。

    如果你选择“**新建...**”，将会显示“**创建发布配置文件**”向导。可以使用此向导从托管提供程序（如 Microsoft Azure）的网站导入发布配置文件，或创建新配置文件并添加你的服务器、凭据以及下一过程中的其他设置。
    
    有关导入发布配置文件或创建新发布配置文件的详细信息，请参阅 [创建发布配置文件](http://msdn.microsoft.com/en-us/library/dd465337.aspx#creating_a_profile)。
    
3. 在“**发布外接程序**”页中，选择“**部署 Web 项目**”链接。
    
    The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/en-us/library/dd465337.aspx).
    

### <a name="to-package-your-add-in"></a>打包您的加载项


1. 在“**发布外接程序**”页中，选择“**打包外接程序**”链接。
    
    将显示“**发布 Office 和 SharePoint 外接程序**”向导。
    
2. 在“**你的网站托管在何处?**”下拉列表中，选择或输入托管外接程序内容文件的网站 URL，然后选择“**完成**”。
    
    You have to specify an address that begins with the HTTPS prefix to complete this wizard. In general, using an HTTPS endpoint for your website is the best approach, but it is not required if you don't plan to publish your add-in to the Office Store. After the package is created, you can open the manifest in Notepad and replace the HTTPS prefix of your website with an HTTP prefix. For more information, see [Why do my add-ins have to be SSL-secured?](http://msdn.microsoft.com/en-us/library/jj591603#bk_q7). 
    
     >**注意**  Azure 网站自动提供 HTTPS 终结点。

    Visual Studio 生成发布外接程序所需的文件，然后打开发布输出文件夹。 
    
如果计划将外接程序提交到 Office 应用商店，可以选择“**执行验证检查**”链接以确定将阻止外接程序被接受的问题。应先解决所有问题，再将外接程序提交到应用商店。

现在，您可以将您的 XML 清单上载到适当位置以 [发布外接程序](../publish/publish.md)。您可以在  `OfficeAppManifests` 文件夹的 `app.publish` 中找到 XML 清单。例如：

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="additional-resources"></a>其他资源



- [发布 Office 外接程序](../publish/publish.md)
    
- [将 Office 与 SharePoint 外接程序和 Office 365 Web 应用提交到 Office 应用商店](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
