
# <a name="troubleshoot-user-errors-with-office-add-ins"></a>解决 Office 外接程序中的用户错误

有时，您的用户在使用您开发的 Office 外接程序时可能会遇到问题。例如，外接程序无法加载或无法访问。使用本文中的信息有助于解决您的用户在使用 Office 外接程序时遇到的常见问题。 

您还可以使用 [Fiddler](http://www.telerik.com/fiddler) 识别和调试外接程序中的问题。

解决用户的问题后，您可以 [在 Office 应用商店中直接回复客户评论](https://msdn.microsoft.com/library/jj635874.aspx)。

## <a name="common-errors-and-troubleshooting-steps"></a>常见错误和故障排除步骤

下表列出了用户可能遇到的常见错误消息以及用户可以采取以解决这些错误的步骤。



|**错误消息**|**解决方案**|
|:-----|:-----|
|应用程序错误：无法访问目录|验证防火墙设置。"目录"指 Office 应用商店。此消息指示用户无法访问 Office 应用商店。|
|应用程序错误： 无法启动此应用程序。 关闭此对话框 忽略此问题或单击 "重新启动"以重试。|确认已安装最新的 Office 更新，或下载 [Office 2013 更新](https://support.microsoft.com/en-us/kb/2986156/)。|
|错误：对象不 支持此属性或方法 "defineProperty"|确认 Internet Explorer 不是在兼容模式下运行。转到“工具”>“**兼容性视图设置**”。|
|很抱歉，我们无法加载 该应用程序，因为您的浏览器 版本不受支持。 单击此处查看 支持的浏览器版本的列表。|确保浏览器支持 HTML5 本地存储，或重置您的 Internet Explorer 设置。有关受支持的浏览器的信息，请参阅 [运行 Office 加载项的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。|

## <a name="outlook-add-in-doesnt-work-correctly"></a>Outlook 外接程序不能正常工作

如果在 Windows 上运行的 Outlook 外接程序不能正常工作，请尝试在 Internet Explorer 中启用脚本调试。 


- 转到“工具”>“**Internet 选项**” > “**高级**”。
    
- 在“**浏览**”下，取消选中“**禁用脚本调试 (Internet Explorer)**”和“**禁用脚本调试 (其他)**”。
    
我们建议仅在解决问题时取消选中这些设置。如果你将其保持未选中状态，你在浏览时会收到提示。解决此问题后，再次选中“**禁用脚本调试(Internet Explorer)**”和“**禁用脚本调试(其他)**”。


## <a name="add-in-doesnt-activate-in-office-2013"></a>外接程序在 Office 2013 中无法激活

如果在用户执行下列步骤时外接程序无法激活：


1. 使用 Microsoft 帐户在 Office 2013 中登录。
    
2. 为其 Microsoft 帐户启用两步验证。
    
3. 尝试插入外接程序时在收到提示的时候验证其身份。
    
确认是否已安装最新的 Office 更新程序，或下载 [Office 2013 更新程序](https://support.microsoft.com/en-us/kb/2986156/)。

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a>外接程序无法在任务窗格中加载，或外接程序清单存在其他问题

请尝试使用[运行时日志记录](https://dev.office.com/docs/add-ins/develop/use-runtime-logging-to-debug-manifest)，针对外接程序清单存在的问题进行调试。

## <a name="additional-resources"></a>其他资源



- [调试 Office Online 中的外接程序](../testing/debug-add-ins-in-office-online.md)
    
- [将 Office 外接程序旁加载到 iPad 和 Mac 上](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [在 iPad 和 Mac 上调试 Office 外接程序](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
- [在 Visual Studio 中创建和调试 Office 外接程序](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
    
- [部署和安装 Outlook 外接程序以供测试](../outlook/testing-and-tips.md)
    
- [使用运行时日志记录调试清单](https://dev.office.com/docs/add-ins/develop/use-runtime-logging-to-debug-manifest)
    
