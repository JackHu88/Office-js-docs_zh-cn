
# Excel、Word 和 PowerPoint 的外接程序命令

外接程序命令是一些能够扩展 Office UI，并在外接程序中启动操作的 UI 元素。您可以在功能区中添加按钮，或者向上下文菜单中添加项目。当用户选择外接程序命令时，它们将启动一些操作，例如运行 JavaScript 代码，或在任务窗格中显示外接程序页。外接程序命令可以帮助用户查找和使用您的外接程序，这可以帮助提高您外接程序的利用率和重复性使用，进而提高客户对它的保留率。

有关此功能的概述，请参阅视频 [Office 功能区中的外接程序命令](https://channel9.msdn.com/events/Build/2016/P551)。


**命令在 Excel Desktop 中运行的外接程序**
![外接程序命令](../../images/addincommands1.png)

**命令在 Excel Online 中运行的外接程序**
![外接程序命令](../../images/addincommands2.png)

## 命令功能
目前支持下列命令功能。

**扩展点**

- 功能区选项卡 - 扩展内置选项卡或新建自定义选项卡。
- 上下文菜单 - 扩展所选上下文菜单。 

**控件类型**

- 简单按钮 - 触发特定操作。
- 菜单 - 包含触发操作的多个操作。

**操作**

- ShowTaskpane - 显示一个或多个在其中加载自定义 HTML 页面的窗格。
- ExecuteFunction - 加载一个不可见的 HTML 页，然后在其中执行一个 JavaScript 函数。 若要在你的函数（例如错误、进度、其他输入）中显示 UI，你可以使用 [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui) API。  

## 支持的平台
目前在以下平台上支持外接程序命令：

- Office for Windows Desktop 2016（版本 16.0.6769.0000 或更高版本）
- 含个人帐户的 Office Online
- 含工作/学校帐户的 Office Online（预览）

即将推出更多平台。

## 开始使用外接程序命令

有关如何在清单中指定外接程序命令的详细信息，请参阅[在清单中定义外接程序命令](http://dev.office.com/docs/add-ins/outlook/manifests/define-add-in-commands)。

若要开始使用外接程序命令，请参阅 GitHub 上的 [Office 外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)。





