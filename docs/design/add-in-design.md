# <a name="design-guidelines-for-office-add-ins"></a>Office 外接程序的设计准则

Office 外接程序可通过提供用户可在 Office 客户端内访问的上下文功能来扩展 Office 体验。通过外接程序，用户可以访问 Office 内的第三方功能以完成更多操作，而无需进行成本高昂的上下文切换。 

 您的外接程序 UX 设计必须与 Office 无缝集成，为用户提供高效、自然的交互。利用外接程序命令（Office UI 扩展）提供对外接程序的访问权限，并使用创建基于 HTML 的自定义 UI 时的建议 [UI 元素](ui-elements/ui-elements.md)和[最佳实践](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)。 
 
 
## <a name="core-office-add-in-design-principles"></a>Office 外接程序的核心设计准则
无论您用于创建自定义 UI 的基础框架是什么，在设计外接程序时请应用以下准则： 

- **对 Office 进行明确设计**。外接程序的功能和外观必须和谐地补充 Office 体验，包括应用 Office 或文档主题。
 
- **提高用户的工作效率**。帮助用户在不影响其他工作的情况下完成一项工作。允许在 Office 文档和外接程序之间进行无缝交互。 

- **使内容优先于 Chrome**。强调外接程序的内容和功能优先于任何辅助 Chrome。通过避免不会为用户体验增加价值的多余 UI 元素，最大限度地利用空间。  

- **始终保持用户的控制权**。允许用户控制体验、了解任何重要决策，并轻松撤消外接程序执行的操作。 

- 
  **针对所有平台和输入方法进行设计**。外接程序设计用于 Office 支持的所有平台，您的外接程序 UI 应该进行优化，以便跨平台和外形规格运行。支持鼠标/键盘和触摸输入设备，确保您的自定义 HTML UI 响应迅速，可适应不同的外形规格。有关详细信息，请参阅[触摸](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#bk_Touch)。 


## <a name="design-language"></a>设计语言
我们建议你采用 Office 设计语言，并使用 [Office UI Fabric](https://dev.office.com/fabric) 创建基于 HTML 的自定义体验。如果你的组织已有设计语言，欢迎你使用，只要最终结果对 Office 用户来说是和谐的体验。 


## <a name="add-in-building-blocks"></a>外接程序构建块
您可以使用两种类型的 UI 元素来创建外接程序： 

- [外接程序命令](ui-elements/ui-elements.md#add-in-commands)使你可以在 Office 应用程序中添加本机 UX 挂钩
- [基于 HTML 的自定义 UI](ui-elements/ui-elements.md#custom-html-based-ui) 使你可以利用 Office 客户端内的 HTML 功能。 

有关如何使用这些构建块的详细信息，请参阅 [UI 元素](ui-elements/ui-elements.md)。  

## <a name="ux-design-patterns"></a>UX 设计模式

为了帮助你为外接程序创建一流的用户体验，我们提供了模版，以演示常规 UX 设计模式。这些模板反映了创建一流的外接程序的[最佳实践](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices)，包括首次运行体验、品牌元素和用户通知的模式。它们使用 [Office UI 结构](https://dev.office.com/fabric)组件和样式，并包含可轻松扩展 Office UI 的元素。

要访问此模板，请参阅 [Office 外接程序 UX 设计模式](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns)报告。此外还提供了 Adobe Illustrator 文件，你可以下载并更新它们，以反映你自己的设计。你还可以将 [Office 外接程序 UX 设计模式代码](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)报告中的代码文件复制到你的加载项项目，并根据需要将其自定义。 

## <a name="recommended-layouts-and-interaction-patterns"></a>建议布局和交互模式
我们为每种外接程序类型提供建议布局，以及帮助您实现所有操作的**端到端**示例。若要了解有关如何对外接程序进行布局的详细信息，请参阅以下内容：

- [任务窗格容器的布局](ui-elements/layout-for-task-pane-add-ins.md)
- [内容外接程序的布局](ui-elements/layout-for-content-add-ins.md) 
- [邮件外接程序的布局](ui-elements/layouts-for-outlook-add-ins.md)

另请参阅交互模式，获取外接程序的常规方案及其相应的交互模式的示例。

## <a name="additional-resources"></a>其他资源

- [Office UI Fabric](https://dev.office.com/fabric) 

