-
#<a name="use-office-ui-fabric-in-office-add-ins"></a>在 Office 外接程序中使用 Office UI Fabric

若要生成 Office 外接程序，我们建议使用 [Office UI Fabric](https://dev.office.com/fabric) 生成用户体验。 

Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。Fabric 提供了以视觉对象为中心的组件，可在 Office 外接程序中进行扩展、返工和使用。由于 Fabric 使用的是 Office 设计语言，因此 Fabric 的用户体验组件看起来像是 Office 的自然扩展。

Fabric 包含以下多个项目：

- **Fabric JS（推荐）**- 仅使用 JavaScript 实现用户体验组件。如果不想依赖 React 框架，我们建议使用此版 Fabric。  
- **Fabric React** - 使用 React 框架实现用户体验组件。
- **Fabric Core** - 包含设计语言的核心元素，如图标、颜色、铅字和网格等。Fabric JS 和 Fabric React 均使用 Fabric Core。 

下面逐步介绍了使用 Fabric JS 的基础知识。  

##<a name="1-add-the-fabric-cdn-references"></a>1.添加 Fabric CDN 引用
若要从 CDN 引用 Fabric，请在页面中添加以下 HTML 代码。

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.2.0/js/fabric.min.js"></script>

就是这么简单。现在，可以开始在外接程序中使用 Fabric 了。 

##<a name="2-use-fabric-icons-and-fonts"></a>2.使用结构图标和字体
使用图标变得非常简单。您只需使用“i”元素并参考相应的类即可。可以通过更改字号来控制图标的大小。例如，下面的代码展示了如何制作使用 themePrimary (#0078d7) 颜色的超大表图标。 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

若要查找 Office UI Fabric 中可用的更多图标，请在“[图标](https://dev.office.com/fabric#/styles/icons)”页上使用搜索功能。找到要在外接程序中使用的图标后，请务必在图标名称前加上前缀 `ms-Icon--`。 

若要了解 Office UI Fabric 中可用的字号和颜色，请参阅[版式](https://dev.office.com/fabric#/styles/typography)和[颜色](https://dev.office.com/fabric#/styles/colors)。

##<a name="3-use-fabric-js-ux-components"></a>3.使用 Fabric JS 用户体验组件

Fabric 提供了多个可在外接程序中使用的用户体验组件，如按钮或复选框。下面列出了我们建议用于外接程序的 Fabric JS 用户体验组件。若要在外接程序中使用其中一个 Fabric 组件，请单击相应的 Fabric 文档链接，然后按**使用此组件**中的说明操作。

> **注意：**随着时间的推移，我们将逐渐添加其他组件。 

- [痕迹导航](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Breadcrumb.md)
- [按钮](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Button.md)
- [复选框](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/CheckBox.md)
- [ChoiceFieldGroup](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ChoiceFieldGroup.md)
- [日期选取器](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/DatePicker.md)（有关如何在外接程序中实现日期选取器的示例，请参阅 [Excel 销售额跟踪程序](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)代码示例。）
- [下拉列表](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Dropdown.md)
- [标签](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Label.md)
- [链接](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Link.md)
- [列表](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/List.md)（请考虑在 CSS 中更改组件的默认样式。）
- [MessageBanner](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBanner.md)
- [MessageBar](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/MessageBar.md)
- [覆盖](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Overlay.md)
- [面板](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Panel.md)
- [透视](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Pivot.md)
- [ProgressIndicator](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/ProgressIndicator.md)
- [搜索框](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/SearchBox.md)
- [缓冲图标](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Spinner.md)
- [表](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Table.md)
- [TextField](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/TextField.md)
- [开关](https://github.com/OfficeDev/office-ui-fabric-js/blob/master/ghdocs/components/Toggle.md)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>将外接程序更新为使用 Fabric JS
如果你一直使用的是旧版 Office UI Fabric，并且想迁移到 Fabric JS，请务必了解新组件，并在外接程序中合并和测试新组件。请注意以下几点，它们有助于你进行更新规划：

- 使用 Fabric JS 时，组件初始化更加简单。对于旧版 Fabric，需要先在外接程序项目中添加 Fabric 组件的 JavaScript 文件（包括对该文件的 `<Script>` 引用），然后初始化组件。在 Fabric JS 中，不再需要添加 Fabric 组件的 JavaScript 文件及关联的 `<Script>` 引用。只需初始化 Fabric 组件即可。   
- 多个组件现在提供可控制用户体验组件行为的函数。例如，复选框控件具有 `toggle` 函数，可以在选中和取消选中状态之间进行切换。 
- 更新了某些图标类名和样式。
- 最明显的变化是在多个组件中使用 `<label>` 元素。`<label>` 元素控制组件样式。可能需要更新用户体验代码，才能使用 `<label>` 元素。例如，更改 Fabric JS 复选框上 `<input>` 元素的 checked 属性值对复选框不会产生任何影响。请改用 `check`、`unCheck` 或 `toggle` 函数。   

##<a name="next-steps"></a>后续步骤
若要获得端到端代码示例以了解如何使用 Fabric JS，我们已经为你准备好了。请参阅以下资源：

- [Excel 销售额跟踪程序](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

##<a name="related-resources"></a>相关资源
若要获得旧版 Fabric 的代码示例或文档，请参阅以下资源：

- [用户体验设计模式（使用 Fabric 2.6.1）](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Office 外接程序 Fabric UI 示例（使用 Fabric 1.0）](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [在 Office 外接程序中使用 Fabric 2.6.1](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

