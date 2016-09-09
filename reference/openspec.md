# 开放性 API 规范

感谢你对我们正在设计的新 API 和功能的关注。我们将在此处提供早期版本的 API 规范，以便收集社区反馈。你的意见有助于确保最终设计满足对你而言十分重要的用例的要求。 

此处所介绍的功能可能处于不同的开发阶段，如早期设计或公开预览。在功能全面推出后，我们会将相关内容从此页面中删除，并会更新我们的文档，以添加新功能的详细信息。 

_**重要提示：**此处列出的功能仍处于设计和审阅阶段，尚未全面推出。这些功能和 API 可能会发生变更。_

## 新 Excel JavaScript API
加入我们，共同审阅我们对新 Excel JavaScript API 的设计。 新更新的 API 包括 customXML 部件、数据透视表刷新、已筛选范围的视图、作为图像的范围和表、向表追加多行等。 

**查看 [Excel JavaScript 1.3 API 页](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_1.3_OpenSpec) 了解详细信息，并提供你的反馈。**

## 新 Word JavaScript API
Word JavaScript API 1.3 更新包含自引入此 API 以来我们实现的最多一组更改。借助新 API，你可以： 

* 创建和更改内存中的文档
* 创建和访问列表对象
* 创建和访问表对象
* 通过更多方法访问和比较范围对象

几乎所有 Word JavaScript API 对象中都已经实现了这些更改。对于 Windows 和 Mac 的桌面版 Word 2016 以及 iPad 上的 Word 2016，此功能现在或很快就会进行预览阶段。请将你的客户端更新为最新的每月内部版本，并开始实现这些强大功能吧！

**请参阅 [Word JS API 1.3](https://github.com/OfficeDev/office-js-docs/tree/WordJs_1.3_Openspec/word) 页面，了解详细信息并提供你的反馈。**

## 文档属性访问权限
我们一直在努力增加让 Web 外接程序能够访问（获取、设置）文档级属性的功能。借助此功能，外接程序可以将文档属性集成到自定义工作流中，也可以读取/设置文档属性。Word 和 Excel 将支持此功能。PowerPoint 可能会支持此功能。此功能还适用于 Excel REST API（Excel 支持 REST 服务）。我们将介绍基本设计理念，并通过用例和代码片段来演示 API 在添加后的工作原理。欢迎你提供设计方面的反馈。 

**请参阅[文档属性开放性规范](https://github.com/OfficeDev/office-js-docs/tree/DocumentProperties_OpenSpec)页面，了解详细信息并提供你的反馈。**

