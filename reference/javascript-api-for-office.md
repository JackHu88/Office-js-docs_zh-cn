
# <a name="javascript-api-for-office-reference"></a>适用于 Office 的 JavaScript API 参考

借助适用于 Office 的 JavaScript API，您可以创建可与 Office 主机应用程序中的对象模型进行交互的 Web 应用程序。您的应用程序将引用 office.js 库中，该库是一个脚本加载程序。Office.js 库加载适用于正在运行外接程序的 Office 应用程序的对象模型。您可以使用以下 JavaScript 对象模型：


1. 公用（必需） - 与 Office 2013 一起引入的 API。这为**所有 Office 主机应用程序**加载 API，并将您的外接程序与 Office 客户端应用程序连接。对象模型包含特定于 Office 客户端的 API 和适用于多个 Office 客户端主机应用程序的 API。在[共享](../reference/shared/shared-api.md)和 **outlook** 下的所有内容都被视为公用 API。**Microsoft.Office.WebExtension** 命名空间（默认状态下使用代码中的别名 [Office](../reference/shared/office.md) 引用该命名空间）包含可以用于将与内容交互的脚本写入 Office 文档、工作簿、演示文稿、邮件项以及 Office 外接程序中的项目的对象。如果外接程序将 Office 2013 及更高版本作为目标，则必须使用这些公用 API。此对象模型使用回调。

1. 特定于主机 - 与 **Office 2016** 一起引入的 API。此对象模型提供特定于主机的强类型对象，这些对象对应于使用 Office 客户端时所看到的熟悉对象，并表示 Office JavaScript API 的未来。特定于主机的 API 目前包括 [Word JavaScript API](../reference/word/word-add-ins-reference-overview.md) 和 [Excel JavaScript API](../reference/excel/application.md)。此对象模型使用承诺模式。

从 TOC 上方的下拉列表选择 Office 客户端，以便根据您的目标主机应用程序筛选内容。

## <a name="supported-host-applications"></a>支持的主机应用程序
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

了解有关[支持的主机和其他要求](../docs/overview/requirements-for-running-office-add-ins.md)的详细信息。

## <a name="open-api-specifications"></a>开放 API 规范

在我们设计和开发新的 API 以用于 Office 外接程序时，我们将使它们适用于[开放 API 规范](openspec.md)页的反馈。了解管道中的新增功能，并提供您对我们的设计规范的宝贵意见。

