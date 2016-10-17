
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-(cdn)"></a>从 适用于 Office 的 JavaScript API 的内容传送网络 (CDN) 引用 适用于 Office 的 JavaScript API 库


[适用于 Office 的 JavaScript](../../reference/javascript-api-for-office.md) 库包含 Office.js 文件和关联的特定于主机应用程序的 .js 文件，例如 Excel-15.js 和 Outlook-15.js。 


引用该 API 最简单的方法是通过添加以下 `<script>` 到你的页面的 `<head>` 标记来使用我们的 CDN：  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

CDN URL 中 `/1/` 前面的 `office.js` 指定使用 Office.js 版本 1 中的最新增量版本。因为适用于 Office 的 JavaScript API 保持向后兼容性，所以最新版本将继续支持版本 1 中之前引入的 API 成员。如果你需要升级现有的项目，请参阅 [更新适用于 Office 的 JavaScript API 的版本和清单架构文件] (../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)。 

如果计划从 Office 应用商店发布你的 Office 外接程序，则必须使用 CDN 引用。本地引用仅适用于内部、开发和调试方案。

> **重要说明：**开发任何 Office 主机应用程序的外接程序时，从页面的 `<head>` 部分内部引用适用于 Office 的 JavaScript 十分重要。这将确保在任何正文元素之前，API 已经完全初始化。Office 主机要求外接程序在激活 5 秒钟内进行初始化。超过此阈值会导致声明外接程序无响应，并且会向用户显示错误消息。       

## <a name="additional-resources"></a>其他资源



- [了解适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Office 外接程序平台概述](../../docs/overview/office-add-ins.md)
    
- [Office 外接程序开发生命周期](../../docs/design/add-in-development-lifecycle.md)
    
- [适用于 Office 的 JavaScript API](../../reference/javascript-api-for-office.md)
    
