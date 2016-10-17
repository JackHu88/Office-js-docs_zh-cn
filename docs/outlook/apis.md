
# <a name="outlook-add-in-apis"></a>Outlook 外接程序 API

要将 API 用于您的 Outlook 外接程序，您必须指定 Office.js 库的位置、要求集、架构和权限。

## <a name="office.js-library"></a>Office.js 库

要与 Outlook 外接程序 API 进行交互，你必须使用 Office.js 中的 JavaScript API。库的 CDN 为 _https://appsforoffice.microsoft.com/lib/1/hosted/Office.js_。提交到 Office 应用商店的外接程序必须按此 CDN 引用 Office.js，它们不能使用本地引用。 

在实施外接程序 UI 的网页（.html、.aspx 或 .php 文件）的 **head** 标记（即 **script** 标记的 **src** 属性）中声明 CDN：


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

添加 API 时，Office.js 的 URL 将保持不变。仅当我们打破现有的 API 行为时，才会更改 URL 中的版本。

> **重要说明：**开发任何 Office 主机应用程序的外接程序时，从页面的 `<head>` 区域内引用适用于 Office 的 JavaScript API。这将确保 API 已先于所有正文元素完全初始化。Office 主机要求外接程序在激活 5 秒钟内进行初始化。超过此阈值会导致声明的外接程序无响应，并且会向用户显示错误消息。  

## <a name="requirement-sets"></a>要求集

所有 Outlook API 均属于邮箱要求集。邮箱要求集具有不同版本，我们发布的每个新 API 集均属于较高版本的要求集。并非所有 Outlook 客户端都支持我们发布的最新 API 集，但如果某个 Outlook 客户端声明支持某个要求集，它将支持该要求集中的所有 API。 

若要控制外接程序在哪些 Outlook 客户端中显示，请在清单中指定最低要求集版本。例如，如果你指定要求集版本 1.3，则外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。 

指定要求集不会将外接程序限定于该版本中的 API。如果外接程序指定要求集 v1.1，却在支持 v1.3 的 Outlook 客户端中运行，该外接程序仍可以使用 v1.3 API。要求集仅控制外接程序在哪些 Outlook 客户端中显示。

要检查大于清单中所指定要求集的要求集中任何 API 的可用性，可以使用标准 JavaScrip：


```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> **注意：**对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

对于这种情况，你应该指定支持关键 API 集的最低要求集，否则外接程序的关键功能将无法运行。您可以在 **Requirements**、**Sets** 和 **Set** 元素中指定清单中的要求集。有关详细信息，请参阅 [Outlook 外接程序清单](../outlook/manifests/manifests.md)和[了解 Outlook API 要求集](..\..\reference\outlook\tutorial-api-requirement-sets.md)。

**Methods** 元素不适用于 Outlook 外接程序，因此，你无法声明对特定方法的支持。


## <a name="permissions"></a>权限

外接程序需要相应的权限才能使用所需的 API。有四个级别的权限。有关详细信息，请参阅[了解 Outlook 外接程序权限](../outlook/understanding-outlook-add-in-permissions.md)。


|**权限级别**|**说明**|
|:-----|:-----|
|受限|允许使用实体，但不允许使用正则表达式。|
|读取项目|除了 _Restricted_ 所允许的权限，它还允许：<ul><li>正则表达式</li><li>Outlook 外接程序 API 读取访问</li><li>获取项目属性和回调令牌</li></ul>|
|读/写|除了 _Read item_ 所允许的权限，它还允许：<ul><li>Outlook 外接程序 API 完全访问权限，但不包括 <b>makeEwsRequestAsync</b></li><li>设置项目属性</li></ul>|
|读/写邮箱|除了 _Read/write_ 所允许的权限，它还允许：<ul><li>创建、读取、写入项目和文件夹</li><li>发送项目</li><li>调用 [makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md#makeewsrequestasyncdata-callback-usercontext)</li></ul>|
总之，你应该指定外接程序所需的最低权限。权限在清单的 **Permissions** 元素中声明。有关详细信息，请参阅 [Outlook 外接程序清单](../outlook/manifests/manifests.md)。有关安全问题的信息，请参阅 [Outlook 外接程序的隐私、权限和安全性](../outlook/../../docs/develop/privacy-and-security.md)。


## <a name="additional-resources"></a>其他资源

- [Outlook 外接程序清单](../outlook/manifests/manifests.md)

- [了解 Outlook API 要求集](../../reference/outlook/tutorial-api-requirement-sets.md)
    
- [Outlook 外接程序的隐私、权限和安全性](../outlook/../../docs/develop/privacy-and-security.md)
    
