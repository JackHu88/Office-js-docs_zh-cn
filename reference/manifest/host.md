
# “主机”元素
指定 Office 外接程序支持的 Office 主机应用程序的类型。

 **外接程序类型：**内容、任务窗格、邮件


## 语法：


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## 属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|Name|string|必需|Office 主机应用程序的类型的名称。|

## 备注

可以指定 **Host** 元素中的 **Name** 属性中的以下值。每个值映射到外接程序支持的 Office 主机应用程序的一个或多个集。



|**Name**|**Office 主机应用程序**|
|:-----|:-----|
| `"Document"`|Word、Word Online 和 iPad 上的 Word|
| `"Database"`|Access Web 应用程序|
| `"Mailbox"`|Outlook、Outlook Web App 和适用于设备的 OWA|
| `"Notebook"`|OneNote Online|
| `"Presentation"`|PowerPoint、PowerPoint Online 和 iPad 上的 PowerPoint|
| `"Project"`|Project|
| `"Workbook"`|Excel、Excel Online、iPad 上的 Excel|

## 备注

有关指定主机支持的详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

