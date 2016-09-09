
# IconUrl 元素
指定用于表示插入 UX 和 Office 应用商店中的 Office 外接程序的图像的 URL。

 **外接程序类型：**内容、任务窗格、邮件


## 语法：


```XML
<IconUrl DefaultValue="string " />
```


## 可以包含：

[Override](../../reference/manifest/override.md)


## 属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|string|必需|指定该设置的默认值，表示为 [DefaultLocale](../../reference/manifest/defaultlocale.md) 元素中指定的区域设置。|

## 注解

对于邮件外接程序，该图标显示在“**文件**” > “**管理外接程序**”UI (Outlook) 或“**设置**” > “**管理外接程序**”UI (Outlook Web App) 中。 对于内容或任务窗格外接程序，图标显示在“**插入**” > “**外接程序**”UI 中。 对于所有外接程序类型，如果你将外接程序发布到 Office 应用商店，则该图标也将用于 Office 应用商店网站上。

该图像必须为以下文件格式之一：GIF、JPG、PNG、EXIF、BMP 或 TIFF。对于内容和任务窗格应用程序，必须将该图像指定为 32 x 32 像素。对于邮件应用程序，该图像必须是 64 x 64 像素。您还应使用 [HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md) 元素对运行在高 DPI 屏幕上的 Office 主机应用程序指定搭配使用的图标。有关详细信息，请参阅_创建有效的 Office 应用商店应用和外接程序_中的[为应用程序创建一致的视觉标识](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)一节。

