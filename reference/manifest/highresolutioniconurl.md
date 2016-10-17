
# <a name="highresolutioniconurl-element"></a>HighResolutionIconUrl 元素
指定用于表示插入 UX 中的 Office 外接程序和高 DPI 屏幕上的 Office 应用商店的图像的 URL。

 **外接程序类型：**内容、任务窗格、邮件


## <a name="syntax:"></a>语法：


```XML
<HighResolutionIconUrl DefaultValue="string " />
```


## <a name="can-contain:"></a>可以包含：

[Override](../../reference/manifest/override.md)


## <a name="attributes"></a>属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|字符串 (URL)|必需|指定该设置的默认值，表示为 [DefaultLocale](../../reference/manifest/defaultlocale.md) 元素中指定的区域设置。|

## <a name="remarks"></a>注解

对于邮件外接程序，图标显示在“**文件**” > “**管理外接程序**”UI 中。对于内容或任务窗格外接程序，图标显示在“**插入**” > “**外接程序**”UI 中。

图像必须以 64 x 64 像素的分辨率采用下列任一文件格式进行保存：GIF、JPG、PNG、EXIF、BMP 或 TIFF。有关详细信息，请参阅 [创建有效的 Office 应用商店应用和外接程序](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)中的_为你的应用创建一致的视觉标识_。

