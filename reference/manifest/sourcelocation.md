
# SourceLocation 元素
指定 Office 外接程序的源文件位置为介于 1 和 2018 个字符之间的 URL。 源位置必须是 HTTPS 地址，而非文件路径。

 **外接程序类型：**内容、任务窗格、邮件


## 语法：


```XML
<SourceLocation DefaultValue="string " />
```


## 包含在：

[DefaultSettings](../../reference/manifest/defaultsettings.md)（内容和任务窗格外接程序）

[FormSettings](../../reference/manifest/formsettings.md)（邮件外接程序）


## 可以包含：

[Override](../../reference/manifest/override.md)


## 属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|必需|指定该设置的默认值，表示为 [DefaultLocale](../../reference/manifest/defaultlocale.md) 元素中指定的区域设置。|
