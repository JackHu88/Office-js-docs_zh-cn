
# Override 元素
提供一种为其他区域设置指定某设置的值的方法。

 **外接程序类型：**内容、任务窗格、邮件


## 语法：


```XML
<Override Locale="string " Value="string " />
```


## 包含在：


||
|:-----|
|[CitationText](../../reference/manifest/citationtext.md)|
|[说明](../../reference/manifest/description.md)|
|[DictionaryName](../../reference/manifest/dictionaryname.md)|
|[DictionaryHomePage](../../reference/manifest/dictionaryhomepage.md)|
|[DisplayName](../../reference/manifest/displayname.md)|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|
|[IconUrl](../../reference/manifest/iconurl.md)|
|[QueryUri](../../reference/manifest/queryuri.md)|
|[SourceLocation](../../reference/manifest/sourcelocation.md)|
|[SupportUrl](../../reference/manifest/supporturl.md)|

## 属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|区域设置|string|必需|为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。|
|值|string|必需|指定表示为指定区域设置的设置的值。|

## 其他资源



- [Office 外接程序的本地化](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
