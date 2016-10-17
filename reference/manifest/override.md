
# <a name="override-element"></a>Override 元素
提供一种为其他区域设置指定某设置的值的方法。

 **外接程序类型：**内容、任务窗格、邮件


## <a name="syntax:"></a>语法：


```XML
<Override Locale="string " Value="string " />
```


## <a name="contained-in:"></a>包含在：


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

## <a name="attributes"></a>属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|区域设置|字符串|必需|为此替代项指定区域设置的区域性名称，采用 BCP 47 语言标记格式，例如 `"en-US"`。|
|值|字符串|必需|指定表示为指定区域设置的设置的值。|

## <a name="additional-resources"></a>其他资源



- [Office 外接程序的本地化](../../docs/develop/localization.md#off15wecon_LocalesManifest)
    
