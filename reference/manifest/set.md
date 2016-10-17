
# <a name="set-element"></a>Set 元素
指定来自适用于 Office 的 JavaScript API 的要求集，Office 外接程序需要该集才能激活。

 **外接程序类型：**内容、任务窗格、邮件


## <a name="syntax:"></a>语法：


```XML
<Set Name="string " MinVersion="n .n ">
```


## <a name="contained-in:"></a>包含在：

[Sets](../../reference/manifest/sets.md)


## <a name="attributes"></a>属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|名称|字符串|必需|[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)名称。|
|MinVersion|字符串|可选|指定您的外接程序所需的 API 集的最低版本。如果 **DefaultMinVersion** 的值已在父 [Sets](../../reference/manifest/sets.md) 元素中指定，则替代该值。|

## <a name="remarks"></a>注解

有关要求集的详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md#specify-office-hosts-and-api-requirements)。

有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。


 >**重要说明** 对于邮件外接程序，则只能使用一个 `"Mailbox"` 要求集。此要求集包含 Outlook 邮件外接程序支持的整个 API 子集，您必须在邮件外接程序清单中指定 `"Mailbox"` 要求集（针对内容和任务窗格外接程序，非可选）。另外，您无法在邮件外接程序中声明对特定方法的支持。

