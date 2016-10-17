
# <a name="requirements-element"></a>Requirements 元素
指定适用于 Office 的 JavaScript API 要求（[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_sets)和/或方法）的最小集，Office 外接程序需要该集才能激活。

 **外接程序类型：**内容、任务窗格、邮件


## <a name="syntax:"></a>语法：


```XML
<Requirements>
   ...
</Requirements>
```


## <a name="contained-in:"></a>包含在：

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>可以包含：



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](../../reference/manifest/sets.md)|x|x|x|
|[Methods](../../reference/manifest/methods.md)|x||x|

## <a name="remarks"></a>备注

有关要求集的详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

