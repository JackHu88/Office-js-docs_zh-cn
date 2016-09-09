
# Sets 元素
指定适用于 Office 的 JavaScript API 的最小子集，Office 外接程序需要该子集才能激活。

 **外接程序类型：**内容、任务窗格、邮件


## 语法：


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## 包含在：

[Requirements](../../reference/manifest/requirements.md)


## 可以包含：

[Set](../../reference/manifest/set.md)


## 属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|可选|为所有子 **Set** 元素指定默认的 [MinVersion](../../reference/manifest/set.md) 属性值。默认值为“1.1”。|

## 注解

有关要求集的详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

有关 **Set** 元素的 **MinVersion** 属性和 **Sets** 元素的 **DefaultMinVersion** 属性的详细信息，请参阅[在清单中设置 Requirements 元素](../../docs/overview/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。

