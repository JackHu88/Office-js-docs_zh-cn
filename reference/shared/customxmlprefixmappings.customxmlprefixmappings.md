
# <a name="customxmlprefixmappings-object"></a>CustomXmlPrefixMappings 对象
表示自定义命名空间前缀映射的集合。

|||
|:-----|:-----|
|**主机：**|Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```
CustomXmlPrefixMappings
```


## <a name="members"></a>成员


**方法**


|**名称**|**说明**|
|:-----|:-----|
|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|将前缀异步添加到命名空间映射，以便查询某个项时使用。|
|[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|异步获取映射到指定前缀的命名空间。|
|[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|异步获取指定命名空间的前缀。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|CustomXmlParts|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Word 的支持。|
|1.0|引入|
