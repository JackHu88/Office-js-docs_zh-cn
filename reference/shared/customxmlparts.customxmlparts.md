
# <a name="customxmlparts-object"></a>CustomXmlParts 对象
表示 [CustomXMLPart](../../reference/shared/customxmlpart.customxmlpart.md) 对象的集合。

|||
|:-----|:-----|
|**主机：**|Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```
Office.context.document.customXmlParts
```


## <a name="members"></a>成员


**方法**


|**名称**|**说明**|
|:-----|:-----|
|[addAsync](../../reference/shared/customxmlparts.addasync.md)|将新的自定义 XML 部件异步添加到文件中。|
|[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|按其 ID 异步获取自定义 XML 部件。|
|[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|异步获取匹配指定的命名空间的自定义 XML 部件的数组。|

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
