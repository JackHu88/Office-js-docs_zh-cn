
# CustomXmlPart 对象
表示 **CustomXMLParts** 集合中的单个 [CustomXMLPart](../../reference/shared/customxmlparts.customxmlparts.md)。

|||
|:-----|:-----|
|**主机：**|Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## 成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|获取指示是否已安装 CustomXMLPart 的值。|
|[id](../../reference/shared/customxmlpart.id.md)|获取 CustomXMLPart 的 GUID|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|获取对照当前 CustomXMLPart 使用的一组命名空间前缀映射 (CustomXMLPrefixMappings)。|

**方法**


|**名称**|**说明**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|为  **CustomXmlPart** 对象事件异步添加事件处理程序。|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|从集合中异步删除此自定义 XML 部件。|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|异步获取此自定义 XML 部件中与指定 XPath 匹配的任何 CustomXmlNodes。|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|异步获取此自定义 XML 部件内的 XML。|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|为  **CustomXmlPart** 对象事件移除事件处理程序。|

**事件**


|**名称**|**说明**|
|:-----|:-----|
|[dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md)|删除节点时发生。|
|[dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md)|插入节点时发生。|
|[dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md)|替换节点时发生。|

## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|CustomXmlParts|
|**最低权限级别**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Word 的支持。|
|1.0|引入|
