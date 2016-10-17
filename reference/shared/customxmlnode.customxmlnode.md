
# <a name="customxmlnode-object"></a>CustomXmlNode 对象
表示文档中的树中的 XML 节点。

|||
|:-----|:-----|
|**主机：**|Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```js
CustomXmlNode
```


## <a name="members"></a>成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|[baseName](../../reference/shared/customxmlnode.basename.md)|获取不带命名空间前缀的节点的基名称（如有）。|
|[nodeType](../../reference/shared/customxmlnode.nodetype.md)|获取 **CustomXMLNode** 的类型。|
|[namespaceUri](../../reference/shared/customxmlnode.namespaceuri.md)|检索 **CustomXMLPart** 的字符串 GUID。|

**方法**


|**名称**|**说明**|
|:-----|:-----|
|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|作为匹配相对 XPath 表达式的 **CustomXMLNode** 对象的数组异步获取节点。|
|[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|异步获取节点的值。|
|[getTextAsync](customxmlnode.gettextasync.md)|异步获取自定义 XML 部件中 XML 节点的文本。|
|[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|异步获取节点的 XML。|
|[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|异步设置节点的值。|
|[setTextAsync](customxmlnode.settextasync.md)|异步设置自定义 XML 部件中 XML 节点的文本。|
|[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|异步设置节点的 XML。|

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
