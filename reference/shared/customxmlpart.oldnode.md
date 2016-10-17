
# <a name="nodedeletedeventargs.oldnode-property"></a>NodeDeletedEventArgs.oldNode 属性
获取刚从 **CustomXmlPart** 对象删除的节点。

|||
|:-----|:-----|
|**主机：**|Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```
var myNode = eventArgsObj.oldNode;
```


## <a name="return-value"></a>返回值

表示刚删除的节点的 [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md)。


## <a name="remarks"></a>备注

请注意，如果正在从文档中删除子树，此节点可能有子级。此节点还将是一个"断开"节点，以便您能从该节点向下查询，但是不能向上查询树 - 该节点似乎是单独存在的。


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




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Word 的支持。|
|1.0|引入|
