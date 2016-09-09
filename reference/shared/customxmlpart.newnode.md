
# NodeInsertedEventArgs.newNode 属性
获取刚刚添加到  **CustomXMLPart** 对象的节点。

|||
|:-----|:-----|
|**主机：**|Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```
var myNode = eventArgsObj.newNode;
```


## 返回值

表示刚刚插入的节点的 [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md)。


## 备注

请注意，如果刚刚将子树添加到文档中，则此节点可能有子级。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

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
