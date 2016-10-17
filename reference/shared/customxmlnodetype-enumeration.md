
# <a name="customxmlnodetype-enumeration"></a>CustomXMLNodeType 枚举
指定节点类型。



|||
|:-----|:-----|
|**主机：**|Word|
|**包含最后一次更改的版本**|1.1|



```js
Office.CustomXMLNodeType
```


## <a name="members"></a>成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.CustomXMLNodeType.Attribute|"attribute"|节点为属性。|
|Office.CustomXMLNodeType.CData|"CData"|节点为 CData 类型。|
|Office.CustomXMLNodeType.NodeComment|"comment"|节点为注释。|
|Office.CustomXMLNodeType.Element|"element"|节点为元素。|
|Office.CustomXMLNodeType.NodeDocument|"nodeDocument"|节点为 Document 元素。|
|Office.CustomXMLNodeType.ProcessingInstruction|"processingInstruction"|节点为处理指令。|
|Office.CustomXMLNodeType.Text|"text"|节点为文本节点。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y|Y|Y|



|||
|:-----|:-----|
|**外接程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Word 的支持。|
|1.0|引入|
