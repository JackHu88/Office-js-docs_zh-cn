
# CustomXmlPart.namespaceManager 属性
获取对照当前 CustomXMLPart 使用的一组命名空间前缀映射 (CustomXMLPrefixMappings)。

|||
|:-----|:-----|
|**主机：**|Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|CustomXmlParts|
|**包含最后一次更改的版本**|1.1|

```js
var nsManager = customXmlPart.namespaceManager);
```


## 返回值

一个 [CustomXmlPrefixMappings](../../reference/shared/customxmlprefixmappings.customxmlprefixmappings.md) 对象。


## 示例




```js
function setXMLPartNamespaceManagerNamespace() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.namespaceManager.addNamespaceAsync("myPrefix", "myNamespace");
    });
}

```




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
