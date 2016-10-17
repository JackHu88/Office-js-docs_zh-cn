
# <a name="bindings-object"></a>Bindings 对象
表示外接程序在文档中所具有的绑定。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|在其中进行的**最后更改**|1.1|

```js
Office.context.document.bindings
```


**属性**

|||
|:-----|:-----|
|名称|说明|
|[document](../../reference/shared/bindings.document.md)|获取表示与此组绑定关联的文档的 **Document** 对象。|

**方法**

|||
|:-----|:-----|
|名称|说明|
|[addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)|将绑定添加到文档中的命名项。|
|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|显示可让用户指定要绑定的选择的 UI。|
|[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|将指定类型的绑定的绑定对象添加到文档中的当前选择。|
|[getAllAsync](../../reference/shared/bindings.getallasync.md)|获取先前创建的所有绑定。|
|[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|按其标识符获取指定的绑定。|
|[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|移除指定绑定。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


|||||
|:-----|:-----|:-----|:-----|
||Office for Windows Desktop|Office Online（在浏览器中）|Office for iPad|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel和 Word 的支持|
|1.1|针对 [addFromNamedItemAsync](../../reference/shared/bindings.addfromnameditemasync.md)、[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md) 和[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) 添加了以 Excel 相关应用程序中表绑定的形式绑定到矩阵数据的支持。|
|1.1|<ul><li>针对 <a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">Document</a> 属性，添加了对 <span class="keyword">Document</span> 对象的访问，该对象表示 Access 相关内容外接程序中的当前 Access 数据库。</li><li>针对所有方法，添加了对 Access 相关内容外接程序中表绑定的支持。 </li></ul>|
|1.0|引入|
