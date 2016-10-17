
# <a name="asyncresult-object"></a>AsyncResult 对象
用于封装异步请求的结果的对象，包括状态和错误信息（如果请求失败）。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```
AsyncResult
```


## <a name="members"></a>成员


**属性**


|**名称**|**Description**|
|:-----|:-----|
|**[asyncContext](../../reference/shared/asyncresult.asynccontext.md)**|获取与传入时状态相同的传递给调用方法的可选 _asyncContext_ 参数的用户定义项。|
|**[error](../../reference/shared/asyncresult.error.md)**|如果出现任何错误，获取提供错误描述的 **Error** 对象。|
|**[status](../../reference/shared/asyncresult.status.md)**|获取异步操作的状态。|
|**[value](../../reference/shared/asyncresult.value.md)**|获取此异步操作的负载或内容（如有）。|

## <a name="remarks"></a>备注

当执行您传递给 "Async" 方法的 _callback_ 参数的函数时，它会接收您可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

以下示例适用于内容和任务窗格外接程序。此示例显示了对  [Document](../../reference/shared/document.getselecteddataasync.md) 对象的**getSelectedDataAsync** 方法的调用。




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

作为  _callback_ 实参 ( `function (result){...}`) 传递的匿名函数具有一个名为  _result_ 的形参，该形参在该函数执行时提供对 **AsyncResult** 对象的访问。完成对 **getSelectedDataAsync** 方法的调用时，系统会执行回调函数，并且以下代码行会访问 **AsyncResult** 对象的 **value** 属性以返回文档中选择的数据：

 `var dataValue = result.value;`

请注意，函数中的其他代码行使用回调函数的  _result_ 参数访问 **AsyncResult** 对象的 **status** 和 **error** 属性。

The **AsyncResult** object is available from the function passed as the argument to the _callback_ parameter of the following methods:



|**Parent Object**|**Method**|
|:-----|:-----|
|**Document**（仅限 Excel、PowerPoint、Project 和 Word）|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|
||[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|
|**Bindings**（仅限 Excel 和 Word）|[addFromPromptAsync](../../reference/shared/bindings.addfrompromptasync.md)|
||[addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md)|
||[getAllAsync](../../reference/shared/bindings.getallasync.md)|
||[getByIdAsync](../../reference/shared/bindings.getbyidasync.md)|
||[releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md)|
|**Binding**（仅限 Excel 和 Word）|[getDataAsync](../../reference/shared/binding.getdataasync.md)|
||[setDataAsync](../../reference/shared/binding.setdataasync.md)|
||[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|
|**TableBinding**（仅限 Excel 和 Word）||
||[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|
||[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|
|**Settings**（仅限 Excel、PowerPoint 和 Word）|[refreshAsync](../../reference/shared/settings.refreshasync.md)|
||[saveAsync](../../reference/shared/settings.saveasync.md)|
|**CustomXmlNode**（仅限 Word）|[getNodesAsync](../../reference/shared/customxmlnode.getnodesasync.md)|
||[getNodeValueAsync](../../reference/shared/customxmlnode.getnodevalueasync.md)|
||[getXmlAsync](../../reference/shared/customxmlnode.getxmlasync.md)|
||[setNodeValueAsync](../../reference/shared/customxmlnode.setnodevalueasync.md)|
||[setXmlAsync](../../reference/shared/customxmlnode.setxmlasync.md)|
|**CustomXmlPart**（仅限 Word）|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|
||[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|
||[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|
|**CustomXmlParts**（仅限 Word）|[addAsync](../../reference/shared/customxmlparts.addasync.md)|
||[getByIdAsync](../../reference/shared/customxmlparts.getbyidasync.md)|
||[getByNamespaceAsync](../../reference/shared/customxmlparts.getbynamespaceasync.md)|
|**CustomXmlPrefixMappings**（仅限 Word）|[addNamespaceAsync](../../reference/shared/customxmlprefixmappings.addnamespaceasync.md)|
||[getNamespaceAsync](../../reference/shared/customxmlprefixmappings.getnamespaceasync.md)|
||[getPrefixAsync](../../reference/shared/customxmlprefixmappings.getprefixasync.md)|
|**Mailbox**（仅限 Outlook）|
  [getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||
  [makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties**（仅限 Outlook）|
  [saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item**（仅限 Outlook）|
  [loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings**（仅限 Outlook）|
  [saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



| |**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**适用于 Mac 的 Outlook**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**项目**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
