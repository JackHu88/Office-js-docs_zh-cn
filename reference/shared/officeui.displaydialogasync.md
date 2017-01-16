# <a name="uidisplaydialogasync-method"></a>UI.displayDialogAsync 方法

在 Office 主机中显示一个对话框。 

## <a name="requirements"></a>要求

|主机|引入版本|包含最后一次更改的版本|
|:---------------|:--------|:----------|
|Word、Excel、PowerPoint|1.1|1.1|
|Outlook|Mailbox 1.4|邮箱 1.4|

此方法在 Word、Excel 或 PowerPoint 外接程序的 DialogAPI [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)以及 Outlook 的邮箱要求集 1.4 中引入。若要指定 DialogAPI 要求集，请在清单中运行以下代码。

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.1"> 
    <Set Name="DialogAPI"/> 
  </Sets> 
</Requirements> 
```

若要指定邮箱 1.4 要求集，请在清单中运行以下代码。

```xml
<Requirements> 
  <Sets DefaultMinVersion="1.4"> 
    <Set Name="Mailbox"/> 
  </Sets> 
</Requirements> 
```

若要在运行时在 Word、Excel 或 PowerPoint 外接程序中检测此 API，请运行以下代码。

```js
if (Office.context.requirements.isSetSupported('DialogAPI', 1.1)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

若要在运行时在 Outlook 外接程序中检测此 API，请运行以下代码。

```js
if (Office.context.requirements.isSetSupported('Mailbox', 1.4)) {  
  // Use Office UI methods; 
} else { 
  // Alternate path 
} 
```

或者，可以先检查 `displayDialogAsync` 方法是否未定义，然后再使用它。

```js
if (Office.context.ui.displayDialogAsync !== undefined) {
  // Use Office UI methods
}
```

### <a name="supported-platforms"></a>支持的平台
有关支持的平台的信息，请参阅[对话框 API 要求集](../requirement-sets/dialog-api-requirement-sets.md)。

## <a name="syntax"></a>语法

```js
Office.context.ui.displayDialogAsync(startAddress, options, callback);
```
##<a name="examples"></a>示例

有关使用 **displayDialogAsync** 方法的简单示例，请参阅 GitHub 上的 [Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example/)。

有关显示身份验证应用场景的示例，请参阅：

- [Microsoft Graph ASP.Net 插入图表中的 PowerPoint 外接程序](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Office 外接程序 Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Excel 外接程序 ASP.NET QuickBooks ](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Office 外接程序 ASP.net MVC 服务器身份验证示例](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)
- [Office 外接程序 Office 365 客户端 AngularJS 身份验证](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)


 
## <a name="parameters"></a>参数

| 参数    | 类型   |说明|
|:---------------|:--------|:----------|
|startAddress|字符串|接受在对话框中打开的初始 HTTPS(TLS) URL。 <ul><li>初始网页必须与父页位于相同的域。初始网页加载后，你可以转到其他域。</li><li>调用 [office.context.ui.messageParent](officeui.messageparent.md) 的所有页也必须都与父页位于相同的域。</li></ul>|
|选项|object|可选。接受用于定义对话框行为的 options 对象。|
|callback|对象|接受用于处理对话框创建尝试的 callback 方法。|
    
### <a name="configuration-options"></a>配置选项
以下配置选项适用于对话框。


| 属性     | 类型   |说明|
|:---------------|:--------|:----------|
|**width**|object|可选。以占当前显示器的百分比的形式，定义对话框的宽度。默认值为 80%。最小分辨率为 250 像素。|
|**height**|object|可选。以占当前显示器的百分比的形式，定义对话框的高度。默认值为 80%。最小分辨率为 150 像素。|
|**displayInIframe**|对象|可选。确定是否应在 IFrame 内显示对话框。**此设置仅适用于 Office Online 客户端**，桌面客户端可忽略此设置。可取值如下：<ul><li>False（默认值）- 对话框将显示为一个新的浏览器窗口（弹出窗口）。对于无法在 IFrame 中显示的身份验证页建议使用此值。 </li><li>True - 对话框将显示为使用 IFrame 的浮动重叠窗口。对于用户体验和性能而言，这是最佳选择。</li>|


## <a name="callback-value"></a>回调值
在你传递给 _callback_ 参数的函数执行后，它会收到你可以从回调函数的唯一参数访问的 [AsyncResult](../../reference/shared/asyncresult.md) 对象。

在传递给 **displayDialogAsync** 方法的回调函数中，可以使用 **AsyncResult** 对象的属性返回以下信息。



|**属性**|**用于**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|访问 [Dialog](../../reference/shared/officeui.dialog.md) 对象。|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|确定操作是成功还是失败。|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|如果操作失败，则访问提供错误信息的 [Error](../../reference/shared/error.md) 对象。|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|如果将用户定义的对象或值作为 _asyncContext_ 参数传递，则对其进行访问。|

### <a name="errors-from-displaydialogasync"></a>DisplayDialogAsync 错误

除常规的平台和系统错误外，以下是调用 **displayDialogAsync** 时出现的特定错误。

|**代码编号**|**含义**|
|:-----|:-----|
|12004|传递给 `displayDialogAsync` 的 URL 域不受信任。该域必须与主机页（包括协议和端口号）具有同一域，或必须在外接程序清单的 `<AppDomains>` 部分中注册。|
|12005|传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。需要使用 HTTPS。（在 Office 的某些版本中，返回 12005 的错误消息与返回 12004 错误消息是相同的。）|
|12007|从任务窗格已经打开了一个对话框。任务窗格外接程序一次只能打开一个对话框。|



## <a name="design-considerations"></a>设计注意事项
下列设计注意事项适用于对话框：

- Office 外接程序随时都可能有一个打开的对话框。
- 用户可以移动每个对话框和调整其大小。
- 每个对话框在打开时都在屏幕上居中显示。
- 对话框按照创建的顺序出现在主机应用程序顶部。

使用对话框可以执行以下操作：

- 显示身份验证页以收集用户凭据。
- 显示来自 ShowTaspane 或 ExecuteAction 命令的错误/进度/输入屏幕。
- 临时增加用户可用于完成一项任务的表面区域。

不要使用对话框与文档进行交互。而是使用任务窗格。 

有关可以用于创建对话框的设计模式，请参阅 GitHub 的 Office 外接程序 UX 设计模式存储库中的 [客户端对话框](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/Patterns/Client_Dialog.md)。
