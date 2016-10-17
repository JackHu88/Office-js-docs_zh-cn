
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>获取和设置 Outlook 外接程序的外接程序元数据

您可以通过使用以下任一项管理 Outlook 外接程序中的自定义数据：

- 漫游设置，可管理用户邮箱的自定义数据。
    
- 自定义属性，可管理用户邮箱中某个项目的自定义数据。
    
两种方法均允许您访问仅可供 Outlook 外接程序访问的自定义数据，但两种方法分别存储数据。也就是说，自定义属性不能访问通过漫游设置存储的数据，反之亦然。数据存储在该邮箱的服务器上，并且在外接程序支持的所有外形因素上的后续 Outlook 会话中可访问。 

## <a name="custom-data-per-mailbox:-roaming-settings"></a>每个邮箱的自定义数据：漫游设置


您可以使用 [RoamingSettings](../../reference/outlook/RoamingSettings.md) 对象指定特定于用户的 Exchange 邮箱的数据，例如用户的个人数据和首选项。当您的邮件外接程序在设计在其上运行的任何设备（台式机、平板电脑或智能手机）上漫游时，可以访问漫游设置。

 对该数据的更改存储在当前 Outlook 会话的这些设置的内存副本中。您应该在更新后显式保存所有漫游设置，以便用户下次在同一设备或任何其他受支持设备上打开您的外接程序时可以使用这些设置。


### <a name="roaming-settings-format"></a>漫游设置格式


**RoamingSettings** 对象中的数据存储为序列化的 JavaScript 对象表示法 (JSON) 字符串。下面是一个结构示例，假定有三个名为 `add-in_setting_name_0`、`add-in_setting_name_1` 和 `add-in_setting_name_2` 的已定义漫游设置。


```js
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a>加载漫游设置


邮件外接程序通常在 [Office.initialize](../../reference/shared/office.initialize.md) 事件处理程序中加载漫游设置。以下 JavaScript 代码示例演示了如何加载现有漫游设置并获取两个设置的值，即“customerName”和“customerBalance”。


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>创建或分配漫游设置


继续前一个示例，下面的 JavaScript 函数  `setAddInSetting` 显示如何使用 [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) 方法用今天的日期设置名为 `cookie` 的设置，并使用 [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) 方法将所有漫游设置重新保存到服务器，以使数据持续存在。如果设置尚不存在， **set** 方法会创建设置，并将其分配给指定的值。 **saveAsync** 方法可异步保存漫游设置。此代码示例将回调方法 `saveMyAddInSettingsCallback` 传递给 **saveAsync**。当异步调用完成时，会使用一个参数  _asyncResult_ 调用 `saveMyAddInSettingsCallback`。此参数是一个 [AsyncResult](../../reference/outlook/simple-types.md) 对象，其中包含异步调用的结果和任何详细信息。可以使用可选的 _userContext_ 参数从异步调用向回调函数传递任何状态信息。


```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a>删除漫游设置


通过扩展前面的示例，以下 JavaScript 函数  `removeAddInSetting` 显示了如何使用 [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) 方法删除 `cookie` 设置并将所有漫游设置保存回 Exchange Server。


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox:-custom-properties"></a>邮箱中每个项目的自定义数据：自定义属性


您可以使用 [CustomProperties](../../reference/outlook/CustomProperties.md) 对象指定用户邮箱中某个项目的特定数据。例如，您的邮件外接程序可以对特定邮件进行分类，并使用自定义属性 `messageCategory` 标记类别。或者，如果您的邮件外接程序使用邮件中的会议建议创建约会，您可以使用自定义属性跟踪这些约会。这可以确保当用户再次打开邮件时，您的邮件外接程序不会再次创建约会。

与漫游设置类似，对自定义属性的更改将存储在当前 Outlook 会话的属性的内存副本中。为确保这些自定义属性在下次会话中可用，请将所有自定义属性保存到服务器。

这些外接程序和项目特定的自定义属性只能使用  **CustomProperties** 对象访问。这些属性不同于 Outlook 对象模型中基于 MAPI 的自定义属性 [UserProperties](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx)，也不同于 Exchange Web 服务 (EWS) 中的扩展属性。您无法使用 Outlook 对象模型或 EWS 访问  **CustomProperties**。

但是，邮件外接程序可以使用 EWS [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 操作获取基于 MAPI 的扩展属性。在服务器端可使用回调令牌访问 **GetItem**，在客户端则使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法进行访问。在 **GetItem** 请求中，在属性集中指定您需要的自定义扩展属性。邮件外接程序还可以使用 **makeEwsRequestAsync** 以及 EWS [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx) 和 [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) 操作来创建和修改扩展属性。




### <a name="using-custom-properties"></a>使用自定义属性


在可以使用自定义属性之前，您必须通过调用 [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法来加载这些属性。如果已为当前项设置了任何自定义属性，则此时会从 Exchanger 服务器中加载这些自定义属性。创建属性包之后，您可以使用 [set](../../reference/outlook/CustomProperties.md) 和 [get](../../reference/outlook/CustomProperties.md) 方法来添加和检索自定义属性。若要保存您对属性包所做的任何更改，则必须使用 [saveAsync](../../reference/outlook/CustomProperties.md) 方法以将更改保存到 Exchange 服务器中。


 >**注释**  由于 适用于 Mac 的 Outlook 不缓存自定义属性，如果用户的网络断开，则 适用于 Mac 的 Outlook 中的邮件外接程序将无法访问其自定义属性。


### <a name="custom-properties-example"></a>自定义属性示例


下面的示例演示使用自定义属性的 Outlook 外接程序的一组简化的方法。可以将此示例用作使用自定义属性的外接程序的起点。 

此示例包括以下方法：


- [Office.initialize](../../reference/shared/office.initialize.md) -- 初始化外接程序并从 Exchange 服务器中加载自定义属性包。
    
-  **customPropsCallback** -- 获取并保存从服务器返回的自定义属性包以供将来使用。
    
-  **updateProperty** -- 设置或更新特定属性，然后将更改保存到服务器。
    
-  **removeProperty** -- 从属性包中删除特定属性，然后将该删除操作保存到服务器。
    



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## <a name="additional-resources"></a>其他资源

    
- 
  [MAPI 属性概述](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- 
  [Outlook 属性概述](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [从 Outlook 外接程序调用 Web 服务](../outlook/web-services.md)
    
- 
  [Exchange 中 EWS 的属性和扩展属性](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- 
  [Exchange 中 EWS 的属性集和响应形状](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


