
# <a name="call-web-services-from-an-outlook-add-in"></a>从 Outlook 外接程序调用 web 服务

您的外接程序可使用运行 Exchange Server 2013 的计算机中的 Exchange Web 服务 (EWS)，该 Web 服务可在为外接程序的 UI 提供源位置的服务器上获得，也可在 Internet 上获得。本文提供展示 Outlook 外接程序如何从 EWS 请求信息的示例。

您用来调用 Web 服务的方法随 Web 服务所在的位置的不同而不同。表 1 列出了可以基于位置调用 Web 服务的不同方法。


**表 1.从 Outlook 外接程序调用 Web 服务的方式**


|**Web 服务位置**|**调用 Web 服务的方法**|
|:-----|:-----|
|托管客户端邮箱的 Exchange 服务器|使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法可调用外接程序支持的 EWS 操作。承载邮箱的 Exchange 服务器还会公开 EWS。|
|为加载项 UI 提供源位置的 Web 服务器|使用标准 JavaScript 技术调用 Web 服务。UI 框架中的 JavaScript 代码将在提供 UI 的 Web 服务器的上下文中运行。因此，此代码可以调用该服务器上的 Web 服务，而不会导致出现跨网站脚本错误。|
|所有其他位置|为提供 UI 源位置的 Web 服务器上的 Web 服务创建代理。如果您不提供代理，跨网站脚本错误将阻止外接程序运行。提供代理的一种方式是使用 JSON/P。有关详细信息，请参阅 [Office 外接程序的隐私和安全性](../../docs/develop/privacy-and-security.md)。|

## <a name="using-the-makeewsrequestasync-method-to-access-ews-operations"></a>使用 makeEwsRequestAsync 方法访问 EWS 操作


可以使用 [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法向承载用户邮箱的 Exchange 服务器发出 EWS 请求。

EWS 服务支持 Exchange 服务器中的不同操作；例如复制、查找、更新或发送项目的项目级操作，以及创建、获取或更新文件夹的文件夹级操作。若要执行 EWS 操作，请创建一个执行该操作的 XML SOAP 请求。当操作完成时，您将获得包含该操作相关数据的 XML SOAP 响应。EWS SOAP 请求和响应遵循 Messages.xsd 文件中定义的架构。正如其他 EWS 架构文件一样，Message.xsd 文件位于承载 EWS 的 IIS 虚拟目录中。 

若要使用  **makeEwsRequestAsync** 方法启动 EWS 操作，需要提供以下项：


- 针对该 EWS 操作的 SOAP 请求的 XML，作为  _data_ 形参的实参
    
- 回调方法（作为  _callback_ 实参）
    
- 该回调方法的任何可选输入数据（作为  _userContext_ 实参）
    
EWS SOAP 请求完成后，Outlook 将使用一个实参（是一个 [AsyncResult](../../reference/outlook/simple-types.md) 对象）调用该回调方法。该回调方法可以访问 **AsyncResult** 对象的两个属性：包含该 EWS 操作的 XML SOAP 响应的 **value** 属性，以及包含作为 **userContext** 形参传递的所有数据的 **asyncContext** 可选属性。通常，回调方法稍后会解析 SOAP 响应中的 XML 以获取所有相关信息，并相应地处理这些信息。


## <a name="tips-for-parsing-ews-responses"></a>解析 EWS 响应的提示


解析 EWS 操作的 SOAP 响应时，请注意下列与浏览器相关的问题：


- 指定使用 DOM 方法  **getElementsByTagName** 时标记名称的前缀，以支持 Internet Explorer。
    
     **getElementsByTagName** 的行为方式不同，具体取决于浏览器类型。例如，EWS 响应可包含以下 XML（出于显示目的进行了格式化和缩写）：
    
```XML
      <t:ExtendedProperty><t:ExtendedFieldURI PropertySetId="00000000-0000-0000-0000-000000000000" 
    PropertyName="MyProperty" 
    PropertyType="String"/>
    <t:Value>{
    ...
    }</t:Value></t:ExtendedProperty>
```

 如下所示的代码可在 Chrome 等浏览器上运行，使 XML 由 **ExtendedProperty** 标记括起来：

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("ExtendedProperty");
```


   
 在 Internet Explorer 上，必须包含标记名称的 `t:` 前缀，如下所示：

```js
    var mailbox = Office.context.mailbox;
    mailbox.makeEwsRequestAsync(mailbox.item.itemId), function(result) {
        var response = $.parseXML(result.value);
        var extendedProps = response.getElementsByTagName("t:ExtendedProperty");
```

- 使用 DOM 属性  **textContent** 获取 EWS 响应中标记的内容，如下所示：
    
```
      content = $.parseJSON(value.textContent);
```

 对于 EWS 响应中的某些标记，其他属性（如 **innerHTML**）可能无法在 Internet Explorer 上正常运行。
    

## <a name="example"></a>示例


以下示例调用  **makeEwsRequestAsync** 以使用 [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) 操作获取项目主题。此示例包括以下三个函数：


-  `getSubjectRequest` — 将项目 ID 作为输入，并返回 SOAP 请求的 XML 以便为指定项目调用 **GetItem**。
    
-  `sendRequest` — 调用 `getSubjectRequest` 以获取所选项目的 SOAP 请求，然后将 SOAP 请求和回调方法 `callback` 传递到 **makeEwsRequestAsync** 以获得指定项目的主题。
    
-  `callback` — 处理包含有关指定项目的任何主题和其他信息的 SOAP 响应。
    

```js
function getSubjectRequest(id) {
   // Return a GetItem operation request for the subject of the specified item. 
   var result = 
'<?xml version="1.0" encoding="utf-8"?>' +
'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
'  <soap:Header>' +
'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
'  </soap:Header>' +
'  <soap:Body>' +
'    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
'      <ItemShape>' +
'        <t:BaseShape>IdOnly</t:BaseShape>' +
'        <t:AdditionalProperties>' +
'            <t:FieldURI FieldURI="item:Subject"/>' +
'        </t:AdditionalProperties>' +
'      </ItemShape>' +
'      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
'    </GetItem>' +
'  </soap:Body>' +
'</soap:Envelope>';

   return result;
}





function sendRequest() {
   // Create a local variable that contains the mailbox.
   var mailbox = Office.context.mailbox;

   mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
   var result = asyncResult.value;
   var context = asyncResult.context;

   // Process the returned response here.
}


```


## <a name="ews-operations-that-add-ins-support"></a>外接程序支持的 EWS 操作


Outlook 外接程序可通过  **makeEwsRequestAsync** 方法访问 EWS 中可用的操作子集。如果您不熟悉 EWS 操作以及使用 **makeEwsRequestAsync** 方法访问操作的方式，则可以使用 SOAP 请求示例开始自定义您的 _data_ 实参。以下是如何使用 **makeEwsRequestAsync** 方法的说明：


1. 在 XML 中，用适当值替换所有项目 ID 和相关 EWS 操作属性。
    
2. 加入 SOAP 请求作为  _makeEwsRequestAsync_ 的 **data** 形参的实参。
    
3. 指定回调方法并调用  **makeEwsRequestAsync**。
    
4. 在回调方法中，验证 SOAP 响应中操作的结果。
    
5. 根据需要使用 EWS 操作的结果。
    
下表列出了外接程序支持的 EWS 操作。若要查看 SOAP 请求和响应的示例，请选择各操作对应的链接。有关 EWS 操作的详细信息，请参阅 [在交换 EWS 操作](http://msdn.microsoft.com/library/cf6fd871-9a65-4f34-8557-c8c71dd7ce09%28Office.15%29.aspx)。


**表 2.支持的 EWS 操作**


|**EWS 操作**|**说明**|
|:-----|:-----|
|
  [CopyItem 操作](http://msdn.microsoft.com/library/bcc68f9e-d511-4c29-bba6-ed535524624a%28Office.15%29.aspx)|在 Exchange 存储的指定文件夹中复制指定项目并在其中放入新项目。|
|
  [CreateFolder 操作](http://msdn.microsoft.com/library/6f6c334c-b190-4e55-8f0a-38f2a018d1b3%28Office.15%29.aspx)|在 Exchange 存储中的指定位置创建文件夹。|
|
  [CreateItem 操作](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx)|在 Exchange 存储中创建指定项目。|
|
  [FindConversation 操作](http://msdn.microsoft.com/library/2384908a-c203-45b6-98aa-efd6a4c23aac%28Office.15%29.aspx)|在 Exchange 存储的指定文件夹中枚举会话列表。|
|
  [FindFolder 操作](http://msdn.microsoft.com/library/7a9855aa-06cc-45ba-ad2a-645c15b7d031%28Office.15%29.aspx)|查找指定文件夹的子文件夹并返回描述这组子文件夹的一组属性。|
|
  [FindItem 操作](http://msdn.microsoft.com/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx)|标识位于 Exchange 存储的指定文件夹中的项目。|
|
  [GetConversationItems 操作](http://msdn.microsoft.com/library/8ae00a99-b37b-4194-829c-fe300db6ab99%28Office.15%29.aspx)|在会话中获取排列为节点的一个或多个项集。|
|
  [GetFolder 操作](http://msdn.microsoft.com/library/355bcf93-dc71-4493-b177-622afac5fdb9%28Office.15%29.aspx)|从 Exchange 存储中获取文件夹的指定属性和内容。|
|
  [GetItem 操作](http://msdn.microsoft.com/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx)|从 Exchange 存储中获取项目的指定属性和内容。|
|
  [MarkAsJunk 操作](http://msdn.microsoft.com/library/1f71f04d-56a9-4fee-a4e7-d1034438329e%28Office.15%29.aspx)|将电子邮件移动到"垃圾邮件"文件夹，并相应地在阻止的发件人名单中添加或删除邮件的发件人。|
|
  [MoveItem 操作](http://msdn.microsoft.com/library/dcf40fa7-7796-4a5c-bf5b-7a509a18d208%28Office.15%29.aspx)|将项目移动到 Exchange 存储中的单个目标文件夹。|
|
  [SendItem 操作](http://msdn.microsoft.com/library/337b89ef-e1b7-45ed-92f3-8abe4200e4c7%28Office.15%29.aspx)|发送位于 Exchange 存储中的电子邮件。|
|
  [UpdateFolder 操作](http://msdn.microsoft.com/library/3494c996-b834-4813-b1ca-d99642d8b4e7%28Office.15%29.aspx)|修改 Exchange 存储中现有文件夹的属性。|
|
  [UpdateItem 操作](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx)|修改 Exchange 存储中现有项目的属性。|

## <a name="authentication-and-permission-considerations-for-the-makeewsrequestasync-method"></a>makeEwsRequestAsync 方法的身份验证和权限注意事项


在使用  **makeEwsRequestAsync** 方法时，将使用当前用户的电子邮件帐户凭据对请求进行验证。 **makeEwsRequestAsync** 方法将管理您的凭据，这样您就不必随请求一起提供身份验证凭据。


 >
  **注释**  服务器管理员必须使用 [New-WebServicesVirtualDirctory](http://technet.microsoft.com/en-us/library/bb125176.aspx)或 [Set-WebServicesVirtualDirecory](http://technet.microsoft.com/en-us/library/aa997233.aspx) cmldet 将客户端访问服务器 EWS 目录上的 _OAuthAuthentication_ 形参设置为 **true**，以使用  **makeEwsRequestAsync** 方法发出 EWS 请求。

你的外接程序必须在其外接程序清单中指定 **ReadWriteMailbox** 权限才能使用 **makeEwsRequestAsync** 方法。有关使用 **ReadWriteMailbox** 权限的信息，请参阅[了解 Outlook 外接程序的权限](../outlook/understanding-outlook-add-in-permissions.md#readwritemailbox-permission)中的 [ReadWriteMailbox 权限](../outlook/understanding-outlook-add-in-permissions.md)部分。


## <a name="additional-resources"></a>其他资源



- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
- [Office 外接程序的隐私和安全性](../../docs/develop/privacy-and-security.md)
    
- [解决 Office 外接程序中的同源策略限制](../../docs/develop/addressing-same-origin-policy-limitations.md)
    
- 
  [Exchange 的 EWS 引用](http://msdn.microsoft.com/library/2a873474-1bb2-4cb1-a556-40e8c4159f4a%28Office.15%29.aspx)
    
- 
  [Outlook 和 Exchange 中的 EWS 的邮件应用程序](http://msdn.microsoft.com/library/821c8eb9-bb58-42e8-9a3a-61ca635cba59%28Office.15%29.aspx)
    
请参阅下文，了解如何使用 ASP.NET Web API 为外接程序创建后端服务：


- [使用 ASP.NET Web API 为 Office 外接程序创建 Web 服务](http://blogs.msdn.com/b/officeapps/archive/2013/06/10/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx)
    
- [使用 ASP.NET Web API 构建 HTTP 服务的基础知识](http://www.asp.net/web-api)
    
