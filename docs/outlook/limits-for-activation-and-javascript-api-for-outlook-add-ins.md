
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Outlook 外接程序的激活和 JavaScript API 的限制

为了向 Outlook 外接程序的用户提供令人满意的体验，您必须了解特定的激活和 API 使用准则，并执行外接程序使其不超过这些限制。这些准则的设置是为了确保单个外接程序不能请求 Exchange Server 或 Outlook 而浪费过长时间来处理其激活规则或对适用于 Office 的 JavaScript API 调用，从而影响 Outlook 和其他外接程序的整体用户体验。在外接程序清单中设计激活规则，使用自定义属性、漫游设置、收件人、Exchange Web 服务 (EWS) 服务请求和响应以及异步调用时，均须遵守这些限制。 

 >**注意** 如果您的外接程序在 Outlook 富客户端中运行，您还必须确认外接程序是在特定运行时资源使用率限制内运行的。 


## <a name="limits-for-activation-rules"></a>激活规则的限制


为 Outlook 外接程序设计激活规则时，请遵循以下准则：


- 将清单的大小限制为 256 KB。如果超出该限制，则无法为 Exchange 邮箱安装 Outlook 外接程序。

- 可为外接程序最多指定 15 条激活规则。如果超出该限制，则无法安装外接程序。
    
- 如果您对所选项目的正文使用 [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) 规则，预计 Outlook 富客户端将仅对正文的前 1 MB 应用规则，而不会超过此限制应用于正文的其他部分。如果正文的前 1 MB 之后存在匹配，您的外接程序将不会激活。如果您期望这成为一种可能的方案，请重新设计激活条件。
    
- 如果你在 **ItemHasKnownEntity** 或 [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) 规则中使用正则表达式，请注意，通常适用于任何 Outlook 主机的下列限制和准则，以及表 1、2 和 3 中所述的限制和准则（因主机而异）：
    
      - 在外接程序的激活规则中最多仅指定五个正则表达式。如果超出该限制，将无法安装外接程序。
    
  - 指定正则表达式，以便 **getRegExMatches** 方法调用在前 50 个匹配项内返回预期结果。
    
  - 可以在正则表达式中指定向前断言，但不支持向后 (?<=text) 和否定向后 (?<!text) 断言。
    

表 1 列出了这些限制并介绍了 Outlook 富客户端与 Outlook Web App 或适用于设备的 OWA 之间正则表达式支持的区别。这种支持不依赖于任何特定类型的设备和项目正文。


 **表 1.各种正则表达式支持的一般区别**


|**Outlook 富客户端**|**Outlook Web App 或适用于设备的 OWA**|
|:-----|:-----|
|使用作为 Visual Studio 标准模板库的一部分提供的 C++ 正则表达式引擎。此引擎符合 ECMAScript 5 标准。 |使用属于 JavaScript 一部分的正则表达式评估，由浏览器提供，且支持 ECMAScript 5 超集。|
|由于正则表达式引擎不同，预计包含基于预定义字符类的自定义字符类的正则表达式将返回与 Outlook Web App 或适用于设备的 OWA 中的 Outlook 富客户端不同的结果。<br/><br/>例如，正则表达式“[\s\S]{0,100}”与一个空格字符或非空格字符的 0 到 100 之间的任意数匹配。此正则表达式在 Outlook 富客户端中与在 Outlook Web App 和适用于设备的 OWA 中返回的结果不相同。作为解决方法，您应将正则表达式重写为""(\s|\S){0,100}"。此变通正则表达式与一个空格字符或非空格字符的 0 到 100 之间的任意数匹配。<br/><br/>您应该在每个 Outlook 主机中对每个正则表达式进行充分的测试，并在正则表达式返回不同的结果时重写该正则表达式。 |您应该在每个 Outlook 主机中对每个正则表达式进行充分的测试，并在正则表达式返回不同的结果时重写该正则表达式。|
|默认情况下，外接程序的所有正则表达式的计算时间限制为 1 秒。超出此限制将导致最多重新计算 3 次。如果超出该重新计算限制，Outlook 富客户端将禁止对任何 Outlook 主机上的同一邮箱运行外接程序。<br/><br/>管理员可使用 **OutlookActivationAlertThreshold** 和 **OutlookActivationManagerRetryLimit** 注册表项覆盖这些计算限制。|不支持与 Outlook 富客户端中相同的资源监视或注册表设置。但将为所有 Outlook 主机上的同一邮箱禁用 Outlook 富客户端上需要很长计算时间的正则表达式的外接程序。|

表 2 列出了这些限制并介绍了每一个 Outlook 应用了正则表达式的项正文部分的区别。如果对项正文应用了正则表达式，则其中某些限制取决于设备和项正文的类型。

**表 2.计算的项正文的大小限制**


||**Outlook 富客户端**|**Outlook Web App、适用于设备的 OWA、OWA for iPad 或 OWA for iPhone**|**Outlook Web App**|
|:-----|:-----|:-----|:-----|
|外形规格|任何支持的设备|Android 智能手机、iPad 或 iPhone|Android 智能手机、iPad 和 iPhone 之外任何支持的设备|
|纯文本项正文|对正文数据的第一个 1 MB 而不对超出该限制的其余正文应用正则表达式。|仅当正文少于 16,000 个字符时激活加载项。|仅当正文少于 500,000 个字符时激活加载项。|
|HTML 项正文|对正文数据的第一个 512 KB 而不对超出该限制的其余正文应用正则表达式。（实际的字符数取决于范围可为每字符 1 到 4 字节的编码。）|对前 64,000 个字符（包括 HTML 标记字符）而不对超出该限制的其余正文应用正则表达式。|仅当正文少于 500,000 个字符时激活加载项。|

表 3 列出了这些限制并介绍了每个 Outlook 主机在计算正则表达式后返回的匹配项的区别。这种支持不依赖于任何特定的设备类型，但是，如果对项正文应用了正则表达式，则该支持可能依赖于项正文的类型。

**表 3.返回的匹配项限制**


||**Outlook 富客户端**|**Outlook Web App 或适用于设备的 OWA**|
|:-----|:-----|:-----|
|返回的匹配项的顺序|假定对于应用于同一个项目的同一个正则表达式， **getRegExMatches** 返回的匹配项在 Outlook 富客户端中与在 Outlook Web App 或适用于设备的 OWA 中的不同。|假定  **getRegExMatches** 返回的匹配项的顺序在 Outlook 富客户端与在 Outlook Web App 或 适用于设备的 OWA 中的不同。|
|纯文本项正文|**getRegExMatches** 返回至多 1,536 (1.5 KB) 个字符的任意匹配项，最多 50 个匹配项。<br/><br/>**注意**：**getRegExMatches** 并不会在返回的数组中以任何特定顺序返回匹配项。通常，假定 Outlook 富客户端中应用于同一项的同一正则表达式的匹配项顺序与 Outlook Web App 和适用于设备的 OWA 中的不同。|**getRegExMatches** 返回的任何匹配项最多为 3,072 个字符 (3 KB) ，最多为 50 个匹配项。|
|HTML 项正文|**getRegExMatches** 返回至多 3,072 (3 KB) 个字符的任意匹配项，最多 50 个匹配项。<br/> <br/> **注意**：**getRegExMatches** 并不会在返回的数组中以任何特定顺序返回匹配项。通常，假定 Outlook 富客户端中应用于同一项的同一正则表达式的匹配项顺序与 Outlook Web App 和适用于设备的 OWA 中的不同。|**getRegExMatches** 返回的任何匹配项最多为 3,072 (3 KB) 个字符，最多为 50 个匹配项。|

## <a name="limits-for-javascript-api"></a>JavaScript API 的限制


除了前面的激活规则的准则外，每个 Outlook 主机还对 JavaScript 对象模型强制实施了特定限制，如表 4 中所述：


**表 4.使用适用于 Office 的 JavaScript API 获取或设置特定数据的限制**


|**功能**|**限制**|**相关 API**|**说明**|
|:-----|:-----|:-----|:-----|
|自定义属性|2500 个字符|[CustomProperties](../../reference/outlook/CustomProperties.md) 对象<br/> <br/>[item.loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法|约会或邮件项目的所有自定义属性的限制。如果外接程序的所有自定义属性的总大小超出此限制，则所有 Outlook 主机将返回错误。|
|漫游设置|32 KB 字符数|[RoamingSettings](../../reference/outlook/RoamingSettings.md) 对象<br/><br/> [context.roamingSettings](../../reference/outlook/Office.context.md) 属性|外接程序的所有漫游设置的限制。如果您的设置超出此限制，则所有 Outlook 主机将返回错误。|
|正在提取已知实体|2000 个字符|[item.getEntities](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/> <br/>[item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/> <br/>[item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md) 方法|Exchange Server 可从项目正文的已知实体中提取的字符限制。Exchange Server 会忽略超出该限制的实体。请注意，这个限制不依赖于外接程序是否使用了  **ItemHasKnownEntity** 规则。|
|Exchange Web 服务|1 MB 字符数|[mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法|**Mailbox.makeEwsRequestAsync** 调用的请求或响应的限制|
|收件人|100 位收件人|[item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) 属性<br/> <br/>[item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) 属性<br/> <br/>[item.resources](../../reference/outlook/Office.context.mailbox.item.md) 属性<br/> <br/>[item.to](../../reference/outlook/Office.context.mailbox.item.md) 属性<br/> <br/>[item.cc](../../reference/outlook/Office.context.mailbox.item.md) 属性<br/> <br/>[Recipients.addAsync](../../reference/outlook/Recipients.md) 方法<br/> <br/>[Recipient.getAsync](../../reference/outlook/Recipients.md) 方法<br/> <br/>[Recipient.setAsync](../../reference/outlook/Recipients.md) 方法|在每个属性中指定的对收件人的限制。|
|显示名称|255 个字符|[EmailAddressDetails.displayName](../../reference/outlook/simple-types.md) 属性<br/><br/> [Recipients](../../reference/outlook/Recipients.md) 对象<br/><br/> **item.requiredAttendees** 属性<br/><br/> **item.optionalAttendees** 属性 <br/><br/>**item.resources** 属性 <br/><br/>**item.to** 属性 <br/><br/>**item.cc** 属性|约会或邮件中显示名称的长度限制。|
|设置主题|255 个字符|[mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [Subject.setAsync](../../reference/outlook/Subject.md) 方法|新的约会窗体中的主题限制，或设置约会或邮件主题的限制。|
|设置地点|255 个字符|[Location.setAsync](../../reference/outlook/Location.md) 方法|设置约会或会议请求地点的限制。|
|新的约会窗体的正文|32 KB 字符数|**Mailbox.displayNewAppointmentForm** 方法|新的约会窗体中正文的限制。|
|显示现有项目的正文|32 KB 字符数|[mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md) 方法|对于 Outlook Web App 和 适用于设备的 OWA：即现有约会或邮件窗体中正文的限制。|
|设置正文|1 MB 字符数|[Body.prependAsync](../../reference/outlook/Body.md) 方法<br/> <br/>[Body.setAsync](../../reference/outlook/Body.md)<br/><br/>[Body.setSelectedDataAsync](../../reference/outlook/Body.md) 方法|设置约会或邮件项目正文的限制。|
|附件数|Outlook Web App 和 适用于设备的 OWA 中可以有 499 个文件|[item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法|可以附加到某个项目进行发送的文件数的限制。通过用户界面和  **addFileAttachmentAsync**，Outlook Web App 和适用于设备的 OWA 通常限制为允许附加最多 499 个文件。Outlook 富客户端不会专门限制附件数。但是，所有 Outlook 主机都遵守用户 Exchange Server 配置的对附件大小的限制。请参阅下一行，获取"附件大小"。|
|附件大小|取决于 Exchange Server|**item.addFileAttachmentAsync** 方法|项目的所有附件的大小都有限制，管理员可以在用户邮箱的 Exchange Server 上对该限制进行配置。对于 Outlook 富客户端，这样便限制了项目的附件数。对于 Outlook Web App 和适用于设备的 OWA，按附件数和所有附件大小限制项目的实际附件，以这两种限制中更少的为准。|
|附件的文件名|255 个字符|**item.addFileAttachmentAsync** 方法|要添加到项目的附件的文件名长度限制。|
|附件的 URI|2048 个字符|**item.addFileAttachmentAsync** 方法|要添加为项目附件的文件名 URI 的限制。|
|附件 ID|100 个字符|[item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/><br/> [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) 方法|要添加或从项目中删除的附件 ID 的长度限制。|
|异步调用|3 次调用|**item.addFileAttachmentAsync** 方法<br/><br/>**item.addItemAttachmentAsync** 方法<br/><br/><br/>**item.removeAttachmentAsync** 方法<br/><br/> [Body.getTypeAsync](../../reference/outlook/Body.md) 方法<br/><br/>**Body.prependAsync** 方法<br/><br/>**Body.setSelectedDataAsync** 方法<br/><br/> [CustomProperties.saveAsync](../../reference/outlook/CustomProperties.md) 方法<br/><br/><br/> [item.LoadCustomPropertiesAysnc](../../reference/outlook/Office.context.mailbox.item.md) 方法<br/><br/><br/> [Location.getAsync](../../reference/outlook/Location.md) 方法<br/><br/>**Location.setAsync** 方法<br/><br/> [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/> [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法<br/><br/>**Recipients.addAsync** 方法<br/><br/> [Recipients.getAsync](../../reference/outlook/Recipients.md) 方法<br/><br/>**Recipients.setAsync** 方法<br/><br/> [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) 方法<br/><br/> [Subject.getAsync](../../reference/outlook/Subject.md) 方法<br/><br/>**Subject.setAsync** 方法<br/><br/> [Time.getAsync](../../reference/outlook/Time.md) 方法<br/><br/> [Time.setAsync](../../reference/outlook/Time.md) 方法|对于 Outlook Web App 或 适用于设备的 OWA：对每次同时异步调用的次数有限制，因为浏览器只允许对服务器进行有限数量的异步调用。 |

## <a name="additional-resources"></a>其他资源



- [部署和安装 Outlook 外接程序以进行测试](../outlook/testing-and-tips.md)
    
- [Outlook 外接程序的隐私、权限和安全性](../outlook/../../docs/develop/privacy-and-security.md)
    
