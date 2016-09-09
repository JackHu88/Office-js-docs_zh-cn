
# 比较 Outlook for Mac 与其他 Outlook 主机之间的 Outlook 外接程序支持的差异

您可以按照与其他主机（包括 Outlook for Windows、适用于设备的 OWA 和 Outlook Web App）相同的方式在 适用于 Mac 的 Outlook 中创建和运行 Outlook 外接程序，而无需为每个主机自定义 JavaScript。从外接程序到 适用于 Office 的 JavaScript API 的相同调用通常按同一方式工作，下表中说明的领域除外。

 >**注释**  适用于 Mac 的 Outlook 仅在 Outlook 阅读模式中支持适用于 Office 的 JavaScript API。

|**区域**|**Outlook for Windows、适用于设备的 OWA、Outlook Web App**|**Outlook for Mac**|
|:-----|:-----|:-----|
|office.js 和 Office 外接程序清单架构支持的版本|Office.js 和架构 v1.1 中的所有 API。|<ul><li>仅适用于阅读模式的 API。可以激活使用 office.js v1.1 中新的和可扩展的 API 的外接程序，但用于撰写模式的这些 API 无法在适用于 Mac 的 Outlook 上正常运行。 </li><li>Schema v1.1。</li></ul>|
|定期约会系列实例|<ul><li>可以获得主约会的项目 ID 和其他属性或定期系列约会的实例 </li><li>可以使用 [mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md#displayappointmentformitemid) 显示定期序列的实例或主项目。</li></ul>|<ul><li>可以获得主约会的项目 ID 和其他属性，但无法获得定期系列约会的实例</li><li>可以显示定期系列的主约会。不显示项目 ID 和定期系列的实例。</li></ul>|
|约会参与者的收件人类型|可以使用 [EmailAddressDetails.recipientType](../../reference/outlook/simple-types.md) 标识参与者的收件人类型。|**EmailAddressDetails.recipientType** 将返回约会参与者的 **undefined**。|
|主机版本字符串 |[diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) 返回的版本字符串的格式取决于主机的实际类型。例如：<ul><li>适用于 Windows 的 Outlook：15.0.4454.1002</li><li>Outlook Web App：15.0.918.2</li></ul>|由 适用于 Mac 的 Outlook 上的  **Diagnostics.hostVersion** 返回的版本字符串实例：15.0 (140325)|
|项目自定义属性|如果网络出现故障，外接程序仍可以访问缓存的自定义属性。|由于 适用于 Mac 的 Outlook 不缓存自定义属性，因此，如果网络出现故障，外接程序将无法对其进行访问。|
|附件详细信息|[AttachmentDetails](../../reference/outlook/Office.context.mailbox.md) 对象中的内容类型和附件名称取决于主机的类型：<ul><li><b>AttachmentDetails.contentType</b> 的 JSON 示例：<b>"contentType": "image/x-png"</b>。 </li><li><b>AttachmentDetails.name</b> 不包含任何文件名扩展名。例如，如果附件是一封主题为“RE: Summer activity”的邮件，则表示附件名称的 JSON 对象将为 <b>"name": "RE: Summer activity"</b>。</li></ul>|<ul><li><b>AttachmentDetails.contentType</b> 的 JSON 示例：<b>"contentType": "image/png"</b></li><li><b>AttachmentDetails.name</b> 始终包含一个文件名扩展名。作为邮件项目的附件包含 .eml 扩展名，约会包含 .ics 扩展名。例如，如果附件是主题为“RE: Summer activity”的电子邮件，那么表示附件名称的 JSON 对象为 <b>"name": "RE: Summer activity.eml"</b>。</li></ul>|
|字符串表示  **dateTimeCreated** 和 **dateTimeModified** 属性中的时区|例如：2014 年 3 月 13 日，星期四，14:09:11 GMT+0800（中国标准时间）|例如：2014 年 3 月 13 日，星期四，14:09:11 GMT+0800 (CST)|
|**dateTimeCreated** 和 **dateTimeModified** 的时间准确度|如果外接程序使用以下代码，准确度精确到毫秒。<br/><pre lang="javascript">JSON.stringify(Office.context.mailbox.item, null, 4);</pre>|准确度精确到秒。|

## 其他资源



- [部署和安装 Outlook 外接程序以进行测试](../outlook/testing-and-tips.md)
    
