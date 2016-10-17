
# <a name="context.roamingsettings-property"></a>Context.roamingSettings 属性
获取一个对象，它表示保存到用户邮箱的 Outlook 外接程序的自定义设置或状态。

|||
|:-----|:-----|
|**主机：**|Outlook|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Mailbox|
|**包含最后一次更改的版本**|1.0|

```
var appSettings = office.context.roamingSettings;
```


## <a name="return-value"></a>返回值

一个 [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx) 对象。


## <a name="remarks"></a>备注

**RoamingSettings** 对象允许你存储和访问用户邮箱中存储的 Outlook 外接程序数据，以便从用于访问该邮箱的任意主机客户端应用程序运行的外接程序可以使用此类数据。


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|Mailbox|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|
