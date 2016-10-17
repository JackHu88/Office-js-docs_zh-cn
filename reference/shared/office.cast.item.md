
# <a name="office.cast.item-property"></a>Office.cast.item 属性
提供特定于撰写或阅读模式的邮件和约会的 IntelliSense。

|||
|:-----|:-----|
|**主机：**|Outlook|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Mailbox|
|**包含最后一次更改的版本**|1.0|



|||
|:-----|:-----|
|**适用的 Outlook 模式**|仅限在 Visual Studio 中的设计时间|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## <a name="return-value"></a>返回值

使您能够为 Outlook 外接程序选择合适的 IntelliSense 的一组方法。


## <a name="remarks"></a>备注

此属性及其方法仅在 Visual Studio 上支持在开发 Outlook 外接程序时使用 IntelliSense。它们对其他开发工具不起任何作用。

**Office.cast.item** 方法用于 Visual Studio 中的设计时，可为 **Office.context.mailbox.item** 属性提供特定的 IntelliSense。例如，当你使用 **toAppointmentCompose** 方法时，IntelliSense 只会显示适用于撰写模式的 **Appointment** 方法和属性。

在运行时， **Office.cast.item** 方法对您的 Outlook 外接程序无效。


## <a name="example"></a>示例

以下示例使用  **toMessageCompose** 方法来转换 **Office.context.mailbox.item** 属性，以便可以在撰写模式下仅显示 **Message** 对象的 IntelliSense。转换之后， `message` 变量将仅显示用于撰写模式的方法和属性的 IntelliSense。


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||Office for Windows Desktop|Office Online（在浏览器中）|Outlook for Mac|
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



|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|
