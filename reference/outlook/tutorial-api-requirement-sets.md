 

# <a name="understanding-outlook-api-requirement-sets"></a>了解 Outlook API 要求集

Outlook 外接程序通过在其[清单](https://msdn.microsoft.com/EN-US/library/office/dn592036.aspx)中使用 [Requirements](https://msdn.microsoft.com/en-us/library/office/fp123693.aspx) 元素来声明所需要的 API 版本。Outlook 外接程序始终包括 `Name` 属性设置为 `Mailbox` 和 `MinVersion` 属性设置为支持外接程序方案的 API 最低要求集的 [Set](https://msdn.microsoft.com/EN-US/library/office/dn592049.aspx) 元素。

例如，下面的清单段表示 1.1 的最低要求集：

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

所有 Outlook API 均属于`Mailbox`[要求集](https://msdn.microsoft.com/EN-US/library/office/dn535871.aspx#SpecifyRequirementSets_intro)。`Mailbox`要求集具有不同版本，我们发布的每个新 API 集均属于较高版本的要求集。并非所有 Outlook 客户端都支持最新的 API 集，但如果某个 Outlook 客户端声明支持某个要求集，它将支持该要求集中的所有 API。

在清单中设置最低要求集版本可控制外接程序会显示在哪个 Outlook 客户端中。如果客户端不支持最低要求集，则不会加载外接程序。例如，如果指定要求集版本 1.3，则意味着外接程序不会显示在任何不支持 1.3 及以上版本的 Outlook 客户端中。

## <a name="using-apis-from-later-requirement-sets"></a>使用更高版本要求集中的 API

设置要求集不会限制外接程序可使用的可用 API。例如，如果外接程序指定要求集 1.1，但它在支持 1.3 的 Outlook 客户端中运行，则外接程序可以使用要求集 1.3 中的 API。

要使用较新的 API，开发人员可使用标准 JavaScript 技术来检查是否存在新 API

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

对于清单中所指定的要求集版本中的任何 API，无需执行此类检查。

## <a name="choosing-a-minimum-requirement-set"></a>选择最低要求集

开发人员应使用包含其方案关键 API 集的最早要求集，如果不使用该要求集，外接程序将不起作用。

## <a name="clients"></a>客户端

下列客户端支持 Outlook 外接程序。

| 客户端 | 受支持的 API 要求集 |
| --- | --- |
| Outlook 2016 for Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2016 for Mac | 1.1 |
| Outlook 2013 for Windows | 1.1、1.2、1.3 |
| Outlook 网页版（Office 365 和 Outlook.com） | 1.1, 1.2, 1.3, 1.4 |
| Outlook Web App（本地 Exchange 2013） | 1.1 |
| Outlook Web App（本地 Exchange 2016） | 1.1, 1.2. 1.3 |
>**注意** 对 Outlook 2013 中的 1.3 版本的支持已作为 [2015 年 12 月 8 日 Outlook 2013 更新 (KB3114349) 的一部分添加](https://support.microsoft.com/en-us/kb/3114349)    
