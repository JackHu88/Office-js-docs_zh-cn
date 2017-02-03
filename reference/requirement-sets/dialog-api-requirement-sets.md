
# <a name="dialog-api-requirement-sets"></a>Dialog API 要求集

要求集是指各组已命名的 API 成员。Office 外接程序使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持外接程序所需的 API。有关详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)。

Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。

|  要求集  |  Office 2013 for Windows | Office 2016 for Windows*   |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | 内部版本 15.0.4855.1000 或更高版本 | 版本 1602（内部版本 6741.0000）或更高版本 | 1.22 或更高版本 | 15.20 或更高版本| 2017 年 1 月 | 版本 1608（内部版本 7601.6800）或更高版本|

>**注意：**通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。若要使用对话框 API，请运行 Office 更新程序，获取最新版本。 

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- 
  [更新频道发布的 Office 365 客户端版本号和内部版本号](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [使用的是哪一版 Office？](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- 
  [在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- 
  [Office Online Server 概述](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集
若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="dialog-api-11"></a>Dialog API 1.1 
Dialog API 1.1 是首版 API。有关 API 的详细信息，请参阅 [Dialog API](../shared/officeui.md) 参考主题。

## <a name="additional-resources"></a>其他资源

- [指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office 外接程序 XML 清单](../../docs/overview/add-in-manifests.md)

