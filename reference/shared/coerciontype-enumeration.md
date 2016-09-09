
# CoercionType 枚举
指定如何强制由调用方法返回或设置的数据。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**邮箱中的最后更改**|1.1|

```js
Office.CoercionType
```

## 成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.CoercionType.Html|"html"|作为 HTML 返回或设置数据。<br/><br/> **注意**  仅适用于 Word 相关外接程序以及 Outlook 相关 Outlook 外接程序（撰写模式）中的数据。|
|Office.CoercionType.Matrix|"matrix"|以不带标题的表格数据形式返回或设置数据。 以包含一维连续文本的字符的数组的数组形式返回或设置数据。 例如，三行两列 **string** 值应为：` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`。<br/><br/> **注意**  仅适用于 Excel 和 Word 中的数据。|
|Office.CoercionType.Ooxml|"ooxml"|以 Office Open XML 形式返回或设置数据。<br/><br/> **注意**  仅适用于 Word 中的数据。|
|Office.CoercionType.SlideRange|"slideRange"|返回一个 JSON 对象，包含所选幻灯片的 ID、标题和索引的数组。例如，`{"slides":[{"id":257,"title":"Slide 2","index":2},{"id":256,"title":"Slide 1","index":1}]}` 对应于两个幻灯片的选区。<br/><br/> **注意**  仅适用于调用 [Document.getSelectedData](../../reference/shared/document.getselecteddataasync.md) 方法获取当前幻灯片或所选幻灯片范围时 PowerPoint 中的数据。|
|Office.CoercionType.Table|"table"|以带可选标题的表格数据形式返回或设置的数据。 以具有可选标题的数组的数组形式返回或设置数据。<br/><br/> **注意**  仅适用于 Access、Excel 和 Word 中的数据。|
|Office.CoercionType.Text|"text"|以文本 ( **string**) 形式返回或设置数据。以一维连续文本的字符形式返回或设置数据。|
|Office.CoercionType.Image|"image"|以映像流形式返回或设置数据。<br/><br/> **注意**  仅适用于 Excel、Word 和 PowerPoint 中的数据。|
PowerPoint 仅支持 **Office.CoercionType.Text**、**Office.CoercionType.Image** 和 **Office.CoercionType.SlideRange**。

Project 仅支持  **Office.CoercionType.Text**。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**Office for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**|Y|||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**外接程序类型**|内容、 Outlook（撰写模式）、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Word Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.1|增加了对[撰写模式 Outlook 外接程序](../../docs/outlook/compose-scenario.md)的支持。|
|1.0|引入|
