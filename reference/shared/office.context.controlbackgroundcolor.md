
# officeTheme.controlBackgroundColor 属性
获取 Office 主题控件的背景色。

 **重要提示：**此 API 目前只适用于 Windows 桌面上 [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) 中的 Excel、Outlook、PowerPoint 和 Word。



|||
|:-----|:-----|
|**主机：**|Excel、Outlook、PowerPoint、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|不在集合中|
|**在其中添加**|1.3|

```
var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
```


## 返回值

十六进制三元色。


## 注解

返回的颜色对应于用户（通过“**文件**” > “**Office 帐户**” > “**Office 主题**”UI）选择的 Office 主题值，这种做法适用于所有 Office 主机应用程序。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||||
|**Outlook**|Y||||
|**PowerPoint**|Y||||
|**Word**|Y||||

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.3|引入|
