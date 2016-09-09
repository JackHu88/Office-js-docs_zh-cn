
# SelectionMode 枚举
指定是否选择（突出显示）要导航到的位置（使用 [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) 方法时）。

|||
|:-----|:-----|
|**Office.js 版本中引入**|1.1|

|||
|:-----|:-----|
|**主机：**|Excel、PowerPoint 和 Word|
|**在其中添加**|1.1|



```
Office.SelectionMode
```


## 成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.SelectionMode.Selected|"selected"|将选择（突出显示）的位置。|
|Office.SelectionMode.None|"none"|将光标移到开始位置。|

## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|引入|
