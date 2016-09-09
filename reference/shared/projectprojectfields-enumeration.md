
# ProjectProjectFields 枚举
指定可供 **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** 方法用作参数的项目域。

|||
|:-----|:-----|
|**主机：**|Project|
|**在其中添加**|1.0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    DurationUnits: 3,
    GUID: 4, 
    Finish: 5, 
    Start: 6, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## 成员


****


|**成员**|**说明**|
|:-----|:-----|
|**CurrencyDigits**|货币十位后的位数。|
|**CurrencySymbol**|货币符号。|
|**CurrencySymbolPosition**|货币符号位置：未指定 = -1 ；值前面没有空格 ($0) = 0 ；值后面没有空格 (0$) = 1；值前面有一个空格 ($ 0) = 2；值后面有一个空格 (0 $) = 3。|
|**GUID**|项目的 GUID。|
|**Finish**|项目完成日期|
|**开始**|项目起始日期。|
|**ReadOnly**|指定项目是否为只读。|
|**版本**|项目版本。|
|**WorkUnits**|项目的工时单位，如天或小时。|
|**ProjectServerUrl**|Project Web App URL，针对存储在 Project 服务器中的项目。|
|**WSSUrl**|SharePoint URL，针对与 SharePoint 列表同步的项目。|
|**WSSList**|SharePoint 列表的名称，针对与任务列表同步的项目。|

## 备注

**ProjectProjectFields** 常数可用作 **[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** 方法的参数。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## 另请参阅



#### 其他资源


[getProjectFieldAsync 方法](../../reference/shared/projectdocument.getprojectfieldasync.md)
