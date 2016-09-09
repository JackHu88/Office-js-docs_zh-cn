
# TableData 对象
表示表中或 [TableBinding](../../reference/shared/binding.tablebinding.md) 中的数据。

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|TableBindings|
|**在其中添加**|1.1|

```
TableData
```

## 成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|[headers](../../reference/shared/tabledata.headers.md)|获取或设置表中的标题。|
|[rows](../../reference/shared/tabledata.rows.md)|获取或设置表中的行。|

## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Word Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel和 Word 的支持|
|1.0|引入|
