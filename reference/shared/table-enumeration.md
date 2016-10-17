
# <a name="table-enumeration"></a>Table 枚举
指定_表格式方法_的 [cellFormat](../../docs/excel/format-tables-in-add-ins-for-excel.md) 参数中 `cells:` 属性的枚举值。

|||
|:-----|:-----|
|**主机：**|Excel|
|**已添加**|1.1|

```
Office.Table
```

## <a name="members"></a>成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.Table.All|"all"|Office.Table.Data|
|Office.Table.Data|"data"|Office.Table.Headers|
|Office.Table.Headers|"headers"|仅标题行。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 的支持。|
|1.1|引入|
