
# <a name="bindingtype-enumeration"></a>BindingType 枚举
 指定应返回的绑定对象的类型。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**包含最后一次更改的版本**|1.1|

```
Office.BindingType
```


## <a name="members"></a>成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.BindingType.Matrix|"matrix"|不带标题行的表格数据。数据作为数组的数组返回，例如在此表单中：` [[row1column1, row1column2],[row2column1, row2column2]]`|
|Office.BindingType.Table|"table"|带有标题行的表格数据。数据作为 [TableData](../../reference/shared/tabledata.md) 对象返回。|
|Office.BindingType.Text|"text"|纯文本。数据作为连续文本的字符返回。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**|Y|||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|添加了对在 Access 相关应用程序中绑定表数据的支持。|
|1.0|引入。|
