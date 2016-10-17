
# <a name="valueformat-enumeration"></a>ValueFormat 枚举
指定由调用方法返回的值（如数字和日期）返回时应用了其格式设置。

|||
|:-----|:-----|
|**主机：**|Excel、Project、Word|
|**添加内容的版本**|1.0|

```
Office.ValueFormat
```


## <a name="members"></a>成员


**值**


|**枚举**|**值**|**说明**|
|:-----|:-----|:-----|
|Office.ValueFormat.Formatted|"formatted"|返回格式化的数据。|
|Office.ValueFormat.Unformatted|"unformatted"|返回无格式化的数据。|

## <a name="remarks"></a>备注

例如，如果将  _valueFormat_ 参数指定为 `"formatted"`，宿主应用程序中格式化为货币的数字或格式化为 mm/dd/yy 的日期将保留其格式设置。如果将  _valueFormat_ 参数指定为 `"unformatted"`，将以其基础顺序序列号形式返回数据。


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Project**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.0|引入|
