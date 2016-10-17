
# <a name="gototype-enumeration"></a>GoToType 枚举
指定要导航到的位置或对象类型。

|||
|:-----|:-----|
|**主机：**|Excel、PowerPoint 和 Word|
|**添加内容的版本**|1.1|

```js
Office.GoToType
```


## <a name="members"></a>成员


**值**


|**枚举**|**值**|**说明**|**支持的客户端**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|"binding"|转至使用特定绑定 ID 的绑定对象。|Excel</br>Word|
|Office.GoToType.NamedItem|"namedItem"|使用该项目的名称转到项目，例如分配到表或范围的名称。在 Excel 中，您可以使用任何对命名范围或表格的结构化引用："Worksheet2!Table1"|Excel|
|Office.GoToType.Slide|"slide"|转至使用特定 ID 的幻灯片。|PowerPoint|
|Office.GoToType.Index|"index"|转至按幻灯片编号或枚举进行的特定索引：</br>**Office.Index.First**</br>**Office.Index.Last**</br>**Office.Index.Next**</br>**Office.Index.Previous**|PowerPoint|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。


有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|引入|
