
# <a name="tablebinding-object"></a>TableBinding 对象
表示两个维度的行和列的绑定，标题可选。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint、Project、Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|TableBindings|
|**包含 Selection 最后一次更改的版本**|1.1|

```
TableBinding
```


## <a name="members"></a>成员


**属性**


|**名称**|**说明**|**Office.js v1.1 的更新**|
|:-----|:-----|:-----|
|[columnCount](../../reference/shared/binding.tablebinding.columncount.md)|获取指定 **TableBinding** 对象中的列数。|添加了对 Access 相关内容外接程序中表绑定的支持。|
|[hasHeaders](../../reference/shared/binding.tablebinding.hasheaders.md)|如果指定的 **TableBinding** 具有标头，则返回 true；否则返回 false。|添加了对 Access 相关内容外接程序中表绑定的支持。|
|[rowCount](../../reference/shared/binding.tablebinding.rowcount.md)|获取指定 **TableBinding** 对象中的行数。|出于性能原因，Access 相关内容外接程序中始终返回 - 1。|

**方法**


|**名称**|**说明**|**Office.js v1.1 的更新**|
|:-----|:-----|:-----|
|[addColumnsAsync](../../reference/shared/binding.tablebinding.addcolumnsasync.md)|将列和值添加到表中。||
|[addRowsAsync](../../reference/shared/binding.tablebinding.addrowsasync.md)|将行和值添加到表中。|添加了对 Access 相关内容外接程序中表绑定的支持。|
|[clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md)|清除绑定表中的格式。|适用于 Excel 相关外接程序的 Office.js v1.1 中的新功能。|
|[deleteAllDataValuesAsync](../../reference/shared/binding.tablebinding.deletealldatavaluesasync.md)|删除表中的所有非标题行及其值，对主机应用程序进行相应切换。|添加了对 Access 相关内容外接程序中表绑定的支持。|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|将数据写入指定的绑定对象表示的文档的绑定部分。|<ul><li>添加了对 Access 相关内容应用程序中表绑定的支持。</li><li>添加了对在将数据写入 Excel 相关外接程序中绑定表时设置格式的支持。</li></ul>|
|[setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|针对绑定表中指定项和数据设置单元格和表格式。|可以在 Excel 相关外接程序中设置表格式。|
|[setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md)|更新绑定表上的表格式选项。|可以在 Excel 相关应用程序中设置表格式。|

## <a name="remarks"></a>备注

**TableBinding** 对象从 [Binding](../../reference/shared/binding.id.md) 抽象对象继承 [id](../../reference/shared/binding.type.md) 属性、[type](../../reference/shared/binding.getdataasync.md) 属性、[getDataAsync](../../reference/shared/binding.setdataasync.md) 方法和 [setDataAsync](../../reference/shared/binding.md) 方法。

在 Excel 中建立表绑定后，用户每向表中添加一个新行都会自动包括在绑定中（ **rowCount** 会增加）。


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**最低权限级别**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|在 Excel 中添加了对[插入表时设置格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)的支持。|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
