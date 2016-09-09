
# MatrixBinding 对象
表示两个维度的行和列的绑定。 

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|MatrixBindings|
|**选择内容中的最后更改**|1.1|

```
MatrixBinding
```


**属性**


|**名称**|**说明**|
|:-----|:-----|
|[columnCount](../../reference/shared/binding.matrixbinding.columncount.md)|获取矩阵数据结构中的列数，作为整数值。|
|[rowCount](../../reference/shared/binding.matrixbinding.rowcount.md)|获取矩阵数据结构中的行数，作为整数值。|

## 备注

**MatrixBinding** 对象从 [Binding](../../reference/shared/binding.id.md) 对象继承 [id](../../reference/shared/binding.type.md) 属性、[type](../../reference/shared/binding.getdataasync.md) 属性、[getDataAsync](../../reference/shared/binding.setdataasync.md) 方法和 [setDataAsync](../../reference/shared/binding.md) 方法。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|MatrixBindings|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.0|引入|
