
# TableData.headers 属性
获取或设置表的标题。

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|TableBindings|
|**包含最后一次更改的版本**|1.1|

```
var hasHeaders = tableBindingObj.headers;
```


## 返回值

 如果表有标题，返回 **true**；否则返回 **false**。 


## 注解

若要指定标题，你必须指定对应于表结构的数组的数组。 例如，若要指定包含两列的表的标题，你需要将 **header** 属性设置为 ` [['header1', 'header2']]`。

如果你将 **headers** 属性指定为 **null**（或者在构造 **TableData** 对象时将此属性留空），那么代码的执行结果如下：


- 如果插入新表，则将创建表的默认列标题。
    
- 如果覆盖或更新现有表，则不会改动现有标题。
    

## 示例

以下示例将创建具有一个标题和三行的单列表。


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}

```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Word Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.0|引入|
