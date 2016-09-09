
# TableBinding.rowCount 属性
获取表中的行数，作为整数值。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|TableBindings|
|**选择内容中的最后更改**|1.1|

```
var rowCount = bindingObj.rowCount;
```


## 返回值

指定 [TableBinding](../../reference/shared/binding.tablebinding.md) 对象中的行数。


## 注解

通过在 Excel 2013 和 Excel Online 中选择单行（使用“**插入**”选项卡上的“**表**”）插入一个空表格时，两个 Office 主机应用程序都会创建后跟单个空白行的单个标题行。 但如果外接程序的脚本创建这个新插入表格的绑定（例如，通过使用 [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) 方法），然后检查 **rowCount** 属性的值，则返回的值将根据电子表格是在 Excel 2013 中还是在 Excel Online 中打开而有所不同。


- 在桌面上的 Excel 中，**rowCount** 将返回 0（标题后的空白行不计数）。
    
- 在 Excel Online 中， **rowCount** 将返回 1（标题后的空白行计算在内）。
    
可通过检查是否存在  `rowCount == 1` 在脚本中解决返回值不同的问题，如果存在，再检查行是否包含所有空字符串。

在 Access 相关内容应用程序中，出于性能原因， **rowCount** 属性始终返回 -1。


## 示例




```js
function showBindingRowCount() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        write("Rows: " + asyncResult.value.rowCount);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|TableBindings|
|**最低权限级别**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel和 Word 的支持|
|1.1|增加了对 Access 相关外接程序的支持。|
|1.0|引入|
