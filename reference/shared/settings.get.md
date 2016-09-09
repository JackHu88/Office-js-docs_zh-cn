
# Settings.get 方法
检索指定设置。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|Settings|
|**包含最后一次更改的版本**|1.1|

```js
var mySetting = Office.context.document.settings.get(name);
```


## 参数



_名称_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**字符串**

&nbsp;&nbsp;&nbsp;&nbsp;要检索的设置的区分大小写的名称。

    



## 返回值

具有映射到 JSON 序列化值的属性名称的  **object**。


## 示例




```js
function displayMySetting() {
    write('Current value for mySetting: ' + Office.context.document.settings.get('mySetting'));
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此属性。空的单元格表示相应的 Office 主机应用程序不支持此属性。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|Settings|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录




|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 PowerPoint Online 的支持。|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|增加了对在 Access 相关内容外接程序中创建设置的支持。|
|1.0|引入|
