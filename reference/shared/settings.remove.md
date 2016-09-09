

# Settings.remove 方法
移除指定设置。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|Settings|
|**包含最后一次更改的版本**|1.1|

```js
Office.context.document.settings.remove(name);
```


## 参数



_名称_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**字符串**

&nbsp;&nbsp;&nbsp;&nbsp;要移除的设置的区分大小写的名称。

    



## 注解

 **null** 是设置的有效值。因此，将 **null** 分配给设置不会将它从设置属性包中删除。


 >**重要提示**：请注意，**Settings.remove** 方法只会对设置属性包的内存中副本产生影响。若要在调用 **Settings.remove** 方法之后的某个时间点以及在关闭外接程序之前保留文档中指定设置的删除状态，你必须调用 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法。


## 示例




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
}
```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。



||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

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
|1.1|增加了对在 Access 相关内容外接程序中创建自定义设置的支持。|
|1.0|引入|
