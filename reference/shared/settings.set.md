

# Settings.set 方法
设置或创建指定设置。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint 和 Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|Settings|
|**包含最后一次更改的版本**|1.1|

```js
Office.context.document.settings.set(name, value);
```


## 参数



_名称_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**字符串**

&nbsp;&nbsp;&nbsp;&nbsp;要设置或创建的设置的名称（区分大小写）。

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;类型：**字符串**、**数字**、**布尔值**、**null**、**对象** 或 **数组**

&nbsp;&nbsp;&nbsp;&nbsp;指定要存储的值。
    

## 注解

如果设置尚不存在，那么 **set** 方法会新建一个具有指定名称的设置；或者，此方法会在设置属性包的内存中副本内设置具有指定名称的现有设置。在你调用 [Settings.saveAsync](../../reference/shared/settings.saveasync.md) 方法后，值会作为数据类型的序列化 JSON 表示形式存储在文档中。每个外接程序的设置的大小上限为 2MB。


 >**重要提示**：请注意，**Settings.set** 方法只会对设置属性包的内存中副本产生影响。为了确保对设置所做的增补或更改在文档下次打开时、在调用 **Settings.set** 方法之后的某个时间点以及在关闭外接程序之前对外接程序生效，你必须调用 **Settings.saveAsync** 方法，将设置保留在文档中。


## 示例




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
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
|1.1|增加了对 Access 相关内容外接程序中自定义设置的支持。|
|1.0|引入|
