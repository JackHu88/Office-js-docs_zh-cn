
# Binding.bindingDataChanged 事件
绑定内的数据更改时发生。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Word|
|**BindingEvents 中的最后更改**|1.1|

```js
Office.EventType.BindingDataChanged
```


## 备注

若要为绑定的 **BindingDataChanged** 事件添加事件处理程序，请使用 [Binding](../../reference/shared/binding.addhandlerasync.md) 对象的 **addHandlerAsync** 方法。事件处理程序会接收 [BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md) 类型的参数。


## 示例




```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|BindingEvents|
|**应用程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录

|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel 和 Word 的支持。|
|1.1|添加了对 Access 相关应用程序中此事件的支持。|
|1.0|引入|
