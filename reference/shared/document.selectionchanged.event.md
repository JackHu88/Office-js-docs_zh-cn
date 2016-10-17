
# <a name="document.selectionchanged-event"></a>Document.SelectionChanged 事件
文档中的选择更改时发生。

|||
|:-----|:-----|
|**主机：**|Excel、PowerPoint 和 Word|
|**引入版本**|1.1|

```
Office.EventType.DocumentSelectionChanged
```

## <a name="remarks"></a>备注

若要为文档的 **SelectionChanged** 事件添加事件处理程序，请使用 [Document](../../reference/shared/document.addhandlerasync.md) 对象的 **addHandlerAsync** 方法。


## <a name="example"></a>示例




```
function addEventHandlerToDocument() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs) {
    doSomethingWithDocument(eventArgs.document);
}

```




## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.0|引入|
