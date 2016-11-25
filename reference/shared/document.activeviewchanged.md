
# <a name="documentactiveviewchanged-event"></a>Document.ActiveViewChanged 事件
用户更改文档的当前视图时出现。

|||
|:-----|:-----|
|**主机：**|PowerPoint|
|**引入版本**|1.1|

```
Office.EventType.ActiveViewChanged
```


## <a name="remarks"></a>备注

若要为文档的 **ActiveViewChanged** 事件添加事件处理程序，请使用 [Document](../../reference/shared/document.addhandlerasync.md) 对象的 **addHandlerAsync** 方法。事件处理程序会接收 [ActiveViewChangedEventArgs](../../reference/shared/document.activeviewchangedeventargs.md) 类型的参数。


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for Mac**|**Office for iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y||Y|Y|

>**注意：此事件不会在 PowerPoint Online 应用场景中触发，因为幻灯片放映模式被视为新会话。若要获取活动视图，必须在 Office.Initialize 期间查询它。
 

|||
|:-----|:-----|
|**引入版本**|1.1|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|
