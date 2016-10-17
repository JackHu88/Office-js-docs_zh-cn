
# <a name="documentactiveviewchangedeventargs-object"></a>DocumentActiveViewChangedEventArgs 对象
提供有关引发 [ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) 事件的视图的信息。

|||
|:-----|:-----|
|**主机：**|PowerPoint|
|**引入版本**|1.1|



## <a name="members"></a>成员


**属性**


|**名称**|**说明**|
|:-----|:-----|
|[activeView](../../reference/shared/document.activeviewchangedeventargs.activeview.md)|获取  **ActiveView** 枚举值，该值指定文档活动视图的状态，例如，用户是否可以编辑文档。|
|[type](../../reference/shared/document.activeviewchangedeventargs.type.md)|获取标识被引发事件的类型的 **EventType** 枚举值。|

## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**引入版本**|1.1|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|
