# <a name="settings.settingschangedeventargs-object"></a>Settings.settingschangedeventargs 对象
提供有关引发了 [settingsChanged 事件](settings.settingschangedevent.md)的设置的信息。

|||
|:-----|:-----|
|**主机：**|Access、Excel |
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Settings|
|**包含最后一次更改的版本**|1.0|

```js
Office.EventType.SettingsChanged
```

## <a name="members"></a>成员

**属性**

|**名称**|**说明**|
|:-----|:-----|
|**[settings](settings.settingschangedeventargs.setting.md)**|获取表示引发了 settingsChanged 事件的设置的 **Settings** 对象。|
|**[type](settings.settingschangedeventargs.type.md)**|获取用于标识所引发事件的种类的 **EventType** 枚举值。|

## <a name="remarks"></a>备注

若要添加 **settingsChanged** 事件的事件处理程序，请使用 [Settings](settings.addhandlerasync.md) 对象的 **addHandlerAsync** 方法。

只有当外接程序的脚本调用 **Settings.saveAsync** 方法将设置的内存中副本保留到文档文件中时，才会触发 **settingsChanged** 事件。调用 **Settings.set** 或 [Settings.remove](settings.set.md) 方法时，不会触发 [settingsChanged](settings.remove.md) 事件。

当您的应用程序在共享（合著）文档中使用时， **settingsChanged** 事件旨在让您处理两个或两个以上用户试图同时保存设置时出现的潜在冲突。


 >**重要提示**：当外接程序正在与任意 Excel 客户端搭配运行时，外接程序的代码可以注册 **settingsChanged** 事件的处理程序。不过，只有当用 Excel Online 中打开的电子表格加载外接程序，_并且_多个用户正在编辑电子表格（共同创作）时才会触发此事件。因此，实际上只有采用共同创作方案的 Excel Online 才支持 **settingsChanged** 事件。



## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||


|||
|:-----|:-----|
|**在要求集中可用**|Settings|
|**最低权限级别**|受限|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录

|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|
