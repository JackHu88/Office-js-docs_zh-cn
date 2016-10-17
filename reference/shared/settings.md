
# <a name="settings-object"></a>Settings 对象
表示作为名称/值对存储在主机文档中的任务窗格或内容外接程序的自定义设置。

|||
|:-----|:-----|
|**主机：**|Access、Excel、PowerPoint 和 Word|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|Settings|
|**包含最后一次更改的版本**|1.1|

```
Office.context.document.settings
```


## <a name="members"></a>成员


**方法**

|||
|:-----|:-----|
|名称|说明|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|为 **settingsChanged** 事件添加事件处理程序。|
|[get](../../reference/shared/settings.get.md)|检索指定设置。|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|读取文档中保存的所有设置并刷新外接程序在内存中保留的这些设置的副本。|
|[remove](../../reference/shared/settings.remove.md)|移除指定设置。|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|为 **settingsChanged** 事件删除事件处理程序。|
|[saveAsync](../../reference/shared/settings.saveasync.md)|保存设置。|
|[set](../../reference/shared/settings.set.md)|设置或创建指定设置。|

**事件**


|**名称**|**说明**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|设置更改时发生。|

## <a name="remarks"></a>备注

通过使用  **Settings** 对象的方法创建的设置将按外接程序和按文档进行保存。即，这些设置仅供创建它们的外接程序使用，并且仅来自保存它们的文档。

设置的名称为  **string**，而值可以为 **string**、 **number**、 **boolean**、 **null**、 **object** 或 **array**。

**Settings** 对象自动作为 [Document](../../reference/shared/document.md) 对象的一部分进行加载，并且在外接程序激活时通过调用相应对象的 [settings](../../reference/shared/document.settings.md) 属性激活。开发者负责在添加或删除设置后调用 [saveAsync](../../reference/shared/settings.saveasync.md) 方法，从而将设置保存到文档中。


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**在要求集中可用**|Settings|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录

|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|对于 **addHandlerAsync** 和 **removeHandlerAsync** 方法，增加了针对 Access 相关内容外接程序中的事件添加和删除事件处理程序的支持。对于 **get**、**refreshAsync**、**remove**、**saveAsync**，和 **set** 方法，增加了对 Access 相关内容外接程序中自定义设置的支持。|
|1.0|引入|