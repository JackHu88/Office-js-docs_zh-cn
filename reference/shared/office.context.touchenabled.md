
# <a name="context.touchenabled-property"></a>Context.touchEnabled 属性
获取外接程序是否将运行在已启用触控的 Office 主机应用程序中。

|||
|:-----|:-----|
|**主机：**|Excel 和 Word|
|**包含最后一次更改的版本**|1.1|

```
var isTouchEnabled = Office.context.touchEnabled;
```


## <a name="return-value"></a>返回值

如果外接程序是在触控设备（如 iPad）上运行，则返回 **True**；否则返回 **False**。


## <a name="remarks"></a>备注

使用  **touchEnabled** 属性确定您的外接程序何时在触摸设备上运行；如有必要，调整控件类型以及外接程序 UI 中元素的大小和间距以适应触摸交互。


## <a name="support-details"></a>支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅 [运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。

||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**|Y|
|**Word**|Y|

|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## <a name="support-history"></a>支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|引入。|
