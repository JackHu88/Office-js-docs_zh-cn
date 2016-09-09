
# Context.mailbox 属性
获取提供对 Outlook 外接程序的 API 程序特别访问的  **mailbox** 对象。

|||
|:-----|:-----|
|**主机：**|Outlook|
|**在 [要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md) 中可用**|邮箱|
|**包含最后一次更改的版本**|1.0|

```js
var outlookOm = Office.context.mailbox;
```


## 返回值


  [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx) 对象。


## 示例

以下代码行访问 JavaScript API for Office 的 [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) 对象。


```js
// Access the Item object.
var item = Office.context.mailbox.item;

```




## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此方法。空的单元格表示相应的 Office 主机应用程序不支持此方法。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**在要求集中可用**|邮箱|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|
