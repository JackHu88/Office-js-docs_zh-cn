
# Context.officeTheme 属性
提供了访问 Office 主题颜色的属性。

 **重要提示：**此 API 目前只适用于 Windows 桌面上 [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) 中的 Excel、Outlook、PowerPoint 和 Word。


|||
|:-----|:-----|
|**主机：**|Excel、Outlook、PowerPoint、Word|
|**在[要求集](../../docs/overview/specify-office-hosts-and-api-requirements.md)中可用**|不在集合中|
|**在其中添加**|1.3|



```js
Office.context.officeTheme
```


## 成员


**属性**

|||
|:-----|:-----|
|名称|说明|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|获取 Office 主题正文的背景色。|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|获取 Office 主题正文的前景色。|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|获取 Office 主题控件的背景色。|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|获取 Office 主题控件的前景色。|

## 注解

通过使用 Office 主题颜色，你可以使外接程序的配色方案与用户（通过“**文件**” > “**Office 帐户**” > “**Office 主题**”UI）选择的当前 Office 主题协调一致，这种做法适用于所有 Office 主机应用程序。 使用 Office 主题颜色适用于 Outlook 和任务窗格外接程序。


## 示例


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## 支持详细信息



|||
|:-----|:-----|
|**最低权限级别**|[受限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**外接程序类型**|内容、任务窗格、Outlook|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.3|引入|
