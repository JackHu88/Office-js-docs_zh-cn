
# <a name="officeapp-element"></a>OfficeApp 元素
Office 外接程序清单中的根元素。

 **外接程序类型：**内容、任务窗格、邮件


## <a name="syntax:"></a>语法：


```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```


## <a name="contained-in:"></a>包含在：

 _none_


## <a name="must-contain:"></a>必须包含：



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Id](../../reference/manifest/id.md)|x|x|x|
|[Version](../../reference/manifest/version.md)|x|x|x|
|[ProviderName](../../reference/manifest/providername.md)|x|x|x|
|[DefaultLocale](../../reference/manifest/defaultlocale.md)|x|x|x|
|[DefaultSettings](../../reference/manifest/defaultsettings.md)|x|x|x|
|[DisplayName](../../reference/manifest/displayname.md)|x|x|x|
|[说明](../../reference/manifest/description.md)|x|x|x|
|[FormSettings](../../reference/manifest/formsettings.md)||x||
|[Permissions](../../reference/manifest/permissions.md)|x||x|
|[Rule](../../reference/manifest/rule.md)||x||

## <a name="can-contain:"></a>可以包含：



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[AlternateId](../../reference/manifest/alternateid.md)|x|x|x|
|[IconUrl](../../reference/manifest/iconurl.md)|x|x|x|
|[HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md)|x|x|x|
|[SupportUrl](../../reference/manifest/supporturl.md)|x|x|x|
|[AppDomains](../../reference/manifest/appdomains.md)|x|x|x|
|[Hosts](../../reference/manifest/hosts.md)|x|x|x|
|[Requirements](../../reference/manifest/requirements.md)|x|x|x|
|[AllowSnapshot](../../reference/manifest/allowsnapshot.md)|x|||
|[Permissions](../../reference/manifest/permissions.md)||x||
|[DisableEntityHighlighting](../../reference/manifest/disableentityhighlighting.md)||x||
|[Dictionary](../../reference/manifest/dictionary.md)|||x|
|[VersionOverrides](../../reference/manifest/versionoverrides.md)|X|X|X|

## <a name="attributes"></a>属性


|||
|:-----|:-----|
|xmlns|定义的 Office 外接程序清单命名空间和架构版本。应始终将此属性设置为 `"http://schemas.microsoft.com/office/appforoffice/1.1"`|
|xmlns:xsi|定义 XMLSchema 实例。应始终将此属性设置为 `"http://www.w3.org/2001/XMLSchema-instance"`|
|xsi:type|定义的 Office 外接程序的类型。应始终将此属性设置为下列值之一：`"ContentApp"`、`"MailApp"` 或 `"TaskPaneApp"`|
