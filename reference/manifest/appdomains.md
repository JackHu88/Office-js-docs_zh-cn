
# <a name="appdomains-element"></a>AppDomains 元素
指定 Office 外接程序将用于加载页面的任何其他域。

 **外接程序类型：**内容、任务窗格、邮件


## <a name="syntax:"></a>语法：


```XML
<AppDomains>
   ...
</AppDomains>
```


## <a name="contained-in:"></a>包含在：

[OfficeApp](../../reference/manifest/officeapp.md)


## <a name="can-contain:"></a>可以包含：

[AppDomain](../../reference/manifest/appdomain.md)


## <a name="remarks"></a>注解

**AppDomains** 和 **AppDomain** 元素用于指定除在 [SourceLocation](../../reference/manifest/sourcelocation.md) 元素中指定的域之外的任何其他域。有关详细信息，请参阅 [Office 外接程序 XML 清单](../../docs/overview/add-in-manifests.md)。

