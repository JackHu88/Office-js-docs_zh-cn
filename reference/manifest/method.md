
# Method 元素
指定来自适用于 Office 的 JavaScript API 的单个方法，Office 外接程序需要该方法才能激活。

 **外接程序类型：**内容、任务窗格


## 语法：


```XML
<Method Name="string "/>
```


## 包含在：

 _ [方法](../../reference/manifest/methods.md)_


## 属性



|**属性**|**类型**|**必需**|**说明**|
|:-----|:-----|:-----|:-----|
|Name|string|必需|指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。|

## 注解

**Methods** 和 **Method** 元素不受邮件外接程序的支持。有关要求集的详细信息，请参阅[指定 Office 主机和 API 要求](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_intro)。


 >**重要** 因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。有关如何执行此操作的详细信息，请参阅[了解适用于 Office 的 JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md#HostAPISupport_UsingIfStatements)。

