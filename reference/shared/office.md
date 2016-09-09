

# Office 对象
表示外接程序的实例，该实例提供对 API 的高级对象的访问权限。

|||
|:-----|:-----|
|**主机：**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**包含最后一次更改的版本**|1.1|

```js
Office
```


## 成员


**属性**

|||
|:-----|:-----|
|名称|说明|
|[context](../../reference/shared/office.context.md)|获取表示外接程序的运行时环境和提供对 API 的高级对象的访问的 Context 对象。|
|[cast.item](../../reference/shared/office.cast.item.md)|在 Visual Studio 中提供专门用于撰写模式或阅读模式下的邮件和约会的 IntelliSense。 <br/><br/><blockquote>**注释**  仅适用于设计时在 Visual Studio 中开发 Outlook 外接程序的情况。 </blockquote>|

**方法**

|||
|:-----|:-----|
|名称|说明|
|[select](../../reference/shared/office.select.md)|创建承诺以基于传入的选择器字符串返回绑定。|
|[useShortNamespace](../../reference/shared/office.useshortnamespace.md)|切换完整  **Office** 命名空间的 **Microsoft.Office.WebExtension** 别名。|

**事件**

|||
|:-----|:-----|
|名称|说明|
|[初始化](../../reference/shared/office.initialize.md)|加载运行时环境和外接程序准备好开始与应用和托管文档交互时发生。|

## 备注

借助 **Office** 对象，开发者可以对 Initialize 事件实现回调函数，并提供对 [Context](../../reference/shared/context.md) 对象的访问权限。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此对象。空的单元格表示相应的 Office 主机应用程序不支持此对象。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|**Office for iPad**|**适用于设备的 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**外接程序类型**|内容、Outlook、任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录


|**版本**|**更改内容**|
|:-----|:-----|
|1.1|增加了对 Office for iPad 中 Excel、PowerPoint 和 Word 的支持。|
|1.1|<ul><li>对于 <a href="6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf.htm">context</a>，添加了对 Access 相关内容外接程序中获取运行时上下文的支持。</p></li><li><p>对于 <a href="23aeb136-da1f-4127-a798-99dc27bc4dae.htm">select</a>，添加了对 Access 相关内容外接程序中选择表绑定的支持。</li><li>对于 <a href="9a4d5c7d-fcc4-4e8f-bef2-f2a8d8b4ae00.htm">useShortNamespace</a>，添加了对 Access 相关内容外接程序的支持。</li><li>对于 <a href="727adf79-a0b5-48d2-99c7-6642c2c334fc.htm">initialize</a>，添加了对 Access 相关内容外接程序中初始化的支持。</li></ul>|
|1.0|引入|

