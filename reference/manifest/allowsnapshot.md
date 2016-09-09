
# AllowSnapshot 元素
指定是否将内容外接程序的快照图像与主机文档一起保存。

 **外接程序类型：**内容


## 语法：


```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```


## 包含在：

[OfficeApp](../../reference/manifest/officeapp.md)


## 注解


 **安全说明：** **AllowSnapshot** 在默认情况下为 **true**。 这样，用户在不支持 Office 外接程序的主机应用程序版本中打开文档时，即可看到该外接程序的图像，或者如果主机应用程序无法连接到托管外接程序的服务器时，会提供该外接程序的静态图像。 但是，这也意味着可以直接从托管该外接程序的文档访问显示在外接程序中的潜在敏感信息。

