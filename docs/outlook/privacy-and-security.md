
# <a name="privacy,-permissions,-and-security-for-outlook-add-ins"></a>Outlook 外接程序的隐私、权限和安全性
最终用户、开发人员和管理员可以使用 Outlook 外接程序的安全模型的分层权限级别来控制隐私和性能。



本文介绍了 Outlook 外接程序可以请求的可能权限，并从以下几个角度审视安全模型：

- Office 商店 — 外接程序完整性。
    
- 最终用户 — 隐私和性能问题。
    
- 开发人员 — 权限选择和资源使用限制。
    
- 管理员 — 设置性能阈值的权限。
    

## <a name="permissions-model"></a>权限模型


客户对外接程序安全的理解可能会影响外接程序采用情况，因此 Outlook 外接程序安全依赖于一个多层权限模型。Outlook 外接程序可能会公开其所需的权限级别，从而确定外接程序可以对客户邮箱数据采取的可能访问和操作。 

清单架构版本 1.1 包含四个级别的权限。 


**表 1.外接程序权限级别**


|**权限级别**|**Outlook 外接程序清单中的值**|
|:-----|:-----|
|受限|受限|
|读取项目|ReadItem|
|读/写项目|ReadWriteItem|
|读/写邮箱|ReadWriteMailbox|
四个级别的权限具有累积性：**读/写邮箱**权限包括**读/写项目**权限、**读取项目**权限和**受限**权限；**读/写项目**权限包括**读取项目**权限和**受限**权限；**读取项目**权限包括**受限**权限。图 1 显示了四个级别的权限并说明了每一层提供给最终用户、开发人员和管理员的功能。有关这些权限的详细信息，请参阅 [最终用户：隐私和性能问题](#end-users-privacy-and-performance-concerns)、[开发人员：权限选择和资源使用限制](#developers-permission-choices-and-resource-usage-limits) 和 [了解 Outlook 外接程序权限](../outlook/understanding-outlook-add-in-permissions.md)。 


**图 1.将四层权限模型与最终用户、开发人员和管理员关联**

![邮件应用程序架构 v1.1 的 4 层权限模型](../../images/olowa15wecon_Permissions_4Tier.png)


## <a name="office-store:-add-in-integrity"></a>Office 商店：加载项完整性


Office 商店托管最终用户和管理员可安装的外接程序。Office 商店强制执行以下措施来维护这些 Outlook 外接程序的完整性：


- 要求加载项的主机服务器始终使用安全套接字层 (SSL) 进行通信。
    
- 要求开发人员在提交加载项时提供身份证明、合约协议和适合的隐私策略。 
    
- 以只读模式存档加载项。
    
- 支持针对可用加载项的用户审阅系统以推广自我管理的社区。
    

## <a name="end-users:-privacy-and-performance-concerns"></a>最终用户：隐私和性能问题


安全模型通过下列方式解决最终用户的安全、隐私和性能问题：


- 受 Outlook 信息权限管理 (IRM) 保护的最终用户邮件不与 Outlook 外接程序交互。
    
- 从 Office 商店安装外接程序之前，最终用户能够查看外接程序可以对其数据进行的访问和采取的操作，且必须明确确认后才能继续操作。未经用户或管理员手动验证，Outlook 外接程序不会自动推送到客户端计算机。
    
- 授予“**受限**”权限可允许 Outlook 外接程序仅具有对当前项目的有限访问权限。授予“**读取项目**”权限可允许 Outlook 外接程序仅访问当前项目上的个人识别信息，例如发件人和收件人姓名以及电子邮件地址。
    
- 最终用户仅能为他/她自己安装低信任度的 Outlook 外接程序。对组织产生影响的 Outlook 外接程序由管理员安装。
    
- 最终用户可以安装支持上下文相关方案的低信任度 Outlook 外接程序，这不仅对用户具有吸引力，同时还可以最大限度地降低用户的安全风险。
    
- 已安装 Outlook 外接程序的清单文件在用户电子邮件帐户中受到保护。
    
- 通过托管 Office 外接程序的服务器传送的数据始终根据安全套接字层 (SSL) 协议进行加密。
    
- 仅适用于 Outlook 富客户端：Outlook 富客户端监视已安装 Outlook 外接程序的性能，实施管治控制，以及禁用在以下方面超过限制的 Outlook 外接程序：
    
      - 激活响应时间
    
  - 激活或重新激活失败次数
    
  - 内存使用率
    
  - CPU 使用率
    

    管治可阻止拒绝服务攻击并将外接程序性能保持在合理的水平。业务栏通知最终用户 Outlook 富客户端已根据此类管治控制禁用的 Outlook 外接程序。
    
- 无论何时，最终用户都可以验证所安装 Outlook 外接程序请求的权限，在 Exchange 管理中心禁用或随后启用任何 Outlook 外接程序。
    

## <a name="developers:-permission-choices-and-resource-usage-limits"></a>开发人员：权限选择和资源使用限制


安全模型向开发人员提供精细级别的权限以供选择，以及严格的性能准则以供遵循。


### <a name="tiered-permissions-increases-transparency"></a>多层权限将增加透明度

开发人员应按照多层权限模型提供透明度，并解决用户有关哪些加载项可以处理其数据和邮箱的问题，间接促进加载项采用：


- 开发人员根据 Outlook 外接程序应激活的方式、Outlook 外接程序读取或写入项目特定属性的需求，或者创建和发送项目的需求来针对 Outlook 外接程序请求适当级别的权限。
    
- 开发人员使用 Outlook 外接程序清单中的 [Permissions](http://msdn.microsoft.com/en-us/library/c20cdf29-74b0-564c-e178-b75d148b36d1%28Office.15%29.aspx) 元素，并根据需要分配 **Restricted**、 **ReadItem**、 **ReadWriteItem** 或 **ReadWriteMailbox** 的值来请求权限。 
    
     >**注意**  注意从清单架构 v1.1 开始就提供 **ReadWriteItem** 权限。

    下面的示例请求“**读取项目**”权限。
    


```XML
  <Permissions>ReadItem</Permissions>
```

- 如果 Outlook 外接程序激活特定类型的 Outlook 项目（约会或邮件）或存在于项目主题或正文中的特定提取的实体（电话号码、地址、URL），开发人员可以请求“**受限**”权限。例如，如果在当前邮件的主题或正文中找到一个或多个实体（共三个）- 电话号码、邮寄地址或 URL，以下规则将激活 Outlook 外接程序。
    
```XML
  <Permissions>Restricted</Permissions>
    <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
    <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
        <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
    </Rule>
</Rule>
```

- 如果 Outlook 外接程序需要读取当前项目的属性而非默认提取实体的属性，或者需要通过当前项目上的外接程序写入自定义属性集，但无需读写其他项目或在用户的邮箱中创建或发送邮件，则开发人员应请求“**读取项目**”权限。例如，如果 Outlook 外接程序需要寻找项目主体或正文中的会议建议、任务建议、电子邮件地址或联系人姓名等实体，或者需要使用一个正则表达式来激活，则开发人员应请求“**读取项目**”权限。
    
- 如果 Outlook 外接程序需要写入撰写的项目的属性（如收件人姓名、电子邮件地址、正文和主题），或需要添加或删除项目附件，那么开发人员应请求“**读/写项目**”权限。
    
- 仅在 Outlook 外接程序需要使用 [mailbox.makeEWSRequestAsync](../../reference/outlook/Office.context.mailbox.md) 方法执行下列一个或多个操作时，开发人员才请求“**读/写邮箱**”权限：
    
      - 读取或写入邮箱中项目的属性。
    
  - 创建、读取、写入或发送邮箱中的项目。
    
  - 创建、读取或写入邮箱文件夹。
    

### <a name="resource-usage-tuning"></a>资源使用调整

开发人员应注意激活资源的使用限制，在他们的开发工作流中加入性能调整功能，以便减少主机对低性能外接程序的拒绝服务机会。开发人员应遵循 [Outlook 外接程序的激活和 JavaScript API 的限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)中所述的设计激活规则准则。如果 Outlook 外接程序适合运行于 Outlook 富客户端之上，那么开发人员应验证该外接程序能否在资源使用限制之内执行。


### <a name="other-measures-to-promote-user-security"></a>提高用户安全性的其他措施

开发人员还应该注意并规划以下内容：


- 开发人员无法在加载项中使用 ActiveX 控件，因为它们不受支持。
    
- 开发人员在将 Outlook 外接程序提交到 Office 商店时应执行以下操作：
    
      - 生成扩展验证 (EV) SSL 证书作为身份证明。
    
  - 在支持 SSL 的 Web 服务器上承载其提交的加载项。
    
  - 生成合规隐私策略。
    
  - 准备好在提交加载项后签订合约协议。
    

## <a name="administrators:-privileges"></a>管理员：权限


安全模型向管理员提供以下权限和责任：


- 可以阻止最终用户安装任何 Outlook 外接程序，包括 Office 商店上的外接程序。
    
- 可以在 Exchange 管理中心上禁用或启用任何 Outlook 外接程序。
    
- 仅适用于 Outlook for Windows：可以通过 GPO 注册表设置覆盖性能阈值设置。
    


## <a name="additional-resources"></a>其他资源



- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
- [Office 外接程序的隐私和安全性](../../docs/develop/privacy-and-security.md)
    
- [Outlook 外接程序 API](../outlook/apis.md)
    
- [在内容和任务窗格外接程序中有关 API 使用的请求权限](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
    
- [Outlook 外接程序的激活和 JavaScript API 限制](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
