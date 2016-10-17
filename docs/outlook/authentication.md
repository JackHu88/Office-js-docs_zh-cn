
# <a name="authenticate-an-outlook-add-in-by-using-exchange-identity-tokens"></a>使用 Exchange 标识令牌对 Outlook 外接程序进行身份验证

您的 Outlook 外接程序可以从 Internet 上的任何地方为您的客户提供信息，无论是从承载该外接程序的服务器、您的内部网络还是云中的其他位置。但是，如果该信息受保护，则您的外接程序将需要一种方法将 Exchange 电子邮件帐户与您的信息服务相关联。Exchange 2013 可以提供一个标识发出请求的电子邮件帐户的令牌，以此来为您的外接程序启用单一登录 (SSO)。您可将此令牌与您应用程序的注册用户相关联，以便每次外接程序连接到您的服务时可以识别用户。

## <a name="identity-tokens"></a>标识令牌


我们的两个示例外接程序使用公开信息，一个显示邮件中地址的 Bing 地图，另一个显示邮件中 YouTube 视频链接的预览。但是，您的外接程序也可以访问非公开信息。您可以使用承载您外接程序的服务器将外接程序链接到您内部网络中的信息，或链接到云中的任何位置。

可以利用各种不同技术来标识外接程序用户以及对其进行身份验证。Exchange 2013 为您的外接程序提供一个标识特定 Exchange 电子邮件帐户的标识令牌，以此来简化用户身份验证过程。您可在服务中将此令牌与注册用户相关联，为使用 Outlook 外接程序的客户启用单一登录 (SSO)。 

要在加载项中使用 SSO，代码将执行以下操作：


* 调用 Outlook 外接程序 API 中的一个函数以返回一个标识令牌。
* 将该令牌连同一个请求一起发送到您的服务器。
* 解包来自服务器的响应以显示来自您服务的信息。
    
在服务器端，情况稍微有些复杂。当服务器收到来自 Outlook 外接程序的请求时，将作如下处理：

* 服务器将验证令牌。你可以使用我们的[托管令牌验证库](../../docs/outlook/use-the-token-validation-library.md)，也可以为你的服务[创建你自己的库](../../docs/outlook/validate-an-identity-token.md)。
* 服务器会查找令牌中的唯一标识符，确定它是否与某个已知标识相关联。你的服务必须对你服务的已知用户[实现与标识符匹配的方法](../../docs/outlook/authenticate-a-user-with-an-identity-token.md)。
* 如果唯一标识符与之前在服务器上使用一组凭据存储的标识符相匹配，则您的服务器可以使用请求的信息进行响应，且无需客户登录到您的服务。
* 如果这个唯一标识符是未知的，则服务器发送一个响应，要求用户用服务器的凭据登录。
* 如果凭据与服务器上的一个已知标识匹配，则您可以将该标识映射到令牌中的唯一标识符，这样，下次请求传入时，您的服务器无需其他登录步骤即可响应。

 >**注释**  这只是有关如何使用标识令牌的一个建议。和以往一样，您在处理标识和身份验证事宜时，一定要确保您的代码满足组织的安全要求。

在以下文章中，我们通过一个简单的 Outlook 外接程序来谈谈具体细节。该外接程序向 Web 服务发送标识令牌和在邮件中发现的电话号码列表。 

- [Exchange 标识令牌揭秘](../outlook/inside-the-identity-token.md)
- [在 Exchange 中使用标识令牌从 Outlook 外接程序调用服务](../outlook/call-a-service-by-using-an-identity-token.md)
- [使用 Exchange 令牌验证库](../outlvalidate-an-identity-token.md ook/use-the-token-validation-library.md)
- [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md )
- [使用 Exchange 的标识令牌对用户进行身份验证](../outlook/validate-an-identity-token.md)


## <a name="additional-resources"></a>其他资源



- [Outlook 外接程序](../outlook/outlook-add-ins.md)
    
- [从 Outlook 外接程序调用 Web 服务](../outlook/web-services.md)
    


