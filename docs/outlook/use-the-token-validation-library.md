
# 使用 Exchange Web 服务托管 API 令牌验证库

可以使用你的外接程序从运行 Exchange Server 2013 或 Exchange Online 的服务器请求的标识令牌来标识你的 Outlook 外接程序的客户端。格式设置为 JSON Web 令牌的令牌为 Exchange 服务器上的电子邮件帐户提供唯一标识符。Exchange Web 服务 (EWS) 托管 API 提供帮助程序类来简化标识令牌的使用。

## 使用验证库的前提条件

若要验证 Exchange 标识令牌，你必须安装 [EWS 托管 API 库](https://www.nuget.org/packages/Microsoft.Exchange.WebServices)。

## 验证 Exchange 标识令牌

EWS 托管 API 验证库提供 **AppIdentityToken** 类来管理 Exchange 标识令牌。下面的方法演示如何创建一个 **AppIdentityToken** 实例并调用 **Validate** 方法来验证该令牌是否有效。该方法采用以下参数：

- *rawToken*：在 Outlook 外接程序中从 [**Office.context.mailbox.getUserIdentityTokenAsync**](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) 方法返回的令牌的字符串表示形式。
- *hostUri*：调用 **getUserIdentityTokenAsync** 的 Outlook 外接程序中页面的完全限定的 URI。

```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
{
    try
    {
        AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
        token.Validate(new Uri(hostUri));

        return token;
    }
    catch (TokenValidationException ex)
    {
        throw new ApplicationException("A client identity token validation error occurred.", ex);
    }
}
```

## 其他资源

- [使用 Exchange 标识令牌对 Outlook 外接程序进行身份验证](../outlook/authentication.md)  
- [Exchange 标识令牌揭秘](../outlook/inside-the-identity-token.md)
- [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)
    
