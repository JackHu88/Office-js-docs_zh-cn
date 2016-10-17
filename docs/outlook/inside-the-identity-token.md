
# <a name="inside-the-exchange-identity-token"></a>Exchange 标识令牌揭秘
了解 Exchange 2013 标识令牌包含的内容。



Exchange 服务器发送到您的 Outlook 外接程序的身份验证标识令牌对您的外接程序并非透明的；您不需要为了将令牌发送到您的服务器而了解令牌中的内容。但是，如果要编写与您的 Outlook 外接程序交互的 Web 服务代码，则需要了解标识令牌包含的内容。

## <a name="what-is-an-identity-token?"></a>什么是标识令牌？


标识令牌是一个 64 位编码的 URL 字符串，由发送它的 Exchange 服务器自签名。令牌不经过加密，您用来验证签名的公钥存储在颁发该令牌的 Exchange 服务器上。令牌包含三个部分：标头、有效负载和签名。在令牌字符串中，各个部分用"."字符分隔开，以便您拆分令牌。

Exchange 2013 使用 JSON Web Token (JWT) 作为标识令牌。有关 JWT 令牌的信息，请参阅 [JSON Web Token (JWT) Internet 草案](http://self-issued.info/docs/draft-goland-json-web-token-00.html)。


### <a name="identity-token-header"></a>标识令牌标头

标头用来标识令牌以便您的 Web 服务知道提供的是什么类型的令牌。下面的示例显示令牌标头的形式。

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "Un6V7lYN-rMgaCoFSTO5z707X-4" }
```

下表描述标识令牌标头的各个部分。


**标识令牌标头的各个部分**


|**声明**|**值**|**说明**|
|:-----|:-----|:-----|
|typ|"JWT"|将令牌标识为 JSON Web Token。Exchange 服务器提供的所有标识令牌都是 JWT 令牌。|
|alg|"RS256"|用来创建签名的哈希算法。Exchange 服务器提供的所有令牌都使用 RS-256 算法。|
|x5t|证书指纹|令牌的 X.509 指纹。|

### <a name="identity-token-payload"></a>标识令牌有效负载

有效负载包含身份验证声明，标识电子邮件帐户和发送令牌的 Exchange 服务器。下面的示例显示有效负载部分的形式。
```js

{ 
   "aud" : "https://mailhost.contoso.com/IdentityTest.html", 
   "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
   "nbf" : "1331579055", 
   "exp" : "1331607855", 
   "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
   "isbrowserhostedapp":"true",
"appctx" : { 
     "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com" "version" : "ExIdTok.V1" "amurl" :         "https://mailhost.contoso.com:443/autodiscover/metadata/json/1" 
     } 
}
```
下表列出标识令牌有效负载的各个部分。


**标识令牌有效负载的各个部分**


|**声明**|**说明**|
|:-----|:-----|
|aud|请求令牌的外接程序的 URL。仅当令牌是从正在客户端浏览器中运行的外接程序发送的时才有效。 如果外接程序使用 Office 外接程序清单架构 v1.1，则此 URL 为第一个  **SourceLocation** 元素中指定的 URL，在 **ItemRead** 或 **ItemEdit** 窗体类型下，先发生者为外接程序清单中 [FormSettings](http://msdn.microsoft.com/en-us/library/0d1a311d-939d-78c1-e968-89ddf7ebc4b4%28Office.15%29.aspx) 元素的一部分。|
|iss|颁发令牌的 Exchange 服务器的唯一标识符。此 Exchange 服务器颁发的所有令牌将具有相同标识符。|
|nbf|令牌开始生效的日期和时间。值是自 1970 年 1 月 1 日以来的秒数。 |
|exp|标记失效的日期和时间，值是自 1970 年 1 月 1 日以来的秒数。|
|appctxsender|发送应用程序上下文的 Exchange 服务器的唯一标识符。|
|isbrowserhostedapp|指示加载项是否承载于浏览器中。|
|appctx|令牌的应用程序上下文。 |
appctx 声明中的信息提供电子邮件帐户的地址和帐户的唯一标识符。下表列出 appctx 声明的各个部分。



|**appctx 声明部分**|**说明**|
|:-----|:-----|
|msexchuid|与电子邮件帐户和 Exchange 服务器关联的唯一标识符。|
|version|令牌的版本号。对于正在运行 Exchange 2013 的服务器提供的所有令牌，该值为"ExIdTok.V1"。|
|amurl|身份验证元数据文档的 URL，该文档包含用于签署令牌的 X.509 证书的公钥。有关如何使用身份验证元数据文档的详细信息，请参阅 [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)。|

### <a name="identity-token-signature"></a>标识令牌签名

通过使用标头中指定的算法，并使用有效负载中指定的服务器位置处的自签名 X 509 证书，对标头和有效负载部分进行哈希处理来创建签名。Web 服务可以验证此签名，以帮助确保标识令牌来自预期的服务器。


## <a name="additional-resources"></a>其他资源



- [使用 Exchange 标识令牌对 Outlook 外接程序进行身份验证](../outlook/authentication.md)
    
- [在 Exchange 中使用标识令牌从 Outlook 外接程序调用服务](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [使用 Exchange 令牌验证库](../outlook/use-the-token-validation-library.md)
    
- [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)
    
