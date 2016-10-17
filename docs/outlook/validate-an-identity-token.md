
# <a name="validate-an-exchange-identity-token"></a>验证 Exchange 标识令牌

Outlook 外接程序可以向你发送一个标识令牌，但你必须在信任请求之前对该令牌进行验证，以确保它来自你预期的 Exchange 服务器。本文中的示例演示如何通过用 C# 编写的验证对象验证 Exchange 标识令牌；但是，你可以使用任何编程语言来进行验证。[JSON Web 令牌 (JWT) Internet 草案](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl)介绍了验证令牌所需的步骤。 

建议你使用分四步的过程验证标识令牌并获取用户的唯一标识符。首先，从 base64 URL 编码的字符串中提取 JSON Web 令牌 (JWT)。然后，确保该令牌格式正确、它是用于 Outlook 外接程序的令牌、它未过期且你可以提取身份验证元数据文档的有效 URL。接下来，从 Exchange 服务器中检索身份验证元数据文档并验证附加到标识令牌的签名。最后，通过将用户的 Exchange ID 与身份验证元数据文档的 URL 进行哈希运算来计算用户的唯一标识符。整个过程可能看似复杂，但每个步骤其实很简单。你可以从 Web 上的以下位置下载包含这些示例的解决方案：[Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)。
 




## <a name="set-up-to-validate-your-identity-token"></a>进行设置以验证标识令牌


本文中的代码示例依赖于 Windows Identity Foundation (WIF) 以及用 JSON 令牌的处理程序扩展 WIF 的 DLL。可从以下位置下载所需程序集：


- [Windows Identity Foundation](http://msdn.microsoft.com/en-us/security/aa570351)
    
- [用于 32 位应用程序的 Windows.IdentityModel.Extensions.dll](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [用于 64 位应用程序的 Windows.IdentityModel.Extensions.dll](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## <a name="extract-the-json-web-token"></a>提取 JSON Web 令牌


**Decode** 工厂方法将来自 Exchange 服务器的 JWT 分割为构成令牌的三个字符串，然后使用 **Base64Decode** 方法（在第二个示例中显示）将 JWT 标头和负载解码为 JSON 字符串。再将这些字符串传递给 **JsonToken** 构造函数，其中会验证 JWT 的内容并返回一个新 **JsonToken** 对象实例。


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

**Base64Decode** 方法可实现在 [JSON Web 令牌 (JWT) Internet 草稿](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl)的“有关无边距实现 base64url 编码的注意事项”附录中介绍的解码逻辑。




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## <a name="parse-the-jwt"></a>分析 JWT


**JsonToken** 对象的构造函数会检查 JWT 的结构和内容，以确定它是否有效。最好在请求身份验证元数据文档之前执行此操作。如果 JWT 不包含正确声明或它在生命周期之外，则可避免对 Exchange 服务器的调用以及关联的延迟。

构造函数调用实用程序方法来确定不同的声明是否存在并在范围内。如果存在问题，该实用程序方法将引发应用程序异常。如果未引发异常，**IsValid** 属性将设置为 **true**，令牌将准备进行签名验证。

本文稍后将对每个实用程序方法进行更多介绍。




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### <a name="validateheader-method"></a>ValidateHeader 方法

**ValidateHeader** 方法通过检查确保所需声明在令牌标头中，并且声明具有正确的值。该标题必须设置如下；否则，该方法将引发应用程序异常并结束。

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<thumbprint>" }
```

```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### <a name="validatelifetime-method"></a>ValidateLifetime 方法

JWT: "nbf" ("not before") 中提供的两个日期给出了令牌生效的日期和时间，“exp”给出了令牌到期的时间。只有在这两个日期之间提供的令牌应视为有效。为了适应服务器和客户端之间时钟设置的细微区别，此方法将对令牌进行验证，最多验证令牌时间的前五分钟和后五分钟。


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

**validFrom** ("nbf") 和 **validTo** ("exp") 日期作为自 Unix 纪元 1970 年 1 月 1 日以来的秒数发送。使用 UTC 来计算日期和时间，以避免因 Exchange 服务器与运行验证代码的服务器之间的时区差异产生任何问题。


### <a name="validateaudience-method"></a>ValidateAudience 方法

标识令牌只对请求它的外接程序有效。**ValidateAudience** 方法检查令牌中的访问群体声明，以确保它与 Outlook 外接程序的预期 URL 匹配。


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### <a name="validateversion-method"></a>ValidateVersion 方法

**ValidateVersion** 方法检查标识令牌的版本并确保它与预期版本匹配。不同版本的令牌可以携带不同的声明。检查版本可确保预期的声明将在标识令牌中。


```js
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### <a name="validatemetadatalocation-method"></a>ValidateMetadataLocation 方法

在 Exchange 服务器中存储的身份验证元数据对象包含对标识令牌中包括的签名进行验证所需的信息。**ValidateMetadataLocation** 方法确保标识令牌中包含身份验证元数据 URL 声明，并实际验证签名的过程发生在下一步中。


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## <a name="validate-the-identity-token-signature"></a>验证标识令牌签名


在了解到 JWT 包含验证签名所需的声明后，可以使用 Windows Identity Foundation (WIF) 和 WIF 扩展验证令牌中的签名。您需要以下信息来验证签名：


- 从 Exchange 服务器发送的原始 base64 URL 编码的标识令牌字符串。
    
- 来自 JWT 的身份验证元数据文档位置。
    
- 来自 JWT 的访问群体 URL。
    
在此示例中，**IdentityToken** 对象的构造函数从 Exchange 服务器获取身份验证元数据文档并验证标识令牌中的签名。如果标识令牌有效，则可使用 **IdentityToken** 对象实例获取标识令牌中包括的唯一用户 ID。




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

**IdentityToken** 对象构造函数中的大部分代码都使用来自 Exchange 服务器的声明设置实例中的属性。该构造函数调用 **GetSecurityTokenHandler** 方法来获取将验证 Exchange 标识令牌的令牌处理程序。**GetSecurityTokenHandler** 方法调用两个实用程序方法 **GetMetadataDocument** 和 **GetSigningCertificate**，它们用于执行从 Exchange 服务器获取签名证书的工作。以下章节对每种方法进行了介绍。


### <a name="getsecuritytokenhandler-method"></a>GetSecurityTokenHandler 方法

**GetSecurityTokenHandler** 方法返回将验证标识令牌的 WIF 令牌处理程序。该方法中的大部分代码都用于初始化令牌处理程序，以执行验证；但是，该方法会调用 **GetSigningCertificate** 方法，以从 Exchange 服务器中检索用于对令牌签名的 X.509 证书。


```C#
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### <a name="getsigningcertificate-method"></a>GetSigningCertificate 方法

**GetSigningCertificate** 方法调用 **GetMetadataDocument** 方法，以从 Exchange 服务器中检索身份验证元数据，然后返回身份验证元数据文档中的第一个 X.509 证书。如果该文档不存在，该方法将引发应用程序异常。


```C#
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### <a name="getmetadatadocument-method"></a>GetMetadataDocument 方法

身份验证元数据文档包含对 Exchange 标识令牌中的签名进行验证所需的信息。该文档作为 JSON 字符串发送。**GetMetatDataDocument** 方法从 Exchange 标识令牌中指定的位置请求文档并返回将 JSON 字符串封装为对象的对象。如果 URL 不包含身份验证元数据文档，该方法会引发应用程序异常。


```C#
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

默认情况下，Exchange 服务器使用自行签署式 X.509 证书对身份验证元数据文档的请求进行身份验证。除非安装追溯到根服务器的证书，否则必须创建证书验证回调方法，否则对身份验证元数据文档的请求将会失败。 

.NET Framework System.Net 命名空间中的 **ServicePointManager** 类使你可以通过设置 **ServerCertificateValidationCallback** 属性来挂起身份验证回调方法。在[验证 X509 证书](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx)一文中可以看到适合开发和测试的证书验证回调方法的示例。


 **安全说明**  如果使用证书验证回调方法，则必须确保其满足组织的安全要求。


## <a name="compute-the-unique-id-for-an-exchange-account"></a>计算 Exchange 帐户的唯一 ID


你可以通过将身份验证元数据文档 URL 与帐户的 Exchange 标识符进行哈希运算来创建 Exchange 帐户的唯一标识符。在拥有此唯一标识符后，你可以将其用于为 Outlook 外接程序 Web 服务创建单一登录 (SSO) 系统。有关将唯一标识符用于 SSO 的详细信息，请参阅[使用 Exchange 的标识令牌对用户进行身份验证](../outlook/authenticate-a-user-with-an-identity-token.md)

**UniqueUserIdentification** 属性通过使用 **System.Security.Cryptography** 命名空间中的标准 SHA256 提供程序来创建 Exchange ID 和身份验证元数据 URL 的经 salt 加密的 SHA256 哈希值。


 **安全说明**  你必须将身份验证元数据文档与 Exchange ID 进行哈希运算以创建帐户的唯一标识符。仅使用 Exchange ID 会向未经授权的用户公开你的服务。且像往常一样，当处理身份验证和安全性时，你必须确保使用通过此方法创建的唯一标识符满足应用程序的安全性要求。




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## <a name="utility-objects"></a>实用程序对象


本文中的代码示例取决于一些为使用的常量提供友好名称的实用程序对象。下表列出了这些实用程序对象。


**表 1：实用程序对象**


|**对象**|**说明**|
|:-----|:-----|
|**AuthClaimsType**|将令牌验证代码所用的声明标识符收集到一个位置。|
|**Config**|提供用于验证标识令牌的常量。 |
|**JsonAuthMetadataDocument**|封装从 Exchange 服务器发送的 JSON 身份验证元数据文档。|

### <a name="authclaimtypes-object"></a>AuthClaimTypes 对象

**AuthClaimTypes** 对象将令牌验证代码所用的声明标识符收集到一个位置。其中既包括标准 JWT 声明，还包括 Exchange 标识令牌中的特定声明。


```C#
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### <a name="config-object"></a>Config 对象

**Config** 对象包含用于验证标识令牌的常量，以及可在服务器没有追溯到根证书的 X509 证书时使用的证书验证回调方法。


 
  **安全说明**  仅当服务器使用默认的自签名证书时，才需使用安全证书回调方法。当证书为自签名证书时，此示例中的回调方法会返回 **false**，因此你需要将其替换为满足组织安全要求的回调方法。有关适合开发和测试的证书验证回调方法的示例，请参阅[验证 X509 证书](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx)。


```C#
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### <a name="jsonauthmetadatadocument-object"></a>JsonAuthMetadataDocument 对象

**JsonAuthMetadataDocument** 对象通过属性公开身份验证元数据文档的内容。


```C#
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string [] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## <a name="additional-resources"></a>其他资源



- [使用 Exchange 标识令牌对 Outlook 外接程序进行身份验证](../outlook/authentication.md)
    
- [Exchange 标识令牌揭秘](../outlook/inside-the-identity-token.md)
    
