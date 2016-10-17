
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a>使用 Exchange 的标识令牌对用户进行身份验证

您可以为信息服务实现单一登录 (SSO) 身份验证方案，从而让使用 Outlook 外接程序的客户能够通过其 Exchange 服务器凭据连接到您的服务。本文演示如何使用简单的基于  **Dictionary** 对象的用户数据存储匹配凭据。

 >**注释**  这只是 SSO 的一个简单示例，不应在生产代码中使用。和以往一样，在处理标识和身份验证事宜时，一定要确保您的代码满足组织的安全要求。


## <a name="prerequisites-for-using-sso-authentication"></a>使用 SSO 身份验证的先决条件


若要将标识令牌用于 SSO，您的服务应用程序需要具有有效的标识令牌。在以下文章中，您可以了解标识令牌，以及如何请求和验证标识令牌：


- [Exchange 标识令牌揭秘](../outlook/inside-the-identity-token.md)
    
- [在 Exchange 中使用标识令牌从 Outlook 外接程序调用服务](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [使用 Exchange 令牌验证库](../outlook/use-the-token-validation-library.md)（如果使用的是托管代码），或 [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)（如果您要编写自己的令牌验证方法）。
    

## <a name="authenticate-a-user"></a>对用户进行身份验证


以下代码示例演示一个简单身份验证对象，该对象将匹配包含一组服务凭据的标识令牌所代表的唯一标识符。 **TokenAuthentication** 类提供了方法 **GetResponseFromService** ，后者将返回先前经过身份验证的令牌的响应，或者要求用户提供可以进行身份验证并与标识令牌关联的凭据。代码并不完整；它假设您将提供以下对象和方法。



|**对象/方法**|**说明**|
|:-----|:-----|
|**LocalCredentials** 对象|代表您的服务的用户凭据。对象结构取决于服务的要求。|
|**IdentityToken** 对象|包含 Outlook 外接程序发送到您的服务的用户标识令牌。该对象必须至少包含用户的唯一 Exchange 标识符和颁发该令牌的服务器的身份验证元数据 URL。此示例使用 [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)一文中定义的标识令牌对象。|
|**JsonResponse** 对象|代表您的服务的响应。该对象可序列化为 JSON 对象。|
|**CallService** 方法|通过包含用户的服务凭据的  **LocalCredentials** 对象以及包含服务请求数据的对象调用您的服务。如果凭据有效，此方法将返回包含请求结果的 **JsonReponse** 对象。如果凭据无效，此方法将返回 **null**。|
|**GetCredentialsResponse** 方法|返回您的 Office 邮件加载项将识别为服务凭据的请求的  **JsonReponse** 对象。|
|**LocalCredentialsAreValid** 方法|如果提供给服务的凭据有效，将返回  **true**；否则，将返回  **false**。|

 >**注释**  这只是有关如何使用标识令牌的一个建议。和以往一样，在处理标识和身份验证事宜时，一定要确保您的代码满足组织的安全要求。


```C#
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## <a name="authenticating-a-user-with-the-managed-validation-library"></a>使用托管验证库对用户进行身份验证


如果要使用托管库验证标识令牌，您无需计算唯一键。 **AppIdentityToken** 类的 **UniqueUserIdentification** 属性可直接用作用户的唯一键。以下代码示例演示为使用 **AppIdentityToken** 类而需要对前面示例中的 **GetResponseFromService** 方法所进行的修改。


```js
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## <a name="additional-resources"></a>其他资源



- [使用 Exchange 标识令牌对 Outlook 外接程序进行身份验证](../outlook/authentication.md)
    
- [在 Exchange 中使用标识令牌从 Outlook 外接程序调用服务](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [使用 Exchange 令牌验证库](../outlook/use-the-token-validation-library.md)
    
- [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)
    
