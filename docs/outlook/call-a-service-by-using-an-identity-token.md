
# <a name="call-a-service-from-an-outlook-add-in-by-using-an-identity-token-in-exchange"></a>在 Exchange 中使用标识令牌从 Outlook 外接程序调用服务

标识令牌为您的每个客户提供了一个唯一标识符，可使用该标识符对您提供的服务进行个性化设置。您的代码可使用将字符串返回您的 Outlook 外接程序的异步方法调用来要求 Exchange 服务器提供标识令牌。该字符串包含 JSON Web Token (JWT) 标识令牌。您的外接程序无需解压缩令牌。相反，它可将令牌传递到 Web 服务上，以便您的服务能够对来自外接程序的请求进行身份验证。

支持您的外接程序的 Web 服务必须在承载外接程序 HTML 和 JavaScript 源文件的服务器上运行。这将阻止出现跨网站脚本错误。您的服务器可让其他 Web 服务来代理请求（如果您的应用程序需要）。

向您的外接程序发送的服务请求添加标识令牌非常简单。您请求令牌，再使用令牌，然后使用 Web 服务响应。以下是借助您使用 **XmlHttpRequest** 方法发送到服务器的简单 XML 文档定义其外观的方式。

## <a name="request-a-token-from-your-exchange-server"></a>从 Exchange 服务器请求令牌


邮件外接程序的此简单初始化方法使用  **getUserIdentityTokenAsync** 方法从 Exchange 服务器请求标识令牌。 _getUserIdentityToken_ 参数是在返回对服务器的异步请求时调用的函数。请参阅下一个步骤以了解回调方法。


```js
var _mailbox;
var _xhr;
// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
        _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}

```


## <a name="use-the-identity-token"></a>使用标识令牌


针对  **getUserIdentityTokenAsync** 方法的回调函数具有一个参数，该参数将用户标识令牌包含在其 **value** 属性中。

此回调函数创建一个  **XMLHttpRequest** 对象以调用 Web 服务。将 **XMLHttpRequest** 对象的 **onreadystatechange** 属性设置为在您的外接程序获取来自 Web 服务的响应时将运行的函数的名称。




```js
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}
```


## <a name="use-the-web-service-response"></a>使用 Web 服务响应


这是另一个用于处理来自 Web 服务的响应的简单函数。它遵循  **XHMHttpResponse** 回调函数的标准模式。此函数等到来自 Web 服务的完整响应传入后，将该响应的内容置入加载项 UI 中。此函数所解析的响应是来自 Web 服务的响应。有关此响应的详细信息，请参阅 [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)。 


```js
function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


## <a name="example:-calling-a-web-service-with-identity-tokens"></a>示例：使用标识令牌调用 Web 服务


标识令牌向您的服务器上运行的 Web 服务提供有关调用服务的客户端的标识信息。若要使用标识令牌，您需要：


- 一个从 Exchange 服务器请求标识令牌并将其发送到 Web 服务的 Outlook 外接程序。本主题中的信息将帮助您创建该外接程序。
    
- 一个在您的服务器上运行的 Web 服务，此服务为验证标识令牌的外接程序提供了 UI。您将在下列主题之一中找到创建 Web 服务所需的信息：
    
      - [使用 Exchange 令牌验证库](../outlook/use-the-token-validation-library.md) - 如果你使用的是我们提供的验证库。
    
  - [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md) - 如果你编写的是你自己的验证代码。
    

### <a name="code-for-the-sample-add-in"></a>示例外界程序的代码


本文描述的外接程序需要以下文件：


- IdentityTest.js – 为外接程序提供业务逻辑的 JavaScript 文件。
    
- IdentityTest.html – 为外接程序提供 UI 的 HTML 文件。
    
您还需要标识测试 Web 服务。有关该 Web 服务的信息，请参阅 [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)。


#### <a name="identitytest.js"></a>IdentityTest.js

以下示例演示了 IdentityTest.js 文件。


```js
var _mailbox;
var _xhr;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.getUserIdentityTokenAsync(getUserIdentityTokenCallback);
    });
}
function getUserIdentityTokenCallback(asyncResult) {
    var token = asyncResult.value;

    _xhr = new XMLHttpRequest();
    _xhr.open("POST", "https://localhost:44300/IdentityTestService/UnpackTokenJSON");
    _xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    _xhr.onreadystatechange = readyStateChange;

    var request = new Object();
    request.token = token;
    request.phoneNumbers = _mailbox.item.getEntities().phoneNumbers;

    _xhr.send(JSON.stringify(request));
}

function readyStateChange() {
    if (_xhr.readyState == 4 &amp;&amp; _xhr.status == 200) {

        var response = JSON.parse(_xhr.responseText);

        if (undefined == response.error) {
            document.getElementById("msexchuid").value = response.token.msexchuid;
            document.getElementById("amurl").value = response.token.amurl;
            document.getElementById("uniqueID").value = response.token.uniqueID;
            document.getElementById("iss").value = response.token.iss;
            document.getElementById("x5t").value = response.token.x5t;
            document.getElementById("nbf").value = response.token.nbf;
            document.getElementById("exp").value = response.token.exp;
        }
        else {
            document.getElementById("error").value = response.error;
        }
    }
}
```


#### <a name="identitytest.html"></a>IdentityTest.html

以下示例演示了 IdentityTest.html 文件。


```HTML
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Identity Test</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <script src="../Scripts/jquery-1.6.2.js"></script>
    <script src="../Scripts/Office/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/Office.js"></script>

    <!-- Add your JavaScript to the following JavaScript file -->
    <script src="../Scripts/IdentityTest.js"></script>
</head>
<body>
    <div id="SectionContent">
        <table style="width: 80%;">
            <tr>
                <th>Claim
                </th>
                <th>Contents
                </th>
            </tr>
            <tr>
                <td style="width: 25%;">Error:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="error" value="None" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">User Exchange ID:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="msexchuid" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Authentication Metadata URL:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="amurl" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Unique identifier:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="uniqueID" />
                </td>
            </tr>
          </tr>
            <tr>
                <td style="width: 25%;">Audience:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="aud" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Issuer:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="iss" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Certificate thumbprint:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="x5t" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid from:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="nbf" />
                </td>
            </tr>
            <tr>
                <td style="width: 25%;">Valid to:
                </td>
                <td style="width: 75%;">
                    <input style="width: 100%;" id="exp" />
                </td>
            </tr>
        </table>
    </div>
</body>
</html>
```


## <a name="next-steps"></a>后续步骤


在了解如何请求标识令牌后，您需要在请求的服务器端使用令牌。下面的文章将帮助您快速入门：


- [使用 Exchange 令牌验证库](../outlook/use-the-token-validation-library.md)
    
- [验证 Exchange 标识令牌](../outlook/validate-an-identity-token.md)
    
- [使用 Exchange 的标识令牌对用户进行身份验证](../outlook/authenticate-a-user-with-an-identity-token.md)
    

## <a name="additional-resources"></a>其他资源



- [使用 Exchange 标识令牌对 Outlook 外接程序进行身份验证](../outlook/authentication.md)
    
- [Exchange 标识令牌揭秘](../outlook/inside-the-identity-token.md)
    
