
# <a name="labs.core.iconnectionresponse"></a>Labs.Core.IConnectionResponse

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

从连接调用中返回的响应信息。

```
interface IConnectionResponse
```


## <a name="properties"></a>属性


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|初始化配置信息，或如果应用尚未初始化，则返回 **null**。|
| `mode: Core.LabMode`|实验室目前正在运行的模式。|
| `hostVersion: Core.IVersion`|服务器的版本信息 ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md))。|
| `userInfo: Core.IUserInfo`|有关用户的信息 ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md))。|
