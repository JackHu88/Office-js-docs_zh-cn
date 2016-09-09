
# Labs.Core.ILabHost

 _**适用范围：** Office 相关应用 | Office 外接程序 | Office Mix | PowerPoint_

提供用于将 Labs.js 连接到主机的抽象层。

```
interface ILabHost
```


## 方法


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

检索由实验室主机所支持的版本。

 **参数**

无。


### connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

初始化与主机的连接。

 **参数**


|||
|:-----|:-----|
| _版本_|客户端可利用的主机版本的列表。|
| _callback_|完成连接后，将触发的回调函数。|

### disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

终止与主机的通信。

 **参数**


|||
|:-----|:-----|
| _completionStatus_|断开连接时，实验室的状态。|
| _callback_|断开连接后，将触发的回调函数。|

### on

 `on(handler: (string: any, any: any): void)`

添加事件处理程序，以处理来自主机的邮件。已解决的承诺将返回到主机。

 **参数**


|||
|:-----|:-----|
| _Handler_|事件处理程序。|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

向主机发送一条消息。

 **参数**


|||
|:-----|:-----|
| _type_|正在发送的邮件类型。|
| _options_|邮件选项。|
| _callback_|接收到消息后触发的回调函数。|

### create

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

创建实验室。存储主机信息，并留出空间来存储配置和其他元素。

 **参数**


|||
|:-----|:-----|
| _options_|作为创建操作的一部分传递的选项。|
| _callback_|创建实验室后触发的回调函数。|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

从主机检索当前实验室配置。

 **参数**


|||
|:-----|:-----|
| _callback_|要检索配置信息的回调函数。|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

在主机上设置一个新的实验室配置。

 **参数**


|||
|:-----|:-----|
| _configuration_|设置的实验室配置。|
| _callback_|设置配置后触发的回调函数。|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

检索实验室的实例配置。

 **参数**


|||
|:-----|:-----|
| _callback_|检索完配置实例后触发的回调函数。|

### getState

 `getState(callback: Core.ILabCallback<any>)`

针对给定用户检索实验室的当前状态。

 **参数**


|||
|:-----|:-----|
| _completionStatus_|返回当前实验室状态的回调函数。|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

针对给定用户设置实验室的状态。

 **参数**


|||
|:-----|:-----|
| _state_|实验室的状态。|
| _callback_|设置状态后触发的回调函数。|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

对某个操作采取尝试。

 **参数**


|||
|:-----|:-----|
| _type_|操作类型。|
| _options_|与操作一同提供的选项。|
| _callback_|返回最终执行的操作的回调函数。|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

采取已完成的操作。

 **参数**


|||
|:-----|:-----|
| _type_|操作类型。|
| _options_|与操作一同提供的选项。|
| _result_|操作的结果。|
| _callback_|返回最终执行的操作的回调函数。|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

对某个操作采取尝试。

 **参数**


|||
|:-----|:-----|
| _type_|get 操作类型。|
| _options_|与 get 操作一同提供的选项。|
| _callback_|返回已完成操作的列表的回调函数。|
