
# <a name="labs.imessagehandler"></a>Labs.IMessageHandler

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

允许定义事件处理程序的接口。

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **参数**


|||
|:-----|:-----|
| `origin`|生成消息的实验室窗口。|
| `data`|消息的内容。|
| `callback`|接收到消息后触发的回调函数。|
