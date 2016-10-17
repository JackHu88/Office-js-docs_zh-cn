
# <a name="labs.connect-(overload)"></a>Labs.connect（重载）

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

初始化与主机的连接。

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## <a name="parameters"></a>参数


|||
|:-----|:-----|
| _labHost_|可选。要连接到的 [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) 实例。如果未指定主机，将使用 [Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md) 构造一个主机。|
| _callback_|建立连接之后触发的回调。|

## <a name="return-value"></a>返回值

返回到主机的连接。

