
# <a name="labs.command"></a>Labs.Command

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

用于在客户端和主机之间传递消息的常规命令。

```
class Command
```


## <a name="properties"></a>属性


|**名称**|**Description**|
|:-----|:-----|
| `public var type: string`|命令的类型。|
| `public var commandData: any`|与命令关联的可选数据。|

## <a name="methods"></a>方法




### <a name="constructor"></a>构造函数

 `function constructor(type: string, commandData?: any)`

说明

 **参数**


|||
|:-----|:-----|
| `type`|命令的类型。|
| `commandData`|与命令关联的可选数据。|
