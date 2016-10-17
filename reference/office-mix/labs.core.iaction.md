
# <a name="labs.core.iaction"></a>Labs.Core.IAction

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

表示实验室操作，即用户与指定实验室的交互。

```
interface IAction
```


## <a name="properties"></a>属性


|||
|:-----|:-----|
| `type: string`|用户所执行的操作类型。|
| `options: Core.IActionOptions`|与用户所执行的操作一同发送的 [Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md) 选项。|
| `result: Core.IActionResult`|该操作的 [Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md) 结果。|
| `time: number`|操作完成的时间，表示自 1970 年 1 月 1 日 00:00:00 UTC 之后经过的以毫秒为单位的时间。|
