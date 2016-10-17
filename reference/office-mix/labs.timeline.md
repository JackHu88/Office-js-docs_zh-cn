
# <a name="labs.timeline"></a>Labs.Timeline

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

提供对 labs.js 时间线功能的访问权限。

```
class Timeline
```


## <a name="methods"></a>方法




### <a name="method"></a>方法

 `function constructor(labsInternal: Labs.LabsInternal)`

创建 **Timeline** 类的新实例。


### <a name="next"></a>下一页

 `public function next(completionStatus: Labs.Core.ICompletionStatus, callback: Labs.Core.ILabCallback<void>): void`

表示时间线应前进到下一张幻灯片。

 **参数**


|||
|:-----|:-----|
| _completionStatus_|表示实验室的当前状态。|
| _callback_|实验室已移动到下一张幻灯片时触发的回调函数。|
