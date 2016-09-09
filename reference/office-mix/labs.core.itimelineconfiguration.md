
# Labs.Core.ITimelineConfiguration

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

用于 [Labs.Timeline](../../reference/office-mix/labs.timeline.md) 的配置选项。允许您指定一组时间线配置选项。

```
interface ITimelineConfiguration
```


## 属性


|||
|:-----|:-----|
| `duration: number`|实验室的持续时间，以秒为单位。|
| `capabilities: string[]`|实验室支持的时间线功能的数组列表，例如，播放、暂停、寻道等等。|
