
# Labs.Core.IConfiguration

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

实验室配置数据结构。

```
interface IConfiguration extends Core.IUserData
```


## 属性


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|与此配置关联的应用程序版本。|
| `components: Core.IComponent[]`|实验室中包含的组件。|
| `name: string`|实验室的名称。|
| `timeline: Core.ITimelineConfiguration`|实验室的时间线配置。|
| `analytics: Core.IAnalyticsConfiguration`|实验室的分析配置。|
