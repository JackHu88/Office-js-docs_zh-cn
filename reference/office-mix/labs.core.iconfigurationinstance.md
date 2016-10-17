
# <a name="labs.core.iconfigurationinstance"></a>Labs.Core.IConfigurationInstance

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

实验室配置实例的基类。实例是对给定用户配置的实例化，包含有关实验室的特定运行配置的翻译视图。此视图可能排除了隐藏的信息（例如，提示和答案），还包含用来标识各种实例的 ID。

```
interface IConfigurationInstance extends Core.IUserData
```


## <a name="properties"></a>属性


|||
|:-----|:-----|
| `appVersion: Core.IVersion`|与此配置关联的实验室版本。|
| `components: Core.IComponentInstance[]`|与实验室相关联的组件。|
| `name: string`|实验室的名称。|
| `timeline: Core.ITimelineConfiguration`|实验室的时间线配置。|
