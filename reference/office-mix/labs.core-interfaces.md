
# <a name="labs.core-interfaces"></a>Labs.Core 接口
**LabsJS.Labs.Core** 模块中的接口

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

**LabsJS.Labs.Core** 模块包含以下接口。

## 


|||
|:-----|:-----|
|[Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)|表示实验室操作，即用户与指定实验室的交互。|
|[Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md)|采取行动的结果。当采取操作时，根据操作的类型，会产生由服务器设置的结果，或者由客户端提供的结果。|
|[Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md)|实验室组件实例的基类。|
|[Labs.Core.IConfigurationInfo](../../reference/office-mix/labs.core.iconfigurationinfo.md)|有关实验室配置的信息。|
|[Labs.Core.IConnectionResponse](../../reference/office-mix/labs.core.iconnectionresponse.md)|从连接调用中返回的响应信息。|
|[Labs.Core.IGetActionOptions](../../reference/office-mix/labs.core.igetactionoptions.md)|作为 **get** 操作的一部分传递的选项。|
|[Labs.Core.ILabCreationOptions](../../reference/office-mix/labs.core.ilabcreationoptions.md)|作为实验室创建操作的一部分传递的选项。|
|[Labs.Core.ILabHostVersionInfo](../../reference/office-mix/labs.core.ilabhostversioninfo.md)|有关实验室主机的版本信息。|
|[Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md)|实验室操作选项的定义。执行给定的操作时传递的选项。|
|[Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)|提供与实验室相关的用户信息。|
|[Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)|[Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) 对象实例，其中包含数值数据（如果有的话）。|
|[Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)|提供实验室版本信息。|
|[Labs.Core.IAnalyticsConfiguration](../../reference/office-mix/labs.core.ianalyticsconfiguration.md)|自定义分析配置信息。允许你指定要加载哪个 IFrame 以显示对用户运行实验室的自定义分析。|
|[Labs.Core.ICompletionStatus](../../reference/office-mix/labs.core.icompletionstatus.md)|实验室的完成状态。传递实验室完成状态，用于指示交互结果。|
|[Labs.Core.ILabCallback](../../reference/office-mix/labs.core.ilabcallback.md)|用于处理 Labs.js 回调方法的接口。|
|[Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)|与实验室关联的对象。该对象包含某一类型字段，指示它是哪种类型的对象。|
|[Labs.Core.ITimelineConfiguration](../../reference/office-mix/labs.core.itimelineconfiguration.md)|用于 [Labs.Timeline](../../reference/office-mix/labs.timeline.md) 的配置选项。允许指定一组时间线配置选项。|
|[Labs.Core.IUserData](../../reference/office-mix/labs.core.iuserdata.md)|用于表示存储在对象上的自定义用户数据的基接口。|
|[Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md)|存储在实验室中的值的基类。|
|[Labs.Core.IConfiguration](../../reference/office-mix/labs.core.iconfiguration.md)|实验室配置数据结构。|
|[Labs.Core.IConfigurationInstance](../../reference/office-mix/labs.core.iconfigurationinstance.md)|实验室配置实例的基类。|
|[Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)|表示实验室组件的基类。|
|[Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)|提供用于将 Labs.js 连接到主机的抽象层。|
|[Labs.Core.ModeChangedEventData](../../reference/office-mix/labs.core.modechangedeventdata.md)|与模式更改事件关联的数据。|
|[Labs.Core.IEventCallback](../../reference/office-mix/labs.core.ieventcallback.md)|用于处理 EventManager 回调的接口。|
