
# LabsJS.Labs

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

LabsJS.Labs 模块包含一系列关键 JavaScript API，您可以使用它们来创建 Office 外接程序（实验室）。API 提供了实验室开发的入口点。

## LabsJS.Labs API 模块

Labs 模块包含以下类型：


### 变量


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|使用此对象可构建默认的 [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) 实例。|

### 函数


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|初始化与主机的连接。|
|[Labs.connect（重载）](../../reference/office-mix/labs.connect-overload.md)|初始化与主机的连接，并提供输入参数。|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|初始化与主机的连接。|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|检索与指定的连接关联的配置信息。|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|将实验室从主机断开，并提供实验室完成状态。|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|打开指定的实验室进行编辑。您可以在编辑模式下指定实验室的配置数据。但是，您无法在使用实验室（即，实验室正在运行）时对其进行编辑。|
|[Labs.takeLab](../../reference/office-mix/labs.takelab.md)|运行指定的实验室并将实验室结果发送到服务器。请注意，在编辑实验室时无法运行它。|
|[Labs.on](../../reference/office-mix/labs.on.md)|为指定事件添加新的处理程序。|
|[Labs.off](../../reference/office-mix/labs.off.md)|移除指定事件的事件处理程序。|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|检索 [Labs.Timeline](../../reference/office-mix/labs.timeline.md) 对象实例，您可以使用它来控制主机播放器控件。|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|将指定 JSON 对象反序列化为一个对象。仅供组件作者使用。|

### 类


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../reference/office-mix/labs.componentinstancebase.md)|用于组件实例初始化的基类。|
|[Labs.ComponentInstance](../../reference/office-mix/labs.componentinstance.md)|表示组件的一个实例，这是运行时对用户的指定组件的实例化。对象包含实验室特定运行的组件的已翻译视图。|
|[Labs.Command](../../reference/office-mix/labs.command.md)|用于在客户端和主机之间传递消息的常规命令。|
|[Labs.LabEditor](../../reference/office-mix/labs.labeditor.md)|**LabEditor** 对象允许您编辑指定实验室，并获取和设置与实验室关联的配置数据。|
|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md)|为当前用户配置的实验室的实例。使用此对象可记录和检索用户的实验室数据。|
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|提供对 labs.js 时间线功能的访问权限。|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|用于保存和跟踪指定实验室的值的容器对象。值可以存储在本地，也可以存储在服务器上。|

### 接口


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|允许您检索与 [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md) 命令关联的数据。|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|允许您定义事件处理程序的接口。|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|提供与 [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx) 对象的交互方式。|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|与 [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx) 命令关联的数据。|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|与采取操作命令关联的数据。|

### 枚举


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|枚举实验室与主机之间的可能的连接状态。|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)|给定实验室的状态值。|
