# <a name="action-element"></a>Action 元素
 指定用户选择 [按钮](./control.md#button-control)或[菜单](./control.md#menu-dropdown-button-controls)控件时将执行的操作。
 
## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 要执行的操作类型|


## <a name="child-elements"></a>子元素

|  元素 |  说明  |
|:-----|:-----|
|  [FunctionName](#functionname) |    指定要执行的函数的名称。 |
|  [SourceLocation](#sourcelocation) |    指定该操作的源文件位置。 |
|  [TaskpaneId](#taskpaneid) | 指定任务窗格容器的 ID。|
  

## <a name="xsi:type"></a>xsi:type
此属性指定当用户选择按钮时所执行的操作类型。它可以是下列值之一：
- ExecuteFunction
- ShowTaskpane

## <a name="functionname"></a>FunctionName

**xsi:type** 为“ExecuteFunction”时的必需元素。指定要执行的函数的名称。函数包含在 [FunctionFile](./functionfile.md) 元素指定的文件中。

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a>SourceLocation
**xsi:type** 为 ShowTaskpane 时的必需元素。指定此操作的源文件位置。 **resid** 属性必须设置为 **Urls** 元素（位于 **Resources** 元素）中 [Url](./resources.md#urls) 元素的 [id](./resources.md) 属性的值。

```xml
 <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
```  

## <a name="taskpaneid"></a>TaskpaneId
可选元素，当 **xsi: type** 是“ShowTaskpane”时。指定任务窗格容器的 ID。具有多个“ShowTaskpane”操作时，如果想要对每个操作使用独立的窗格，则使用不同的 **TaskpaneId**。为共享相同窗格的不同操作使用同一 **TaskpaneId**当用户选择共享同一 **TaskpaneId** 的命令时，窗格容器将保持打开状态，但窗格的内容将被替换为相应的操作“SourceLocation” 

>**注意：**在 Outlook 中不支持此元素。

以下示例介绍了共享相同 **TaskpaneId** 的两个操作。 


```xml
 <Action xsi:type="ShowTaskpane">
    <TaskpaneId>MyPane</TaskpaneId>
    <SourceLocation resid="aTaskPaneUrl" />
  </Action>

  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>MyPane</TaskpaneId>
    <SourceLocation resid="anotherTaskPaneUrl" />
  </Action>
```  
