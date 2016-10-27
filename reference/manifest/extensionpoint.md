# <a name="extensionpoint-element"></a>ExtensionPoint 元素

 定义 Office UI 中外接程序公开功能的位置。**ExtensionPoint** 元素是 [FormFactor](./formfactor.md) 的子元素。 

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xsi:type**  |  是  | 定义的扩展点类型。|


## <a name="extension-points-for-word,-excel,-powerpoint,-and-onenote-add-in-commands"></a>适用于 Word、Excel、PowerPoint 和 OneNote 外接程序命令的扩展点

- **PrimaryCommandSurface** - Office 中的功能区。
- **ContextMenu** - Office UI 中右键单击时出现的快捷菜单。

下面的示例演示如何将  **ExtensionPoint** 元素与 **PrimaryCommandSurface** 和 **ContextMenu** 属性值配合使用，以及应彼此配合使用的子元素。


 >**重要信息**  对于包含 ID 属性的元素，请确保提供唯一 ID。我们建议您将您的公司名称与您的 ID 配合使用。例如，使用以下格式。<CustomTab id="mycompanyname.mygroupname">


```XML
 <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="Contoso Tab">
            <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
             <!-- <OfficeTab id="TabData"> -->
              <Label resid="residLabel4" />
              <Group id="Group1Id12">
                <Label resid="residLabel4" />
                <Icon>
                  <bt:Image size="16" resid="icon1_32x32" />
                  <bt:Image size="32" resid="icon1_32x32" />
                  <bt:Image size="80" resid="icon1_32x32" />
                </Icon>
                <Tooltip resid="residToolTip" />
                <Control xsi:type="Button" id="Button1Id1">

                   <!-- information about the control -->
                </Control>
                <!-- other controls, as needed -->
              </Group>
            </CustomTab>
          </ExtensionPoint>

        <ExtensionPoint xsi:type="ContextMenu">
          <OfficeMenu id="ContextMenuCell">
            <Control xsi:type="Menu" id="ContextMenu2">
                   <!-- information about the control -->
            </Control>
           <!-- other controls, as needed -->
          </OfficeMenu>
         </ExtensionPoint>
```

**子元素**
 
|**Element**|**说明**|
|:-----|:-----|
|**CustomTab**|如果想要（使用 **PrimaryCommandSurface**）向功能区添加自定义选项卡，则为必需项。如果使用 **CustomTab** 元素，则不能使用 **OfficeTab** 元素。**id** 属性是必需的。|
|**OfficeTab**|如果想要（使用 **PrimaryCommandSurface**）扩展默认 Office 功能区选项卡，则为必需项。如果使用 **OfficeTab** 元素，则不能使用 **CustomTab** 元素。有关详细信息，请参阅 [OfficeTab](officetab.md)。|
|**OfficeMenu**|如果正（使用 **ContextMenu**）将外接程序命令添加到默认上下文菜单中，则为必需项。**id** 属性必须设置为： <br/>适用于 Excel 或 Word 的  - **ContextMenuText**当用户选定文本，然后右键单击所选定的文本时显示上下文菜单上的项。 <br/>适用于 Excel 的  - **ContextMenuCell**当用户右键单击电子表格中的某个单元格时显示上下文菜单上的项。|
|**Group**|选项卡上的一组用户界面扩展点。一个组可以有最多六个控件。 **id** 属性是必需项。它是最多使用 125 个字符的字符串。|
|**Label**|必需。组标签。**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **ShortStrings** 元素的子元素，而 ShortStrings 元素是 **Resources** 元素的子元素。|
|**Icon**|必需。指定将在小型设备上使用或在显示过多按钮的情况下使用的组图标。**resid** 属性必须设置为 **Image** 元素的 **id** 属性的值。**Image** 元素是 **Images** 元素的子元素，而 Images 元素是 **Resources** 元素的子元素。**size** 属性给出图像的大小（以像素为单位）。要求三种图像大小：16、32 和 80。也同样支持五种可选大小：20、24、40、48 和 64。|
|**Tooltip**|可选。组的工具提示**resid** 属性必须设置为 **String** 元素的 **id** 属性的值。**String** 元素是 **LongStrings** 元素的子元素，而 LongStrings 元素是 **Resources** 元素的子元素。|
|**Control**|每个组都要求至少有一个控件。 **Control** 元素可以是 **Button** 或 **Menu**。使用  **Menu** 可指定按钮控件的下拉列表。当前，仅支持按钮和菜单。请参阅 [按钮控件](#button-controls)和 [菜单控件](#menu-controls)各节了解详细信息。<br/>**注意** 为了使故障排除变得更简单，我们建议一次性添加 **Control** 元素和相关的 **Resources** 子元素。

## <a name="extension-points-for-outlook-add-in-commands"></a>Outlook 外接程序命令的扩展点

- [CustomPane](#custompane) 
- [MessageReadCommandSurface](#messagereadcommandsurface) 
- [MessageComposeCommandSurface](#messagecomposecommandsurface) 
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface) 
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module)（仅能在 [DesktopFormFactor](./formfactor.md) 中使用。）

### <a name="custompane"></a>CustomPane

CustomPane 扩展点在满足指定规则时将定义激活的外接程序。仅适用于阅读窗体，并显示水平窗格。 

**子元素**

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **RequestedHeight** | 否 |  显示窗格在桌面计算机上运行时所请求的高度（以像素为单位）。请求的高度可以是 32 到 450 像素。  |
|  **SourceLocation**  | 是 |  外接程序的源代码文件的 URL。这是指 [Resources](./resources.md) 元素中的 **Url** 元素。  |
|  **Rule**  | 是 |  指定外接程序激活时间的规则或规则集。有关详细信息，请参阅 [Outlook 外接程序的激活规则](../../docs/outlook/manifests/activation-rules.md)。 |
|  **DisableEntityHighlighting**  | 否 |  指定是否应关闭实体突出显示。 |


#### <a name="custompane-example"></a>CustomPane 示例
```xml
<ExtensionPoint xsi:type="CustomPane">
   <RequestedHeight>100< /RequestedHeight> 
   <SourceLocation resid="residReadTaskpaneUrl"/>
   <Rule xsi:type="RuleCollection" Mode="Or">
     <Rule xsi:type="ItemIs" ItemType="Message"/>
     <Rule xsi:type="ItemHasAttachment"/>
     <Rule xsi:type="ItemHasKnownEntity" EntityType="Address"/>
   </Rule>
</ExtensionPoint>
```

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface
此扩展点将按钮置于邮件阅读视图的命令界面中。在 Outlook 桌面，它显示在功能区中。

**子元素**

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](./customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface
此扩展点将按钮置于使用电子邮件撰写窗体的外接程序的功能区上。 

**子元素**

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](./customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```
### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

此扩展点将按钮置于向会议的组织者显示的窗体的功能区上。 

**子元素**

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](./customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

此扩展点将按钮置于向会议与会者显示的窗体的功能区上。 

**子元素**

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](./customtab.md) |  将命令添加到自定义功能区选项卡。  |

#### <a name="officetab-example"></a>OfficeTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>CustomTab 示例
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

此扩展点将按钮置于模块扩展的功能区上。 

**子元素**

|  元素 |  说明  |
|:-----|:-----|
|  [OfficeTab](./officetab.md) |  将命令添加到默认功能区选项卡。  |
|  [CustomTab](./customtab.md) |  将命令添加到自定义功能区选项卡。  |

