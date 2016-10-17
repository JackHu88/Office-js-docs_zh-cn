# <a name="formfactor-element"></a>FormFactor 元素

指定对给定外形规格的外接程序的设置。例如，将 `Host` 定义为类型 `MailHost` 和 `DesktopFormFactor` 适用于 Outlook for Desktop 但是_不_适用于 Outlook Web App 或 Outlook.com。它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。

每个 FormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](./functionfile.md) 和 [ExtensionPoint 元素](./extensionpoint.md)。 

支持下列 FormFactor：

- `DesktopFormFactor`（适用于 Windows 或 Mac 客户端的 Office）

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | 说明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](./functionfile.md)     | 是      | 包含 JavaScript 函数的文件的 URL。|
| [GetStarted](./getstarted.md)         | 否       | 定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。 |

## <a name="formfactor-example"></a>FormFactor 示例

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
