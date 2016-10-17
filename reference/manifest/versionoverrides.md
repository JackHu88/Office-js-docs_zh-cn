# <a name="versionoverrides-element"></a>VersionOverrides 元素

根元素包含由外接程序实现的外接程序命令的信息。**VersionOverrides** 是清单中 [OfficeApp](./officeapp.md) 元素的子元素。该元素在清单架构 v1.1 及更高版本中受支持，但在 VersionOverrides v1.0 架构中定义。 

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **xmlns**       |  是  |  架构位置必须是 `http://schemas.microsoft.com/office/mailappversionoverrides`|
|  **xsi:type**  |  是  | 架构版本目前唯一的有效值是 `VersionOverridesV1_0`。 |


## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Description**    |  否   |  描述外接程序。这会替代清单中任何父级部分中的 `Description` 元素。说明文本包含在 [Rescources](./resources.md) 元素中的 **LongString** 元素的子元素中。**Description** 元素的 `resid` 属性被设置为包含文本的 `String` 元素的 `id` 属性的值。|
|  **Requirements**  |  否   |  指定外接程序要求的最低要求集和 Office.js 的版本。这会替代清单中父级部分中的 `Requirements` 元素。| 
|  [Hosts](./hosts.md)                |  是  |  指定 Office 主机的集合。子级 Hosts 元素替代清单中父级部分中的 Hosts 元素。  |
|  [Resources](./resources.md)    |  是  | 定义其他清单元素引用的资源集合（字符串、URL 和图像）。|



### <a name="versionoverrides-example"></a>VersionOverrides 示例
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information on resources -->
   </Resources>
</VersionOverrides>
...
</OfficeApp>
```
