# Hosts 元素

指定 Office 客户端应用程序中 Office 外接程序将激活的位置。 包含 **Host** 元素及其设置的集合。 

当该元素被包括在 [VersionOverrides](./versionoverrides.md) 节点中时，它将替代清单中父级部分中的 **Hosts** 元素。 

## 子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Host](#host)    |  是   |  描述主机及其设置。 |

> **注意：** Outlook 需要 `Hosts` 包含 `MailHost` 的 `Host` 定义。

---- 

## Host 元素
指定外接程序应该激活的单个 Office 应用程序类型，例如“文档”、“工作簿”、“演示文稿”、“项目”、“邮箱”和“笔记本”。

### 属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  是  | 描述这些设置所适应的 Office 主机。|

### 子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  是   |  定义受影响的外形规则。 |


### xsi:type
控制所包含的设置也适用的 Office 主机类别（Word、Excel、PowerPoint、Outlook 和 OneNote）。 值必须为以下值之一：

- `MailHost` (Outlook)    


## 主机示例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
