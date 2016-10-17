
# <a name="labs.labeditor"></a>Labs.LabEditor

 _**适用范围：**Office 相关应用? | Office 外接程序? | Office Mix? | PowerPoint_

**LabEditor** 对象允许编辑指定实验室，并获取和设置与实验室关联的配置数据。

```
class LabEditor
```


## <a name="methods"></a>方法


### <a name="getconfiguration"></a>getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

检索当前实验室配置。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _callback_|检索完配置后触发的回调函数。|

### <a name="setconfiguration"></a>setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

设置一个新的实验室配置。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _configuration_|要设置的配置。|
| _callback_|设置完配置后触发的回调函数。|

### <a name="done"></a>done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

指示用户已完成编辑实验室的操作。

 **参数**


|**名称**|**Description**|
|:-----|:-----|
| _callback_|完成实验室编辑器操作后触发的回调函数。|
