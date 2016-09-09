
# LabsJS JavaScript API 引用
大致了解 LabsJS JavaScript对象模型。

 _**适用范围：** Office 相关应用程序 | Office 外接程序 | Office Mix | PowerPoint_

LabsJS 引用中记录了 [TypeScript](http://www.typescriptlang.org/) 文件 **labs-1.0.42.d.ts** ，该文件将 LabsJS 对象模型分类成模块。

## LabsJS 对象模型

LabsJS 对象模型分为五个模块：


- [LabsJS.Labs](../../reference/office-mix/labsjs.labs.md)。实验室模块包含关键 API 集，您可以使用其创建实验室本身。它们是实验室开发的入口点。
    
- [LabsJS.Labs.Core](../../reference/office-mix/labsjs.labs.core.md)。LabsJS 和演示文稿驱动程序（在这种情况下为 Office Mix）共用的核心界面、数据结构和类，可在两者之间架起一座桥梁。
    
- [LabsJS.Labs.Core.Actions](../../reference/office-mix/labsjs.labs.core.actions.md)。这些 API 表示指明实验室当前行为的操作，对创建新组件（而非默认组件）或使用新驱动程序（而非 Office Mix）建立连接的开发人员非常有用。
    
- [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md)。这些 API 允许您查询服务器上之前发生的操作。
    
- [LabsJS.Labs.Components](../../reference/office-mix/labsjs.labs.components.md)。这些 API 表示目前对实验室可用的四个默认组件（活动、选择、输入和动态）。
    
每个模块都包含一系列成员，其中包含下列成员中的一个或多个：


- 类
    
- 接口
    
- 函数
    
- 枚举
    
- 变量
    



## 其他资源



- [TypeScript](http://www.typescriptlang.org/)
    
