
# ProjectViewTypes 枚举
指定 **[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** 方法可以识别的视图的类型。

|||
|:-----|:-----|
|**主机：**|Project|
|**在其中添加**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## 成员


****


|**成员**|**说明**|
|:-----|:-----|
|**Gantt**|甘特图视图|
|**NetworkDiagram**|网络图视图。|
|**TaskDiagram**|任务图视图。|
|**TaskForm**|任务表单视图。|
|**TaskSheet**|任务表视图。|
|**ResourceForm**|资源表视图。|
|**ResourceSheet**|资源表视图。|
|**ResourceForm**|资源表视图。|
|**ResourceGraph**|资源图视图。|
|**TeamPlanner**|工作组规划器视图。|
|**TaskDetails**|任务详细信息视图。|
|**TaskNameForm**|任务名称表单视图。|
|**ResourceNames**|资源名称视图。|
|**日历**|日历视图。|
|**TaskUsage**|任务使用视图。|
|**ResourceUsage**|资源使用视图。|
|**时间线**|时间线视图。|

## 备注

**[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** 方法返回与活动视图对应的 **ProjectViewTypes** 常数值和名称。


## 支持详细信息


下列矩阵中的大写字母 Y 表示相应的 Office 主机应用程序支持此枚举。空的单元格表示相应的 Office 主机应用程序不支持此枚举。

有关 Office 主机应用程序和服务器要求的详细信息，请参阅[运行 Office 外接程序的要求](../../docs/overview/requirements-for-running-office-add-ins.md)。


**支持的主机（按平台）**


||**Office for Windows Desktop**|**Office Online（在浏览器中）**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**应用程序类型**|任务窗格|
|**库**|Office.js|
|**命名空间**|Office|

## 支持历史记录



****


|**版本**|**更改内容**|
|:-----|:-----|
|1.0|引入|

## 另请参阅



#### 其他资源


[getSelectedViewAsync 方法](../../reference/shared/projectdocument.getselectedviewasync.md)
