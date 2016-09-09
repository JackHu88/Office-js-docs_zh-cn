# 使用运行时日志记录来调试外接程序命令

Office 16 桌面客户端具有可用于记录有用信息的全新功能。 此外，此工具可以帮助你诊断外接程序清单中的错误，如果你正在使用外接程序命令创建清单，那么使用这一工具尤为方便。 

虽然此功能的完整文档还在编制之中，但你可以在本文中找到在通过外接程序命令分析清单时使用此工具调试问题的方法。

##开启运行时日志记录

**重要说明**：运行时日志记录会产生**性能下降**。 请仅在需要调试外接程序中的问题时启用此功能

1. 确保你有支持运行时日志记录的版本。 你需要版本等于或大于 **16.0.7019** 的 **Office 16 桌面**客户端版本
2. 在 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\` 下添加 `RuntimeLogging` 注册表项 
3. 将此项的默认值设置为你想要在其中写入日志的文件的完整路径。 请参阅 [示例注册表项](RuntimeLogging/EnableRuntimeLogging.zip)（解压缩）

你的注册表应如下所示：![](http://i.imgur.com/Sa9TyI6.png)

如果需要关闭此功能，只需将该项从该注册表中删除。 

##诊断命令中的问题
运行时日志记录可用于检测**清单中的问题**，这些问题难以捕获，例如，资源 ID 之间不匹配，或长度无效，并且无法通过 XSD 架构验证发现。 

以下是尝试执行操作的步骤：
 
1. 按照 [自述](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md) 中的说明旁加载你的外接程序。 
2. 如果看不到功能区按钮项目并且外接程序对话框中没有显示任何内容，请查看日志
3. 搜索外接程序的 ID（该 ID 是你在清单中所定义的），以查找属于该外接程序的消息。 日志将此 ID 报告为 `SolutionId`。建议你一次只旁加载一个外接程序，以避免查看过多不属于你的外接程序的消息。 

在下面的示例中，运行时日志记录已帮助标识出一个指向不存在资源文件的控件。 修补程序可更正输入错误（如果有的话），或实际添加缺少的资源。

![](http://i.imgur.com/f8bouLA.png) 

##日志记录的已知问题
运行时日志记录仍具有已知 bug。 你可能会看到几条令人困惑或分类不当的消息。 例如：

- 后接 `Unexpected Parsed manifest targeting different host` 的消息 `Medium  Current host not in add-in's host list` 未正确进行分类。 它们不是错误，你可以放心地忽略它们。
- 消息 `Unexpected   Add-in is missing required manifest fields  DisplayName` 不包含有问题的外接程序的 SolutionId。 但是，很可能这与你正在调试的外接程序无关。 
- 对系统而言，任何 `Monitorable` 消息都会视其为错误。 有时这些消息表示清单中的问题（例如一个已跳过但未引起清单运行失败的拼写错误的元素）。 

