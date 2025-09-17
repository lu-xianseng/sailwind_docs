# 第 1 章 Logic 自动化服务器

本章介绍 SailWind Logic 中包含的自动化服务器功能。

## OLE 背景

"OLE 解决的根本问题：如何设计一个系统，使得来自不同供应商的二进制组件 - 在世界不同地区、不同时间使用不同编程语言编写 - 能够保证互操作？"

Dr. Dobbs's Journal, 1995年1月

从技术上讲，OLE 代表对象链接与嵌入(Object Linking and Embedding)，但这个名称现在已不够全面。如今，OLE 是基于微软框架的一系列不断增长的功能（包括对象链接与嵌入、自动化、复合文档、编辑、ActiveX、DCOM 等），允许异构应用程序相互通信。

自动化是 OLE 的一项功能，它定义了应用程序与使用相同协议的其他应用程序共享其数据和功能的协议。自动化涉及服务器和客户端。服务器是提供某些数据和功能的实体。客户端（也称为控制器）是使用服务器数据和功能的实体。

不同组织的人员可以在不同时间使用不同的开发工具和语言编写服务器和客户端。唯一需要的共同点是它们使用相同的自动化通信协议。SailWind Logic 是一个自动化服务器，因为它向其他应用程序提供了一些数据（其数据库）和一些功能（如打开设计文件和选择对象）。像 Excel 这样的客户端可以使用自动化协议访问 SailWind Logic 的数据和功能。

SailWind Logic 中的自动化实现允许第三方公司和用户：

- 将其产品与 SailWind Logic 集成

- 扩展 SailWind Logic 的功能集

- 自定义现有的 SailWind Logic 功能

- 使用标准脚本语言（如 Basic）自动化 SailWind Logic 任务

此外，这可以独立于发布周期完成。

## 多版本环境中的自动化

要在同一台机器上存在多个 SailWind 软件版本的环境中使用自动化，您必须了解并应用每个软件安装的正确 COM 版本。

- COM 自动化脚本 这些脚本扩展了主要创作应用程序（例如 SailWind Layout）的功能。这些脚本连接到工具实例以访问其功能。该功能包含在应用程序内的 COM 对象中。正是这些 COM 对象由注册器注册。

- COM 对象 这些对象封装了产品功能，并在用户创建的 COM 自动化脚本中被引用。由于产品功能在不同版本之间有所不同，与每个版本关联的 COM 对象具有唯一版本。例如（本示例中使用的 COM 编号仅用于示例，不反映所提及版本的真实 COM 编号），任何 PADS 9.5 发布对象的版本为 1，X-ENTP VX.2.3 对象的版本为 2，PADS-Pro VX.2.3 对象的版本为 3，PADS VX.2.3 对象的版本为 4 等。对象版本控制使每个版本的产品功能可用于客户创建的脚本。注册这些 COM 对象是注册器（PADS VX.2.3）和配置器（PADS9.5）的主要功能。

## 对象层次结构

下图显示了对象层次结构。

![](/logic/scripts/1/_page_2_Figure_1.jpeg)

![](/logic/scripts/1/_page_2_Figure_2.jpeg)

## 自动化服务器示例

本节提供的通用示例旨在方便您开发专为 SailWind Logic 设计的自动化客户端，或将现有应用程序与 SailWind Logic 集成。

此示例演示了以下技术：

- 为 SailWind Logic 服务器创建自动化客户端的基础知识。如何连接到现有（或新）的 SailWind Logic 实例，如何断开连接，以及如何从 SailWind Logic 访问简单数据。

- 添加事件，当 SailWind Logic 状态更改时通知客户端。

- 访问 SailWind Logic 数据库。

- 客户端与 SailWind Logic 之间的全双工交叉探测功能。

提示：

- 您需要在系统上安装适当的开发工具才能正确运行这些示例。换句话说，要运行 Microsoft Excel 版本的任何示例，请在系统上安装 Excel。要运行 Visual Basic 版本，请安装 Visual Basic。要运行 Visual C++ 版本，请安装 Visual C++。

- 如果遇到示例或您自己的客户端应用程序的困难，请参阅以下主题，其中提供了常见自动化问题的提示。

[故障排除](#page-3-1) [SailWind Logic 自动化客户端示例](#page-3-2)

### 故障排除

如果示例未按描述工作，请检查以下事项：

- 确保只有一个 SailWind Logic 实例正在运行。根据定义，服务器是唯一的。同一服务器的多个实例可能会使客户端感到困惑。

- 退出并重新启动 SailWind Logic 在某些情况下可能会有所帮助，因为 SailWind Logic 在启动时会自动注册为自动化服务器。使用任务管理器确保没有其他 SailWind Logic 实例在后台运行。

- 在退出客户端应用程序之前，确保从 SailWind Logic 断开所有客户端连接。

- 如果上述方法无效，可能需要重新启动机器。

### SailWind Logic 自动化客户端示例

本节详细讨论自动化示例。

目标

介绍使用 Microsoft Visual C++ 5.0（带 MFC）、Microsoft Visual Basic 5.0 和 Microsoft Excel 97 编写 SailWind Logic 自动化客户端应用程序的基础知识。为这些应用程序各提供一个示例。

每个版本的示例都演示了如何：

- 准备客户端与 SailWind Logic 自动化服务器之间的通信。

- 准备指向 SailWind Logic 顶层自动化对象的指针。

- 连接和断开与 SailWind Logic 的连接，无论 SailWind Logic 是否正在运行。

- 从 SailWind Logic 访问简单数据。

- 添加客户端通知（也称为客户端回调），以在 SailWind Logic 中发生更改时通知客户端。

- 访问 SailWind Logic 数据库组件并检索每个组件的详细信息。

- 添加全双工交叉探测并演示服务器锁定机制以提高自动化性能。

> [!tip]
>
> - 此示例的源代码位于 C:\SailWind Projects\Samples\Scripts\Logic\AutomationClient 文件夹中（如果您在 SailWind Logic 安装期间安装了 OLE 示例（典型安装））。代码有详细注释，您可以在开发自己的客户端应用程序时使用它。
>
> - 在运行示例之前，请确保将 OLE 示例从分发 CD 复制到硬盘。
>
> - 要运行示例客户端，请启动 SailWind Logic 并打开原理图文件。

规范

客户端示例是一个对话框，可以在运行时连接和断开与 SailWind Logic 服务器的连接。它可以连接到现有的运行实例或新的 SailWind Logic 实例。

Visual Basic 版本的示例见[图2](#page-5-0)。Visual C++ 版本的示例看起来几乎相同。Excel 97 版本不同，因为 Excel 主要是一个应用程序，而不是开发环境。此示例的 Excel 版本在 Excel 工作表中输出信息。

客户端示例允许在客户端和服务器之间进行选择交叉探测；SailWind Logic 中选定对象的列表始终与客户端列表框中的选定项目列表匹配，反之亦然。选择交叉探测支持多选。如果在列表框中双击某个项目，则会在对话框中显示有关该对象的扩展信息。

客户端自动刷新其值，始终匹配 SailWind Logic 中选定的组件。同样，客户端中的用户选择更改会更改 SailWind Logic 选择列表。

![](/logic/scripts/1/_page_5_Figure_1.jpeg)

![](/logic/scripts/1/_page_5_Figure_2.jpeg)

## 代码示例

本帮助文件中的许多主题都提供了 Basic 代码示例，以说明如何使用自动化属性、方法或事件。

您可以按照以下主题中的描述运行代码示例。

代码示例始终采用以下形式：

| Sub Main       |
|----------------|
| ' 执行某些操作 |
| End Sub        |

免责声明：

SailWind Logic 自动化服务器帮助中的代码示例是免费软件。这些示例作为对其用户的礼遇提供。免费软件按"原样"提供，不对免费软件做出任何明示或暗示的保证，包括对适销性或特定用途适用性的任何暗示保证。

[在 SailWind Logic 中运行代码示例](#page-6-1) [在 SailWind Logic 外部运行代码示例](#page-7-0) [增强示例代码](#page-8-0) [代码示例故障排除](#page-8-1)

### 在 SailWind Logic 中运行代码示例

使用 Sax Basic Engine 运行示例。

步骤

1. 选择示例，包括 Sub Main 和 End Sub 语句。

2. 使用编辑菜单 > 复制将其复制到剪贴板。

3. 在 SailWind Logic 中选择工具菜单 > Basic 脚本 > Basic 脚本编辑器。将出现 Sax Basic Engine 对话框。

4. 使用编辑菜单 > 粘贴将代码示例粘贴到 Basic 编辑器中。

5. 从 Sax Basic Engine 对话框的工具栏中选择开始图标以运行示例。

6. 下图显示了 Application.ActiveDocument 属性属性的代码示例粘贴到 Sax Basic Engine 对话框中的情况。

![](/logic/scripts/1/_page_7_Picture_2.jpeg)

图3. 代码示例粘贴到 Sax Basic Engine 对话框

![](/logic/scripts/1/_page_7_Picture_4.jpeg)

### 在 SailWind Logic 外部运行代码示例

将代码示例与其他 Basic 可编写脚本的应用程序（也称为宿主应用程序）一起使用，例如 Visual Basic、Microsoft Excel 和 Microsoft Word。

步骤

1. 将 SailWind Logic 自动化服务器引用导入宿主应用程序。例如要使用 Excel 作为宿主应用程序：

	- a. 在 Excel 中选择工具/引用。将出现引用对话框。

	- b. 在列表框中，打开 SailWind Logic 类型库复选框。

	- c. 选择确定。

2. 将代码示例粘贴到宿主应用程序的 Basic 编辑器中。要在 Excel 中执行此操作：

	- a. 选择视图/代码。

	- b. 选择编辑/粘贴。

3. 在代码示例的开头添加代码，使用 Basic GetObject 函数将宿主应用程序连接到 SailWind Logic。

4. 将 Basic GetObject 函数返回的对象添加到代码示例中每个 SailWind Logic 自动化方法或属性的开头。

5. 在示例末尾添加代码以断开应用程序与 SailWind Logic 的连接。

6. 下图显示了 Application.ActiveDocument 属性属性的代码示例粘贴并修改到 Microsoft Excel Basic 编辑器中的情况。有关更多信息，请参阅宿主应用程序帮助文件。

图4. Microsoft Excel 97 Basic 编辑器中的代码示例

### 增强示例代码

SailWind Logic 自动化服务器帮助中的代码示例已缩减为最少行数，以快速说明如何使用 SailWind Logic 自动化属性和方法。您可以将这些代码示例作为开发自己的 SailWind Logic 功能的基础。

以下列表提供了可能的增强思路：

- 添加严格的类型检查，在代码开头使用 Option Explicit 声明，并使用 Dim Basic 关键字声明所有变量。这确保 SailWind Logic 可以正确解释您的变量，并可以在可能存在潜在问题的地方生成编译错误。

- 使用 Basic On Error 关键字添加错误检查。错误检查可改善代码对运行时错误的反应。

- 使用 Basic UserDialog 关键字在自定义对话框中输出信息，而不是使用 Basic MsgBox 关键字在简单对话框中输出。

代码示例免责声明

SailWind Logic 自动化服务器帮助中的代码示例是免费软件。这些示例作为对其用户的礼遇提供。免费软件按"原样"提供，不对免费软件做出任何明示或暗示的保证，包括对适销性或特定用途适用性的任何暗示保证。

### 代码示例故障排除

如果代码示例在您的系统上无法正确运行，您可以检查以下事项：

- 确保在 SailWind Logic 中打开了设计。

  几乎所有代码示例都要求在运行示例时打开原理图文件。

- 检查示例对设计文件所做的任何假设。

  例如，某些代码示例可能假设组件 U1 存在。这些假设在每个示例之前的文本中明确说明。如果这些假设不成立，示例代码将无法正常运行。您可以根据自己的原理图调整示例。

- 如果您修改了示例代码以在 SailWind Logic 外部运行，请确保您正确应用了对代码所需的所有更改，如["代码示例"](#page-6-0) [第21页](#page-6-0)中所述。

  常见错误包括忘记导入 SailWind Logic 引用和忘记为 SailWind Logic 自动化常量添加前缀。

免责声明

SailWind Logic 自动化服务器帮助中的代码示例是免费软件。这些示例作为对其用户的礼遇提供。免费软件按"原样"提供，不对免费软件做出任何明示或暗示的保证，包括对适销性或特定用途适用性的任何暗示保证。