# 第 2 章 自动化服务器参考

SailWind Logic 自动化对象层次结构严格遵循微软标准。

所有 SailWind Logic 自动化对象都基于微软 IDispatch 接口实现其接口。

以下列表标识了自动化对象：

- Application Object

- Attributes Collection Object Property

- Attribute Object

- Component Object

- Document Object

- Field Object

- Fields Collection Object

- Gate Object

- Library Object

- LibraryItem Object

- Measure Object

- Net Object

- Objects Collection Object

- PartType Object

- Pin Object

- Sheets Collection Object

- Sheet Object

- View Object

以下列表标识了自动化常量：

- PlogNetListVersion

- PlogGateVisibility

- PlogDefaultPosition

- PlogMeasureFormat

- PlogLibraryItemType

- PlogObjectType

- PlogUnit

- PlogGridType

- PlogPinElectricalType

- PlogASCIIVersion

## Application 对象

Application 对象是 SailWind Logic 自动化服务器对象层次结构中的根级对象，代表整个 SailWind Logic 应用程序。此对象通常是自动化客户端在访问任何 SailWind Logic 对象、属性或方法之前连接的第一个对象。

### Application.ActiveDocument 属性

此属性返回活动的 SailWind Logic 文档。

**用法**

```

ActiveDocument As Document  on page 419

```

**参数**

无

**描述**

活动文档代表打开的电路图。

**示例**

以下示例代码使用 [Document.Name 属性](#page-84-0) 检索打开的电路图名称。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

MsgBox ""You are working with " & ActiveDocument.Name

End Sub

```

### Application.Application 属性

此属性返回 SailWind Logic Application 对象。

**用法**

```

Application As Application  on page 401

```

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的自动化客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### Application.DefaultFilePath 属性

此属性设置或返回 SailWind Logic 用于打开电路图文件的路径。

**用法**

```

DefaultFilePath As String

```

**参数**

无

**描述**

此属性检查 *SailWindlogic.ini* 文件中的 FileDir 文件夹条目。当您将此属性设置为新值时，*.ini* 文件条目也会更改。

例如，当您使用默认安装设置安装 SailWind Logic 时，路径为 "*C:\<install\_folder>\<version>\Programs*"。

**示例**

以下示例代码更改 SailWind Logic 默认文件路径并通知客户端此更改。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

oldPath = DefaultFilePath

DefaultFilePath="C:\TEMP"

MsgBox "The default file path used to be " & oldPath & " and it was just 

changed to " & DefaultFilePath 

End Sub

```

### Application.FullName 属性

此属性返回 SailWind Logic 应用程序的文件名，包括其路径。

**用法**

```

FullName As String

```

**参数**

无

**描述**

例如，此函数可以返回字符串 "C:\<install\_folder>\*<version>*\Programs \SailWindLogic.exe"。

**示例**

以下示例代码显示 SailWind Logic 的通用名称和实际的 .exe 名称。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

MsgBox "Hi, my name is " & Name & " and I am located in " & FullName 

End Sub

```

### Application.Libraries 属性

此属性返回可用库的集合或特定库。

**用法**

```

Libraries as Collection

Libraries(name  as String) as Library Object

```



**参数**

| 参数 | 描述                                                    |
|----------|----------------------------------------------------------------|
| name     | 要检索的库名称。不应包含通配符。 |

**描述**

无

**示例**

此示例显示可用库的数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

MsgBox "库数量: " & Libraries.Count

### Application.Name 属性

此属性返回 SailWind Logic 应用程序的名称。

**用法**

Name As String

**参数**

无

**描述**

例如，在 SailWind Logic 中，此属性返回字符串 "PowerLogic"。

此属性是 Application 对象的默认属性。

**示例**

以下示例代码显示 SailWind Logic 的通用名称、版本和实际的 .exe 名称。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

MsgBox "您好，我的名称是 " & Name & " 版本 " & Version & "，位于 " & FullName

```

End Sub

**相关主题**

[Application.FullName 属性](#page-5-0)

[Application.Version](#page-12-0) 属性

### Application.ObjectType 属性

此属性返回此对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

无

**示例**

以下示例测试 ObjectType 属性。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main Set items = GetLibraryItems() For Each item In items If item.ObjectType <> plogObjectTypeLibraryItem Then MsgBox "测试失败" End If Next item End Sub

### Application.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As [Application.Application 属性](#page-3-0)

**参数**

无

**描述**

无

### Application.ProgressBar 属性

此属性设置或返回进度条的当前值，以百分比表示。

**用法**

ProgressBar As Integer

**参数**

无

**描述**

此属性允许您检索在 SailWind Logic 中运行的长时间批处理进程的当前进度，或显示长时间 Basic 脚本的进度。

将值设置为大于0或小于100以停用进度条。将此属性与 [Application.StatusBarText](#page-11-0) 属性一起使用。

**示例**

以下示例演示了此属性。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

 StatusBarText="我的批处理进程..." '显示进度文本

```

For i = 0 to 100

ProgressBar = i

Next

ProgressBar = -1 '停用进度条

StatusBarText="" '隐藏进度文本

End Sub

**相关主题**

[Application.StatusBarText](#page-11-0) 属性

[Application.ProgressChange 事件](#page-32-0)

### Application.StatusBarText 属性

此属性设置或返回 SailWind Logic 状态栏上显示的文本。

**用法**

StatusBarText As String

**参数**

无

**描述**

要将 SailWind Logic 状态栏文本设置为空，请将此属性设置为空字符串(" ")。

**示例**

以下示例在 SailWind Logic 状态栏上显示消息。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

StatusBarText="Cool! I can even print my own messages in here!"

```

End Sub

**相关主题**

[Application.ProgressBar 属性](#page-10-0)

[Application.ProgressChange 事件](#page-32-0)

### Application.Version 属性

此属性返回 SailWind Logic 版本号。

**用法**

Version As String

**参数**

无

**描述**

此属性以字符串形式返回应用程序版本，格式为：<主版本号>.<次版本号>，例如"3.0"。

**示例**

以下示例显示应用程序名称和版本。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main MsgBox "Hi, my name is " & Name & " version " & Version

End Sub

**相关主题**

[Application.Name 属性](#page-7-0)

[Application.FullName 属性](#page-5-0)

### Application.Visible 属性

此属性设置或返回 SailWind Logic 是否可见。

**用法**

Visible As Boolean

**参数**

无

**描述**

此属性通常用于以下情况：

- 当自动化客户端使用异步 OLE 自动化调用（如 Basic 函数 CreateObject）启动 SailWind Logic 自动化服务器时。自动化服务器始终以不可见方式启动（这是客户端/服务器规则）。因此，如果需要，您可以使用此属性使 SailWind Logic 可见。

- 当客户端尝试关闭 SailWind Logic（参见[Application.Quit 方法](#page-28-0)）时，通过使 SailWind Layout 不可见，断开与它的连接，并让服务器正常关闭。

- 当客户端需要使服务器窗口成为活动窗口，使其显示在其他应用程序窗口之上时。

**示例**

以下示例使 SailWind Logic 不可见，等待一秒，然后再次使其可见。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main Visible = False Wait 1 Visible = True End Sub

### Application.CreateLibrary 方法

此方法返回特定库或可用库的集合。

**用法**

Libraries as Collection

Libraries(Name as String) as Library

**参数**

| 参数 | 描述 |
|------|------|
| name | 要检索的库名称。不应包含通配符。 |

**描述**

无

**示例**

以下示例显示可用库的数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

MsgBox "Number of libraries: " & Libraries.Count

### Application.ExportLibraryItems 方法

此方法从库项生成 PADS 格式的 ASCII 文件。

**用法**

ExportLibraryItems (filename as String, items as Collection)

**参数**

| 参数 | 描述 |
|------|------|
| filename | 要导出库项的文件名。不要指定扩展名。 |
| items | 可选。要导出的项集合。如果省略，则导出集合中的所有项。 |

**描述**

最多可以生成四个文件。

**示例**

此示例子程序将库中以"R"开头的库项导出到指定文件。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set coll = GetLibraryItems(, "R*")

```

```

ExportLibraryItems("C:\sample", coll)

```

### Application.GetConfigParamInt 方法

此方法从 SailWindlogic.ini 文件的指定参数和节中检索整数值。

**用法**

GetConfigParamInt(

sectionName as String,

paramName as String,

defaultValue as Integer) as Integer

**参数**

| 参数 | 描述 |
|------|------|
| sectionName | 包含参数名的节名称。 |
| paramName | 要检索其关联整数值的参数名称。 |
| defaultValue | 如果找不到参数名，则返回默认值。 |

**返回值**

参数值或默认值。

**示例**

```

MsgBox Application.GetConfigParamInt("general", " Object Selector Filter", 

 0)

```

### Application.GetConfigParamString 方法

此方法从 SailWindlogic.ini 文件的指定参数和节中检索字符串。

**用法**

GetConfigParamString(

sectionName as String,

paramName as String,

defaultValue as String) as String

**参数**

| 参数 | 描述 |
|------|------|
| sectionName | 包含参数名的节名称。 |
| paramName | 要检索其关联字符串的参数名称。 |
| defaultValue | 如果找不到参数名，则返回默认值。 |

**返回值**

参数值或默认值。

**示例**

```

MsgBox Application.GetConfigParamString("directories", " FileDir", " 

 C:\SailWind Projects\")

```

### Application.GetLibraryItems 方法

此方法返回所有可用库中的库项集合、给定类型的所有项集合或特定项。

**用法**

```

GetLibraryItems (type as PlogLibraryItemType, name as String) 

 as LibraryItem Object

```

**参数**

| 参数 | 描述 |
|------|------|
| type | 要检索的项类型。可选；默认值为 plogLibraryItemTypeAll。 |
| name | 要检索的项名称。可选；可以包含通配符、范围、列表。 |

**描述**

**无**

**示例**

此示例显示可用库项的数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

MsgBox "Number of library items: " & GetLibraryItems().Count

### Application.LockServer 方法

此方法锁定 SailWind Logic 自动化服务器。

**用法**

LockServer()

**参数**

**无**

**描述**

当客户端对 SailWind Logic 服务器进行多次 OLE 调用时，使用此函数可加快 OLE 调用处理速度。

![](/logic/scripts/2/_page_19_Picture_9.jpeg)

**警告：**

切勿忘记解锁已锁定的服务器。

服务器锁定机制通过将 OLE 服务器方法或属性调用的访问时间加快2到8倍来提高处理速度。服务器锁定机制很危险，因为它会禁用服务器中的许多内部后台任务，例如内存清理和视觉更新。这些服务器任务必须定期执行，但被禁用以便可以快速处理 OLE 传入调用。

通常，当客户端需要对服务器进行数百次连续调用时，您会锁定服务器。不要将服务器锁定太长时间（不超过几分钟）。当客户端只需要进行少量 OLE 调用时，不需要锁定服务器；在这种情况下，速度不会大大提高。

**示例**

以下示例代码显示如何使用此方法。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

LockServer

' 执行一些耗时的操作，这些操作会对 SailWind Logic 自动化服务器进行多次调用

UnlockServer

**相关主题**

[Application.UnlockServer 方法](#page-30-0)

### Application.Measure 方法

该方法创建并返回测量值对象。

**用法**

```

Measure(Value As Variant, [DefaultUnit As String=""]) As Measureon page 195

```

**参数**

| 参数        | 描述                                                                                            |
|-------------|------------------------------------------------------------------------------------------------|
| Value       | 表示测量值的字符串或数字，包含可选前缀和<br>可选物理单位。 |
| DefaultUnit | [可选 page 314] 包含默认前缀和/或物理单位的字符串。                        |

**描述**

该属性解析如"100pF"的字符串值，并创建特殊对象，从中可以提取额外信息如实值、单位名称或数量名称。

如果 *Value* 参数包含单位信息，则忽略 *DefaultUnit* 参数。

如果 *Value* 参数不包含单位信息且 *DefaultUnit* 为空，则创建新的尺寸/维度测量，*Value* 参数将默认解释为当前SailWind Logic设计单位(mils)中的数字。

如果解析器无法识别Value参数中的测量值，它将创建一个虚拟[测量对象](#page-170-0)，其[Measure.Value](#page-179-0)属性等于0.0，[Measure.Text](#page-176-0)属性等于*Value*参数。

注意Measure("100", "pF")和Measure("100pF")之间有区别。虽然两个版本都表示相同的物理测量值100pF，但第一个版本存储了不含单位的精确文本表示。

**示例**

以下示例为部件"C1"添加测量类型属性，假设它存在于打开的设计中。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

Set C1 = ActiveDocument.Components("C1")

C1.Attributes.Add "Capacitance", Measure("500pF")

End Sub

```

**相关主题**

[Attribute.Measure 方法](#page-51-0)

### Application.ModelessCmd 方法

该方法执行无模命令。

**用法**

ModelessCmd( command As String ) As Void

**参数**

command

**返回值**

无

**描述**

该方法允许您执行应用程序中无模命令对话框中可用的所有无模命令。

**示例**

以下示例代码将活动工作表切换到第二个工作表。有关运行此示例的更多信息，请参见"代码示例"。

Sub Main

ModelessCmd("sh 2")

### Application.OpenDocument 方法

该方法打开一个SailWind Logic原理图文件。

**用法**

OpenDocument(filename As String) As Documenton page 419

**参数**

| 参数     | 描述               |
|----------|-------------------|
| filename | 要打开的文件名。 |

**返回值**

如果函数成功，返回值是新打开的[文档对象](#page-71-0)。

如果函数失败，返回值为Nothing。

**描述**

如果*filename*不包含文件的完整路径，SailWind Logic使用[Application.DefaultFilePath属性](#page-4-0)指定的路径来定位文件。

如果*filename*为Nothing，或为空字符串，则创建一个新的空白原理图文件。

如果无法找到或打开*filename*指定的文件，返回值为当前文档。

该方法不检查当前打开的文件是否已保存。客户端有责任使用[Document.Saved属性](#page-91-0)检查打开的示意图是否需要保存。

如果在处理过程中发生SailWind Logic故障，该方法会生成一个[异常](#page-290-0)。

**示例**

以下示例代码打开*DEMO.SCH*，假设它存在于[Application.DefaultFilePath属性](#page-4-0)属性指定的文件夹中。然后示例显示打开的文件名。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

OpenDocument(DefaultFilePath & "\Samples\DEMO.SCH")

```

MsgBox ActiveDocument.FullName & " has just been opened."

```

**相关主题**

[Application.ActiveDocument属性](#page-2-0) [Application.DefaultFilePath属性](#page-4-0) [Document.Name属性](#page-84-0)

[Document.Saved属性](#page-91-0)

### Application.OpenDocumentNoLock 方法

该方法打开SailWind Logic原理图文件而不锁定文件。

**用法**

```

OpenDocumentNoLock(filename As String) As Documenton page 419

```

**参数**

| 参数     | 描述               |
|----------|-------------------|
| filename | 要打开的文件名。 |

**返回值**

如果函数成功，返回值是新打开的文档。

如果函数失败，返回值为当前文档。

### Application.OpenTempDocument 方法

该方法的工作方式与OpenDocumentNoLock相同，但另外不会将文件名添加到MRU(最近使用)列表中。此方法主要用于宏测试。

**用法**

OpenTempDocument(filename As String) As Documenton page 419

**参数**

| 参数     | 描述               |
|----------|-------------------|
| filename | 要打开的文件名。 |

**返回值**

如果函数成功，返回值是新打开的文档。

如果函数失败，返回值为当前文档。

### Application.Quit 方法

该方法关闭SailWind Logic。

**用法**

Quit ()

**参数**

无

**描述**

![](/logic/scripts/2/_page_28_Picture_8.jpeg)

**警告：** 不要使用此方法。

Microsoft要求所有自动化应用程序对象实现此方法。然而，调用此方法违反了一些重要的客户端/服务器规则：

- 服务器在所有客户端断开连接之前不能关闭。根据定义，客户端在从服务器断开连接后不可能调用Quit方法(或任何其他服务器方法)，因此客户端不可能关闭服务器。

- 客户端无法知道是否有其他客户端连接到服务器。因此，它不应该关闭服务器。

- 服务器有自己的关闭管理流程：当最后一个客户端从服务器断开连接时，仅当其图形用户界面(GUI)不活动(不可见)时，服务器才会自动关闭。否则，服务器保持活动状态。

要强制关闭SailWind Logic，自动化客户端需要使用[Application.Visible](#page-13-0)属性使SailWind Logic不可见，然后断开连接。如果此时没有其他客户端连接到SailWind Logic，SailWind Logic会自动关闭。如果自动化客户端是在Sax Basic Engine中运行的Basic脚本，它永远无法成功关闭SailWind Logic自动化服务器。

**相关主题**

[Application.Quit事件](#page-33-0)

### Application.RunMacro 方法

该方法运行SailWind Logic宏。

**用法**

RunMacro(filename As String, macroname As String)

**参数**

| 参数      | 描述                    |
|-----------|------------------------|
| filename  | 要使用的宏文件名。 |
| macroname | 要运行的宏名。      |

**描述**

如果filename不包含文件的完整路径，SailWind Logic使用[Application.DefaultFilePath属性](#page-4-0)指定的路径来定位文件。

如果*filename*为空，或为空字符串，则使用当前默认的SailWind Logic宏文件。

如果*macroname*为空，或不是有效的宏名，则不执行任何操作。

如果在处理过程中发生SailWind Logic故障，该方法会生成一个[异常](#page-290-0)。

**示例**

以下示例代码运行SailWind Logic宏MACRO1，记录在宏文件*MACROS.MCR*中，假设宏存在于[Application.DefaultFilePath](#page-4-0)属性指定的文件夹中。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

```

RunMacro(DefaultFilePath & "\MACROS.MCR", "MACRO1")

### Application.UnlockServer 方法

该方法解锁SailWind Logic自动化服务器。

**用法**

UnlockServer()

**参数**

无

**返回值**

无

**示例**

以下示例代码展示了如何使用此方法。有关运行此示例的更多信息，请参见代码示例。

Sub Main

LockServer

' 执行一些耗时的操作，这些操作会多次调用SailWind Logic自动化服务器

UnlockServer

End Sub

**相关主题**

[Application.LockServer方法](#page-19-0)

### Application.OpenDocument 事件

此事件在程序打开新文档后触发。

**用法**

Application\_OpenDocument (Doc As Documenton page 419)

**参数**

| 参数 | 描述                      |
|------|--------------------------|
| Doc  | 已打开的文档对象。 |

**描述**

此事件在 SailWind Logic 打开新文档后触发。

**相关主题**

[Application.OpenDocument 方法](#page-24-0)

### Application.ProgressChange 事件

此事件在状态栏值更改后触发。

**用法**

Application\_ProgressChange

**参数**

无

**描述**

此事件在进度条值更改后触发。

### Application.Quit 事件

此事件用于关闭程序。

**用法**

Application\_Quit ()

**参数**

无

**描述**

此事件在 SailWind Logic 退出前触发。

## Attributes 集合对象属性

Attributes 集合对象是 SailWind Logic 组件对象属性的集合。

组件对象即 [Component.Attributes 属性](#page-54-0)。

以下列表标识了 Attributes 属性和方法：

[Attributes.Application 属性](#page-35-0) [Attributes.Count 属性](#page-36-0) [Attributes.Item 属性](#page-37-0) [Attributes.Parent 属性](#page-39-0) [Attributes.Add 方法](#page-40-0) [Attributes.Delete 方法](#page-42-0)

### Attributes.Application 属性

此属性返回 SailWind Logic Application 对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的大型自动化客户端应用程序中。使用该属性可快速识别对象所属的应用程序。

### Attributes.Count 属性

此属性返回属性数量。

**用法**

Count As Long

**参数**

无

**描述**

无

**示例**

以下示例代码检索名为 U1 的部件的属性数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

```

 MsgBox "部件 U1 的属性数量为 " & 

 ActiveDocument.Components("U1").Attributes.Count

```

### Attributes.Item 属性

此属性根据索引或名称返回属性。

**用法**

Item(*index* As Long) As [Attributeon page 69](#page-44-0)

Item(*name* As String) As [Attributeon page 69](#page-44-0)

**参数**

|  | 参数 | 描述                                             |
|--|------|-------------------------------------------------|
|  | index | 要检索的属性在集合中的索引。 |
|  | name  | 要检索的属性名称。              |

**描述**

这是 Attributes 集合对象的默认成员。

如果 *index* 或 *name* 参数无效，此属性会生成[异常](#page-290-0)。

**示例**

以下示例展示了两种不同的方法来遍历打开的原理图中组件对象的所有属性。通常推荐第二种方法，因为它更简洁高效。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

' 方法1：使用 Object.Item 属性

Sub Main

Set attrs = ActiveDocument.Components("U1").Attributes 

For I=1 To attrs.Count 

Set thisAttr = attrs.Item(I) 

' 对此属性 thisAttr 执行某些操作 

Next I

```

' 方法2：不使用 Object.Item 属性（推荐方法）

Sub Main

For Each nextAttr in ActiveDocument.Components("U1").Attributes

' 对此属性 nextAttr 执行某些操作

Next nextAttr

### Attributes.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Documenton page 419

**参数**

无

**描述**

无

### Attributes.Add 方法

此方法添加新属性。

**用法**

Add(*name* As String, [*value* As Variant]) As Attribute Objects

**参数**

| 参数 | 描述                                                                                                                                                         |
|------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| name | 新属性的名称。                                                                                                                                              |
| value | [可选参数on page 314] 新属性的值。属性类型可以是<br>Boolean、Byte、Single、Integer、PortInt、Long、Double、String 或 Measure 对象。 |

**返回值**

打包为 Attribute Objects 的新属性。

**描述**

SailWind Logic 仅支持字符串属性值，因此任何其他类型都将转换为字符串类型。

如果 name 参数已是现有属性或不是有效的属性名称，此属性会生成[异常](#page-290-0)。

**示例**

以下示例代码向组件 U1 和网络 +5V 添加了几个不同类型的属性，假设它们存在于打开的原理图中。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

With ActiveDocument.Components("U1").Attributes

' 忽略当属性已存在时生成的异常

On Error Resume Next

.Add("COST", 12.99)

.Add("HEIGHT", 120)

```

.Add("COMMENT", "待办：获取规格书")

End With

ActiveDocument.Nets("+5V").Attributes.Add "Voltage", Measure("5V")

End Sub

**相关主题**

[Attributes.Delete 方法](#page-42-0)

[Attribute.Measure 方法](#page-51-0)

[Application.Measure 方法](#page-21-0)

### Attributes.Delete 方法

此方法删除属性。

**用法**

Delete(*index* As Long)

Delete(*name* As String)

**参数**

| 参数 | 描述                                           |
|------|-----------------------------------------------|
| index | 要删除的属性在集合中的索引。 |
| name  | 要删除的属性名称。              |

**返回值**

**无**

**描述**

在以下情况下，此属性会生成[异常](#page-290-0)：

- 如果 name 参数不是现有属性。

- 如果 name 参数不是有效的可删除属性。

- 如果 index 参数大于现有属性的数量。

**示例**

以下示例代码删除打开的原理图中所有组件的 COST 属性。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

For Each nextComp In ActiveDocument.Components

' 忽略当属性不存在时生成的异常

On Error Resume Next 

nextComp.Attributes.Delete("COST")

```

Next nextComp

End Sub

**相关主题**

[Attributes.Add 方法](#page-40-0)

## Attribute 对象

属性对象表示 SailWind 逻辑组件对象（部件）的一个属性。

该对象通常从[属性集合对象属性](#page-34-0)中获取。

### Attribute.Application 属性

该属性返回 SailWind 逻辑应用对象。

**用法**

Application As [应用对象](#page-1-0)

**参数**

无

**描述**

该属性标识对象为 SailWind 逻辑自动化对象。所有自动化服务器应用都有一个应用对象，所有自动化对象都有一个应用属性。该属性通常用于处理来自不同源（如不同的自动化服务器应用）的大量对象的大型自动化客户端应用中。使用该属性可快速识别对象所属的应用。

### Attribute.Name 属性

该属性返回属性的名称。

**用法**

Name As String

**参数**

无

**描述**

无

**示例**

以下示例代码列出了组件 U1 的所有属性，并将该列表放入自定义对话框。此示例使用 SailWind Logic 中 Sax Basic 引擎的用户对话框编辑器。更多信息请参阅 Sax Basic 编辑器在线帮助。

有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Dim ListAttrs$(10000)

Sub Main

index = 0

For Each nextAttr In ActiveDocument.Components("U1").Attributes

ListAttrs$(index) = nextAttr.Name 

index = index + 1

Next nextAttr

' This piece of code is automatically generated by the SailWind Logic Basic 

 Dialog Editor. 

 Begin Dialog UserDialog 180,238,"U1 Attributes" ' %GRID:10,7,1,1

 ListBox 10,7,160,203,ListAttrs(),.ListBox1

 OKButton 10,210,160,21

```

End Dialog

Dim dlg As UserDialog

Dialog dlg

End Sub

**相关主题**

[Attribute.值](#page-49-0) 属性

### Attribute.Parent 属性

该属性返回对象的父级。

**用法**

Parent As [文档对象](#page-71-0)

**参数**

无

**描述**

SailWind Logic 仅支持组件和网络的属性；因此，父级只能是组件或网络。

### Attribute.Value 属性

该属性设置或返回属性的值。

**用法**

Value As Variant

**参数**

无

**描述**

这是一个默认属性。

属性的值可以是 Boolean、Byte、Single、Integer、PortInt、Long、Double、String 或 Measure 对象类型。

如果值类型与属性类型不匹配，该属性会生成[异常](#page-290-0)。

**示例**

以下示例代码将打开的原理图中所有 LED 组件的 COST 属性更改为 US\$3.99。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main For Each nextComp In ActiveDocument.Components If nextComp.PartType="LED" Then ' avoid exceptions generated when that attribute does not exist if nextComp.Attributes("COST") is Nothing Then nextComp.Attributes("COST").Value = 3.99 End If End If Next nextComp

**相关主题**

[Attribute.名称属性](#page-46-0)

### Attribute.Measure 方法

该方法创建并返回测量值对象。

**用法**

Measure As [应用.测量方法](#page-21-0)

**参数**

无

**描述**

该属性解析属性字符串值（如"100pF"）并创建特殊对象，从中可以提取其他信息，如实值、单位名称、数量名称等。

如果内部解析器无法识别属性中的测量值，它会创建一个虚拟的[应用.测量](#page-21-0) [方法](#page-21-0)对象，其[测量.值](#page-179-0) 属性属性等于0.0，[测量.文本](#page-176-0) 属性属性等于[Attribute.值](#page-49-0) 属性属性。

**示例**

以下示例显示假设存在于打开的原理图中的部件 C1 的电容器"值"属性("电容")的数量名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set C1 = ActiveDocument.Components("C1") 

MsgBox C1.Attributes("Value").Measure.Name 'shows "Capacitance"

```

End Sub

**相关主题**

[应用.测量方法](#page-21-0)

[测量.名称属性](#page-172-0)

## Component 对象

组件对象表示存在于打开的原理图中的物理部件。组件对象由一个或多个可见或未使用的门对象组成。

以下按钮列出了组件对象中的属性和方法。

### Component.应用属性

该属性返回 SailWind 逻辑应用对象。

**用法**

Application As Application on page 401

**参数**

无

**描述**

该属性标识对象为 SailWind 逻辑自动化对象。所有自动化服务器应用都有一个应用对象，所有自动化对象都有一个应用属性。该属性通常用于处理来自不同源（如不同的自动化服务器应用）的大量对象的大型自动化客户端应用中。使用该属性可快速识别对象所属的应用。

### Component.属性属性

该属性返回组件的所有属性集合。

**用法**

Attributes As Attributes

Attributes(*name* As String) As [属性第69页](#page-44-0)

**参数**

| 参数 | 描述                              |
|------|-----------------------------------|
| name | 现有组件属性的名称。              |

**描述**

当传递现有属性名称给该属性时，它返回该组件[属性第69页](#page-44-0)对象。否则，它返回[属性集合对象](#page-34-0) [属性](#page-34-0)中所有组件属性的集合。

**示例**

以下示例代码使用[Attribute.计数属性](#page-36-0)属性检索组件 U1 的属性数量，假设它存在于打开的原理图中。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set attrs = ActiveDocument.Components("U1").Attributes

MsgBox "There are " & attrs.Count & " attribute(s) in component U1."

```

### Component.门属性

该属性返回该组件中已使用门的对象集合。

**用法**

Gates As [对象第219页](#page-194-0)

Gates(*name* As String) As [门第155页](#page-130-0)

**参数**

| 参数 | 描述               |
|------|-------------------|
| name | 现有门的名称。     |

**描述**

该属性仅返回部件的已使用门。

当传递现有门名称给该属性时，它返回该[门对象](#page-130-0)。否则，它返回[对象集合对象](#page-194-0)中组件所有已使用门的集合。

**示例**

以下示例遍历组件 U1 中的所有已使用门。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 For Each nextGate In ActiveDocument.Components("U1").Gates

 'Do something with gate

```

Next

### Component.名称属性

该属性返回组件的名称。

**用法**

Name As String

**参数**

无

**描述**

例如，该属性为组件 U1 返回字符串"U1"。

该属性是组件对象的默认属性。

**示例**

以下示例代码检索打开的原理图中的所有组件，并将该列表放入自定义对话框列表框。当在列表框中选择组件时，示例在 SailWind Logic 中选择该组件。此示例使用 SailWind Logic 中 Sax Basic 引擎的用户对话框编辑器。更多信息请参阅 Sax Basic 编辑器在线帮助。

有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```
Dim ListComps$(10000)

Sub Main

index = 0 

For Each nextComp In ActiveDocument.Components 

ListComps$(index) = nextComp.Name 

index = index + 1 

Next nextComp 

' This piece of code is automatically generated by the SailWind Logic Basic 

 Dialog Editor. 

Begin Dialog UserDialog 180,238,"Parts",.CallbackFunc ' %GRID:10,7,1,1 

ListBox 10,7,160,203,ListComps(),.ListBox1

```

OKButton 10,210,160,21 End Dialog Dim dlg As UserDialog Dialog dlg End Sub

当对话框中发生某些事情时，系统会自动调用以下函数；它用于轻松处理用户操作。

```

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)

Select Case Action%

Case 2 ' Value changing or button pressed

If DlgItem$="ListBox1" Then

ActiveDocument.SelectObjects(plogObjectTypeAll, , False)

'get part by name 

Set comp = ActiveDocument.Components(ListComps(SuppValue%))

'select part

comp.Selected = True

'activate sheet where first gate of the part is located

comp.Gates(1).Sheet.Activate

End If

End Select

End Function
```

### Component.ObjectType 属性

此属性返回对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

此属性返回 [PlogObjectType.](#page-286-0)

SailWind Logic 自动化服务器中的所有 SailWind Logic 数据库对象都实现了此属性，以弥补 Basic 中没有 Visual C++® QueryInterface 函数等效功能的问题。

此属性通常用于：

- 在异构的 [对象集合](#page-194-0) [对象](#page-194-0) 中识别 SailWind Logic 数据库对象的类型。

- 当实现一个依赖于作为参数传递的 SailWind Logic 数据库对象类型的通用例程时。例如：

Sub DoSomething(dbObject As Object)

Select Case dbObject.ObjectType

Case plogObjectTypeComponent

' 对组件对象执行特定操作

Case plogObjectTypeNet

' 对网络对象执行特定操作

Case plogObjectTypePin

' 对引脚对象执行特定操作

Case plogObjectTypeGate

' 对门对象执行特定操作

Case Else MsgBox "不是 SailWind Logic 数据库对象" End Select End Sub

### Component.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Documenton page 419

**参数**

无

**描述**

无

### Component.PartType 属性

此属性返回组件的部件类型。

**用法**

PartType As String

**参数**

无

**示例**

以下示例代码获取组件 U1 的部件类型，假设它存在于打开的示意图中。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

Sub Main

MsgBox "U1 的部件类型是 " & ActiveDocument.Components("U1").PartType

End Sub

**相关主题**

[Component.PartTypeObject](#page-63-0) 属性

### Component.PartTypeLogic 属性

此属性返回组件部件类型的逻辑系列。

**用法**

PartTypeLogic As String

**参数**

无

**描述**

无

**示例**

以下示例代码获取组件 U1 部件类型的逻辑系列，假设它存在于打开的示意图中。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

Sub Main

```

MsgBox "U1 的逻辑系列是 " & 

 ActiveDocument.Components("U1").PartTypeLogic

```

End Sub

**相关主题**

[PartType.Logic](#page-213-0) 属性

### Component.PartTypeObject 属性

此属性返回此组件的部件类型对象。

**用法**

PartTypeObject As [PartType on page 234](#page-209-0)

**参数**

无

**描述**

无

**示例**

以下示例代码获取组件 U1 的部件类型对象，假设它存在于打开的示意图中。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

```
Sub Main


MsgBox "U1 的部件类型是 " & 

 ActiveDocument.Components("U1").PartTypeObject.Name


End Sub
```

**相关主题**

[Component.PartType](#page-61-0) 属性

### Component.PCBDecal 属性

此属性返回组件的 PCB 封装。

**用法**

PCBDecal As String

**参数**

无

**描述**

无

**示例**

以下示例代码获取组件 C1 的当前封装，假设它存在于打开的示意图中。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

Sub Main

MsgBox "C1 PCB 封装是 " & ActiveDocument.Components("C1").PCBDecal

### Component.Pins 属性

此属性返回组件的所有引脚集合。

**用法**

Pins As [Objectson page 219](#page-194-0)

Pins(*name* As String) As [Pinon page 246](#page-221-0)

**参数**

| 参数 | 描述              |
|------|------------------|
| name | 现有引脚的名称。 |

**描述**

返回组件的所有引脚，包括：

- 可见的已使用门引脚（已使用门的引脚）。

- 不可见的"信号"引脚（不属于任何门，但与某些信号相关联）。

- 不可见的未使用门引脚（属于未使用的门）。

- 不可见的未使用部件引脚（不属于任何门且不与任何信号相关联）。

当传递现有引脚 *name* 给此属性时，它将返回该 [引脚对象](#page-221-0)。否则，它将返回 [对象集合对象](#page-194-0) 中组件的所有引脚集合。

**示例**

以下示例计算部件 U1 中所有类型引脚的数量。

```
Sub Main

 SigPins = 0

 AllPins = 0

 UsedGatePins = 0

 UnusedGatePins = 0

 UnusedPartPins = 0


 UnconnectedPins = 0

Set U1 = ActiveDocument.Components("U1")

'获取总引脚数

AllPins = U1.Pins.Count

For Each p In U1.Pins

If p.Gate Is Nothing Then '如果引脚不属于任何门

If p.Net Is Nothing Then '如果引脚未连接则为未使用

UnusedPartPins = UnusedPartPins + 1

Else

SigPins = SigPins + 1 '否则是信号引脚

End If

Else

 If p.Gate.Sheet Is Nothing Then '如果引脚属于门但门未安装

UnusedGatePins = UnusedGatePins + 1

Else

UsedGatePins = UsedGatePins + 1

End If

End If

'如果引脚未连接

If p.Net Is Nothing Then

 UnconnectedPins = UnconnectedPins + 1 End If Next MsgBox "总引脚数是 " & AllPins & vbCr & _ "信号引脚数是 " & SigPins & vbCr & _ "已使用门引脚数是 " & UsedGatePins & vbCr & _ "未使用门引脚数是 " & UnusedGatePins & vbCr & _ "未使用部件引脚数是 " & UnusedPartPins & vbCr & _ vbCr & _ "未连接引脚数是 " & UnconnectedPins, _ "U1 部件的引脚信息"
```

### Component.Selected 属性

该属性设置或返回组件是否被选中✅。

**用法**

Selected As Boolean

**参数**

无

**描述**

当组件的一个或多个门被选中✅时，该组件即被视为选中✅状态。您也可以使用 [Document.SelectObjects 方法](#page-111-0) 或 [Objects.Select 方法](#page-206-0) 来选中✅ SailWind Logic 数据库对象。

**示例**

以下示例代码仅选中✅组件 U1（假设它存在于打开的电路图中），并激活其所在的图纸。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```
Sub Main ActiveDocument.SelectObjects(,,False)

 ActiveDocument.Components("U1").Selected = True

 ActiveDocument.Components("U1").Gates(1).Sheet.Activate

End Sub
```

**相关主题**

[Document.SelectionChange 事件](#page-117-0)

### Component.UnusedGates 属性

该属性返回此组件中所有未使用门的对象集合。

**用法**

UnusedGates As [Objects 第219页](#page-194-0)

UnusedGates(*name* As String) As [Gate 第155页](#page-130-0)

**参数**

| 参数 | 描述               |
|------|-------------------|
| name | 现有门的名称。      |

**描述**

该属性仅返回部件的未使用门。

当传入现有门名称时，返回该 [门对象](#page-130-0)。否则，返回组件所有未使用门的集合，封装为 [Objects 集合对象](#page-194-0)。

**示例**

以下示例获取组件 U1 中门的总数。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 Set u1 = ActiveDocument.Components("U1")

 MsgBox "部件 U1 的总门数为 " & (u1.Gates.Count + 

 u1.UnusedGates.Count)

```

### Component.Delete 方法

该方法从电路图中删除此部件（包括其所有门）。

**用法**

Delete

**参数**

无

**描述**

无

**示例**

以下示例从电路图中删除部件 U1。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

ActiveDocument.Components("U1").Delete

## Document 对象

Document 对象表示 SailWind Logic 中当前打开的电路图文件。通常通过 [Application.ActiveDocument 属性](#page-2-0) 获取此对象。以下是 Document 属性、方法和事件的列表：

[Document.ActiveSheet 属性](#page-73-0) [Document.ActiveView](#page-74-0) 属性 [Document.AncestorSheets 属性](#page-75-0) [Document.Application 属性](#page-76-0) [Document.Components 属性](#page-77-0) [Document.Fields 属性](#page-79-0) [Document.FullName 属性](#page-80-0) [Document.Gates 属性](#page-81-0) [Document.GridX 属性](#page-82-0) [Document.GridY 属性](#page-83-0) [Document.Name 属性](#page-84-0) [Document.Nets 属性](#page-85-0) [Document.Parent 属性](#page-87-0) [Document.PartTypes](#page-88-0) 属性 [Document.Path 属性](#page-89-0) [Document.Pins 属性](#page-90-0) [Document.Saved 属性](#page-91-0) [Document.Sheets 属性](#page-92-0) [Document.Activate 方法](#page-93-0) [Document.ExportASCII 方法](#page-94-0) [Document.ExportNetList 方法](#page-95-0) [Document.GenerateECO 方法](#page-96-0) [Document.GetColor 方法](#page-97-0) [Document.GetObjects 方法](#page-99-0) [Document.ImportASCII 方法](#page-102-0) [Document.ImportECO 方法](#page-103-0) [Document.IntegrityTest](#page-104-0) 方法 [Document.Save 方法](#page-105-0) [Document.SaveAs 方法](#page-106-0) [Document.SaveAsNoLock 方法](#page-107-0) [Document.SaveAsTemp](#page-108-0) 方法 [Document.SaveNoLock 方法](#page-109-0) [Document.SaveTemp 方法](#page-110-0) [Document.SelectObjects 方法](#page-111-0) [Document.SetColor 方法](#page-114-0) [Document.Save 事件](#page-116-0)

[Document.SelectionChange 事件](#page-117-0)

### Document.ActiveSheet 属性

该属性返回此文档中的活动图纸。

**用法**

ActiveSheet As [Sheet 第265页](#page-240-0)

**参数**

无

**描述**

无

**示例**

以下示例获取活动图纸的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

MsgBox "活动图纸是 " & ActiveDocument.ActiveSheet.Name

### Document.ActiveView 属性

该属性返回 SailWind Logic 的活动视图。

**用法**

ActiveView As [View 第294页](#page-269-0)

**参数**

无

**描述**

活动视图表示打开电路图中活动图纸的主视图。

**示例**

以下示例代码将当前 SailWind Logic 视图平移到活动图纸的原点（使用 [View.Pan](#page-279-0) 方法）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

ActiveDocument.ActiveView.Pan(0,0)

### Document.AncestorSheets 属性

该属性返回此文档中不是任何其他图纸子图纸的图纸集合。

**用法**

AncestorSheets As [Sheets 第265页](#page-240-0)

AncestorSheets(*name* As String) As [Sheet 第265页](#page-240-0)

**参数**

• *name*

现有图纸的名称。

**描述**

当传入现有图纸名称时，返回封装为 [Sheet 对象](#page-240-0) 的该图纸。否则，返回打开电路图中根图纸的集合，封装为 [Sheets 集合对象](#page-261-0)。

**示例**

以下示例获取活动文档中祖先图纸的总数。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

 MsgBox "祖先图纸数量为 " & ActiveDocument.AncestorSheets.Count

### Document.Application 属性

该属性返回 SailWind Logic Application 对象。

**用法**

Application As Application 第401页

**参数**

无

**描述**

此属性标识对象为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有 Application 对象，所有自动化对象都有 Application 属性。此属性通常用于处理来自不同源（如不同自动化服务器应用程序）的大量对象的自动化客户端应用程序中，用于快速识别对象所属的应用程序。

### Document.Components 属性

该属性返回所有组件的集合。

**用法**

Components As [Objects 第219页](#page-194-0)

Components(*name* As String) As [Component 第77页](#page-52-0)

**参数**

| 参数 | 描述                |
|------|--------------------|
| name | 现有组件的名称。     |

**描述**

当传入现有组件名称时，返回该 [组件对象](#page-52-0)。否则，返回所有组件的集合，封装为 [Objects 集合对象](#page-194-0)。

**示例**

以下示例代码使用 [Objects.Count 属性](#page-196-0) 获取打开电路图中的组件数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set comps = ActiveDocument.Components 

MsgBox "在 " & ActiveDocument.Name & " 中有 " & comps.Count & " 个部件"

```

End Sub

以下示例代码使用 [Component.Pins 属性](#page-65-0) 和 [Objects.Count 属性](#page-196-0) 获取组件 U1 的引脚数量（假设它存在于打开的电路图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set compU1 = ActiveDocument.Components("U1") 

MsgBox "组件 " & compU1.Name & " 有 " & compU1.Pins.Count & " 个引脚。"

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Sheet.Components 属性](#page-243-0)

### Document.Fields 属性

该属性返回字段的集合。

**用法**

Fields As [Objects 第219页](#page-194-0)

**参数**

无

**描述**

无

**示例**

以下示例获取文档中的字段数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox "此文档中的字段数量为 " & 

 ActiveDocument.Fields.Count()
```
### Document.FullName 属性

此属性返回文档的文件名，包含完整路径。

**用法**

FullName As String

**参数**

无

**描述**

无

**示例**

以下示例获取打开的原理图名称和位置。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox "Hi, you are using " & ActiveDocument.FullName & " located in " & 

 ActiveDocument.Path

```

End Sub

**相关主题**

[Document.Name 属性](#page-84-0)

[Document.Path 属性](#page-89-0)

### Document.Gates 属性

此属性返回所有门电路的集合。

**用法**

Gates As [Objects 第219页](#page-194-0)

Gates(*name* As String) As [Gate 第155页](#page-130-0)

**参数**

| 参数 | 描述               |
|------|-------------------|
| name | 现有门电路的名称。 |

**描述**

当传入现有门电路名称时，返回对应的[门电路对象](#page-130-0)。如果门电路名称不存在，则返回[对象集合](#page-194-0)中的所有门电路。

**示例**

以下示例代码使用[Objects.Count](#page-196-0) [属性](#page-196-0)获取打开的原理图中的门电路数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set gates = ActiveDocument.Gates 

MsgBox "There are " & gates.Count & " gate(s) in " & ActiveDocument.Name

```

End Sub

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Sheet.GetObjects 方法](#page-258-0)

[Sheet.Gates 属性](#page-245-0)

### Document.GridX 属性

此属性设置或返回文档的X轴网格。

**用法**

GridX([*type* As [PlogGridType](#page-287-1) = plogGridDesign], [*unit* As [PlogUnit\]](#page-287-0)) As Double

**参数**

| 参数 | 描述                          |
|------|------------------------------|
| type | [可选 第314页] 要设置或返回的网格类型。 |
| unit | [可选] 设置或返回网格值的单位。 |

**示例**

以下示例代码将打开的原理图的显示网格设置为与设计网格相同的值。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

curGridX = ActiveDocument.GridX

curGridY = ActiveDocument.GridY

ActiveDocument.GridX(plogGridDisplay) = curGridX

ActiveDocument.GridY(plogGridDisplay) = curGridY


End Sub
```

![](/logic/scripts/2/_page_82_Picture_11.jpeg)

**提示** 不能为GridX和GridY分配不同的值。GridY属性不可写，因此GridY会采用分配给GridX的值。

**相关主题**

[Document.GridY 属性](#page-83-0)

### Document.GridY 属性

此属性设置或返回文档的Y轴网格。

**用法**

GridY([*type* As [PlogGridType](#page-287-1) = plogGridDesign], [*unit* A[sPlogUnit\]](#page-287-0)) As Double

**参数**

| 参数 | 描述                          |
|------|------------------------------|
| type | [可选 第314页] 要设置或返回的网格类型。 |
| unit | [可选] 设置或返回网格值的单位。 |

**示例**

以下示例代码将打开的原理图的显示网格设置为与设计网格相同的值。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

curGridX = ActiveDocument.GridX

curGridY = ActiveDocument.GridY

ActiveDocument.GridX(plogGridDisplay) = curGridX

ActiveDocument.GridY(plogGridDisplay) = curGridY

```

End Sub

![](/logic/scripts/2/_page_83_Picture_11.jpeg)

**提示** 不能为GridX和GridY分配不同的值。GridY属性不可写，因此GridY会采用分配给GridX的值。

**相关主题**

[Document.GridX 属性](#page-82-0)

### Document.Name 属性

此属性返回文档的名称。

**用法**

Name As String

**参数**

无

**描述**

例如，如果当前原理图文件是\My Documents\SailWind Projects\Samples\demo.sch，此函数返回字符串"demo.sch"。

此属性是Document对象的默认属性。

**示例**

以下示例代码获取打开的原理图名称和位置。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

MsgBox "Hi, you are using " & ActiveDocument.Name & " located in " & 

 ActiveDocument.Path

```

End Sub

**相关主题**

[Document.FullName 属性](#page-80-0)

[Document.Path 属性](#page-89-0)

### Document.Nets 属性

此属性返回所有网络的集合。

**用法**

Nets As [Objects 第219页](#page-194-0)

Nets(*name* As String) As [Net 第208页](#page-183-0)

**参数**

| 参数 | 描述             |
|------|----------------|
| name | 现有网络的名称。 |

**描述**

当传入现有网络名称时，返回对应的[网络对象](#page-183-0)。如果网络名称不存在，则返回[对象集合](#page-194-0)中的所有网络。

**示例**

以下示例代码使用[Objects.Count](#page-196-0) [属性](#page-196-0)获取打开的原理图中的网络数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set nets = ActiveDocument.Nets 

MsgBox "There are " & Nets.Count & " net(s) in " & ActiveDocument.Name 

End Sub

```

以下示例代码使用[Net.Pins 属性](#page-191-0)和[Objects.Count 属性](#page-196-0)获取网络VCC连接的引脚数量（假设打开的原理图中存在VCC网络）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set netVCC = ActiveDocument.Nets("VCC") 

MsgBox "Net " & netVCC.Name & " connects " & netVCC.Pins.Count & " 

 pin(s)." 

End Sub

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Sheet.Nets 属性](#page-247-0)

### Document.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Application 第401页

**参数**

无

**描述**

无

### Document.PartTypes 属性

此属性返回所有零件类型的集合。

**用法**

PartTypes As [Objects 第219页](#page-194-0)

PartTypes(name As String) As [PartType 第234页](#page-209-0)

**参数**

| 参数 | 描述                |
|------|-------------------|
| name | 现有零件类型的名称。 |

**描述**

当传入现有零件类型名称时，返回对应的[PartType 第77页](#page-52-0)对象。否则返回[对象集合](#page-194-0)中的所有零件类型。

**示例**

以下示例代码使用[Objects.Count 属性](#page-196-0)获取打开的原理图中的零件类型数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set pkgs = ActiveDocument.PartTypes 

MsgBox "There are " & pkgs.Count & " part type(s) in " &

```

End Sub

**相关主题**

[Document.GetObjects 方法](#page-99-0)

ActiveDocument.Name

### Document.Path 属性

此属性返回文档的路径。

**用法**

Path As String

**参数**

无

**描述**

无

**示例**

以下示例代码获取打开的原理图名称和路径。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

```

MsgBox "Hi, you are using " & ActiveDocument.Name & " located in " & 

 ActiveDocument.Path

```

End Sub

**相关主题**

[Document.FullName 属性](#page-80-0)

[Document.Name 属性](#page-84-0)

### Document.Pins 属性

此属性返回所有引脚的集合。

**用法**

Pins As [Objects 第219页](#page-194-0)

Pins(*name* As String) As [Pin 第246页](#page-221-0)

**参数**

|  | 参数 | 描述             |
|--|------|-----------------|
|  | name | 现有引脚的名称 |

**描述**

当传入现有引脚名称时，返回该[Pin 对象](#page-221-0)。否则返回包含所有引脚的[Objects 集合对象](#page-194-0)。

**示例**

以下示例代码使用[Objects.Count](#page-196-0)[属性](#page-196-0)获取打开的原理图中的引脚数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

Set pins = ActiveDocument.Pins

MsgBox "There are " & Pins.Count & " pin(s) in " & ActiveDocument.Name

```

End Sub

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Sheet.Nets 属性](#page-247-0)

### Document.Saved 属性

此属性设置或返回文档是否已保存。

**用法**

Saved As Boolean

**参数**

无

**描述**

此属性通常在SailWind Logic中使用[Application.OpenDocument 方法](#page-24-0)打开新原理图前使用，以避免出现*重新加载前保存旧文件？*的提示消息。

**示例**

以下示例代码将打开的原理图保存状态设为True，然后打开demo.sch（假设该文件存在于[Application.DefaultFilePath 属性](#page-4-0)指向的文件夹中），并获取新打开的原理图名称。有关运行此示例的更多信息，请参见代码示例。

Sub Main ActiveDocument.Saved = True OpenDocument(DefaultFilePath & "\DEMO.SCH") MsgBox ActiveDocument.FullName & " has just been opened."

End Sub

**相关主题**

[Document.Save 方法](#page-105-0)

[Document.SaveAs 方法](#page-106-0)

### Document.Sheets 属性

此属性返回所有图纸的集合。

**用法**

Sheets As [Sheets 第265页](#page-240-0)

Sheets(*name* As String) As [Sheet 第265页](#page-240-0)

**参数**

| 参数 | 描述                |
|------|--------------------|
| name | 现有图纸的名称。 |

**描述**

当传入现有图纸名称时，返回打包为[Sheet](#page-240-0)[对象](#page-240-0)的该图纸。否则返回打开的原理图中所有图纸的集合，打包为[Sheets 集合对象](#page-261-0)。

**示例**

以下示例代码使用[Sheets.Count](#page-263-0)[属性](#page-263-0)获取打开的原理图中的图纸数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

Set shts = ActiveDocument.Sheets

MsgBox "There are " & shts.Count & " sheet(s) in " & ActiveDocument.Name

```

### Document.Activate 方法

此方法激活与文档关联的窗口。

**用法**

Activate ()

**参数**

无

**返回值**

无

**描述**

这是Microsoft的要求。但由于SailWind Logic是SDI（单文档界面）服务器应用程序，此函数无效。

### Document.ExportASCII 方法

此方法从当前原理图生成SailWind Logic ASCII文件。

**用法**

ExportASCII(*path* As String)

**参数**

| 参数 | 描述              |
|------|------------------|
| path | 输出文件的路径。 |

**描述**

无

**示例**

以下示例从活动文档创建ASCII文件。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

ActiveDocument.ExportASCII DefaultFilePath & "\ascii.txt"

### Document.ExportNetList 方法

此方法从当前原理图生成PADS格式网表。

**用法**

ExportNetList(*path* As String, [*ver* As [PlogNetListVersion\]](#page-288-1))

**参数**

|  |  | 参数 | 描述                                                                    |
|--|--|------|-----------------------------------------------------------------------|
|  |  | path | 要导出网表的文件名。                                                  |
|  |  | ver  | [可选 第314页] 导出版本（用于向后兼容）。 |

**返回值**

无

**描述**

如果指定文件不存在，则创建新文件。如果文件已存在，则覆盖现有文件。

如果函数失败，此属性会生成[异常](#page-290-0)。

**示例**

以下示例代码创建具有指定名称(padsnet.asc)的网表文件。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

ActiveDocument.ExportNetList(DefaultFilePath & "\padsnet.asc")

### Document.GenerateECO 方法

此方法将打开的原理图与指定的网表或其他原理图进行比较并生成ECO文件。

**用法**

GenerateECO(*fileToComp* As String, *path* As String)

**参数**

| 参数       | 描述                                                               |
|------------|-------------------------------------------------------------------|
| fileToComp | 要与打开的原理图进行比较的原理图或网表文件名。 |
| path       | 生成的ECO文件名。                                           |

**返回值**

无

**描述**

无

**示例**

以下示例代码比较PADS格式网表文件(padsnet.asc)并创建具有指定名称(eco2pcb.eco)的ECO文件。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

ActiveDocument.GenerateECO(DefaultFilePath & "\padsnet.asc", 

 DefaultFilePath & "\eco2pcb.eco")

```

### Document.GetColor 方法

此方法返回指定文档元素的颜色。

**用法**

GetColor(*colorType* as PlogDocumentColor) as Integer

**参数**

| 参数      | 描述                     |
|-----------|-------------------------|
| colorType | 指定文档元素。 |

**返回值**

调色板中的颜色索引。

**示例**

此示例还包括SetColor方法。

```

doc = Application.ActiveDocument

```

```

msg=""

msg = msg & doc.GetColor(plogDocumentColorBackground) & ", "

msg = msg & doc.GetColor(plogDocumentColorSelection) & ", "

msg = msg & doc.GetColor(plogDocumentColorConnection) & ", "

msg = msg & doc.GetColor(plogDocumentColorBus) & ", "

msg = msg & doc.GetColor(plogDocumentColorLine) & ", "

msg = msg & doc.GetColor(plogDocumentColorPart) & ", "

msg = msg & doc.GetColor(plogDocumentColorHierarchicalComp) & ", "

msg = msg & doc.GetColor(plogDocumentColorText) & ", "

msg = msg & doc.GetColor(plogDocumentColorTextBox) & ", "

```

|  |  | msg = msg & doc.GetColor(plogDocumentColorRefDes) & ", "       |
|--|--|----------------------------------------------------------------|
|  |  | msg = msg & doc.GetColor(plogDocumentColorRefDesBox) & ", "    |
|  |  | msg = msg & doc.GetColor(plogDocumentColorPartType) & ", "     |
|  |  | msg = msg & doc.GetColor(plogDocumentColorPartTypeBox) & ", "  |
|  |  | msg = msg & doc.GetColor(plogDocumentColorPartText) & ", "     |
|  |  | msg = msg & doc.GetColor(plogDocumentColorPartTextBox) & ", "  |
|  |  | msg = msg & doc.GetColor(plogDocumentColorPinNumber) & ", "    |
|  |  | msg = msg & doc.GetColor(plogDocumentColorPinNumberBox) & ", " |
|  |  | msg = msg & doc.GetColor(plogDocumentColorNetName) & ", "      |
|  |  | msg = msg & doc.GetColor(plogDocumentColorNetNameBox) & ", "   |
|  |  | msg = msg & doc.GetColor(plogDocumentColorField) & ", "        |
|  |  | msg = msg & doc.GetColor(plogDocumentColorFieldBox)            |

MsgBox msg

curr\_bkg\_color = doc.GetColor(plogDocumentColorBackground)

doc.SetColor(plogDocumentColorBackground, 2)

MsgBox "Press any key"

doc.SetColor(plogDocumentColorBackground, curr\_bkg\_color)

### Document.GetObjects 方法

该方法返回一个 SailWind Logic 数据库对象集合。

**用法**

GetObjects([*type* As [PlogObjectType](#page-286-0) = plogObjectTypeAll], [*value* As String], [*selected* As Boolean = False]) As [Objects](#page-194-0)

**参数**

| 参数     | 描述                                                                 |  |
|----------|---------------------------------------------------------------------|--|
| type     | [可选](#page-314) 要获取的 SailWind Logic 数据库对象类型。           |  |
| name     | [可选] 要获取的对象的值或名称。                                      |  |
| selected | [可选] True 表示仅获取选中✅的对象。False 表示获取所有对象。           |  |

**返回值**

返回的对象是一个 [Objects 集合对象](#page-194-0)。如果没有对象满足请求，则返回空集合。

**描述**

该方法的所有参数都是可选的，这意味着可以不带任何参数调用它，或者使用任意参数组合调用。更多信息请参见下面的示例。

*Name* 参数支持通配符 ("U*")、逗号分隔的项目列表 ("U1, U2, R1")、以及由两个对象名称和短划线字符指定的范围 ("U1 - U10, U12, R1 - R20")。短划线必须用空格包围，因为短划线是对象名称中的合法符号。每个名称只允许一个通配符，不能在范围中使用通配符。可以传递 *name* 如 "U*, R*, C1 – C100"，但不能传递 *name* 如 "U*1*" 或 "C1* - C10*"。

要获取相同类型的所有对象，请使用相应的对象 Document 属性而不是此方法。例如，要获取打开的原理图中的所有门，请使用 [Document.Gates 属性](#page-81-0) 而不是 Document.GetObjects(plogObjectTypeGate)。

如果 *type* 参数不是有效的 SailWind Logic 数据库对象类型，此属性会生成一个 [异常](#page-290-0)。

**示例**

以下示例代码展示了使用此方法的不同方式，使用 [Objects.Count 属性](#page-196-0) 显示每种方式检索到的对象数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

```

Dim objs As Object

```

' 示例1：获取所有类型的对象

```

```

Set objs = ActiveDocument.GetObjects

```

MsgBox "示例1: " & objs.Count & " 个对象。"

' 示例2：获取所有选中✅的对象

Set objs = ActiveDocument.GetObjects(,,True)

MsgBox "示例2: " & objs.Count & " 个选中✅的对象。"

' 示例3：获取所有网络对象

Set objs = ActiveDocument.GetObjects(plogObjectTypeNet)

MsgBox "示例3: " & objs.Count & " 个网络对象。"

' 示例4：获取所有名称为"VCC"的网络对象（至少有1个）

Set objs = ActiveDocument.GetObjects(plogObjectTypeNet, "VCC")

MsgBox "示例4: " & objs.Count & " 个VCC网络对象。"

' 示例5：获取所有名称以U开头的部件对象

Set objs = ActiveDocument.GetObjects(plogObjectTypeComponent, "U*")

MsgBox "示例3: " & objs.Count & " 个U*部件对象。"

End Sub

**相关主题**

[Document.SelectObjects 方法](#page-111-0)

[Document.Components 属性](#page-77-0)

[Document.Gates 属性](#page-81-0)

[Document.Nets 属性](#page-85-0)

[Document.Pins 属性](#page-90-0)

[Document.PartTypes](#page-88-0) 属性

[Sheet.GetObjects 方法](#page-258-0)

[Sheet.Components 属性](#page-243-0)

[Sheet.Gates 属性](#page-245-0)

[Sheet.Nets 属性](#page-247-0)

[Sheet.PartTypes](#page-251-0) 属性

### Document.ImportASCII 方法

该方法导入一个 SailWind Logic ASCII 文件。

**用法**

ImportASCII(*name* As String)

**参数**

| 参数 | 描述                     |
|------|-------------------------|
| name | 要导入的现有ASCII文件名。 |

**描述**

如果指定的文件名不存在或其格式不正确，该函数将失败并生成一个 [异常](#page-290-0)。

**示例**

以下示例代码将指定文件(demo.txt)导入到当前原理图中。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

ActiveDocument.ImportASCII(DefaultFilePath & "\demo.txt")

End Sub

**相关主题**

[Document.ExportNetList 方法](#page-95-0)

[Document.ExportASCII 方法](#page-94-0)

### Document.ImportECO 方法

该方法导入一个ECO文件。

**用法**

ImportECO(*name* As String)

**参数**

| 参数 | 描述                   |
|------|-----------------------|
| name | 要导入的现有ECO文件名。 |

**描述**

如果指定的文件名不存在或其格式不正确，该函数将失败并生成一个 [异常](#page-290-0)。

**示例**

以下示例代码将指定的ECO文件(*eco2sch.eco*)导入到当前原理图中。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

ActiveDocument.ImportECO(DefaultFilePath & "\eco2sch.eco")

End Sub

**相关主题**

[Document.ExportNetList 方法](#page-95-0)

### Document.IntegrityTest 方法

该方法运行完整性测试。

**用法**

IntegrityTest() as Boolean

**参数**

无

**返回值**

如果测试通过且无错误返回TRUE，其他情况返回FALSE。

### Document.Save 方法

该方法保存文档。

**用法**

Save

**参数**

无

**描述**

无

**示例**

以下示例代码在需要时保存打开的原理图。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

If ActiveDocument.Saved = False Then ActiveDocument.Save

End Sub

**相关主题**

[Document.SaveAs 方法](#page-106-0)

[Document.Saved 属性](#page-91-0)

[Document.Save 事件](#page-116-0)

### Document.SaveAs 方法

该方法以新名称保存文档。

**用法**

Save(*name* As String)

**参数**

| 参数 | 描述         |
|------|-------------|
| name | 新文件名。   |

**示例**

以下示例代码为打开的原理图创建自定义备份。如果原理图文件名为*XXX.SCH*，则备份文件将以XXX (Backup on 12-28-1998 4h40).SCH的名称保存在同一文件夹中。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

```

```

curSchematicName = ActiveDocument.FullName

```

```

theDate = Month(Date) & "-" & Day(Date) & "-" & Year(Date) & " at " & 

 Hour(Now) & "h" & Minute(Now)

```

```

bakSchematicName = Left$(curSchematicName, Len(curSchematicName)-4) & 

 " (Backup on " & theDate & ").sch"

```

ActiveDocument.SaveAs(bakSchematicName)

OpenDocument(curSchematicName)

End Sub

**相关主题**

[Document.Save 方法](#page-105-0)

[Document.Saved 属性](#page-91-0)

[Document.Save 事件](#page-116-0)

### Document.SaveAsNoLock 方法

此方法使用新名称保存文档并释放文件锁。

**用法**

SaveAsNoLock(*filename* as String)

**参数**

| 参数      | 描述            |
|-----------|-----------------|
| filename  | 文件的新名称。  |

**返回值**

### Document.SaveAsTemp 方法

此方法与 SaveAsNoLock 工作方式相同，但不会将文件名添加到 MRU (最近使用) 列表中。此方法主要用于宏测试。

**用法**

SaveAsTemp(*filename* as String)

**参数**

| 参数      | 描述            |
|-----------|-----------------|
| filename  | 文件的新名称。  |

**返回值**

### Document.SaveNoLock 方法

此方法保存文档并释放文件锁。

**用法**

SaveNoLock

**参数**

无

**返回值**

### Document.SaveTemp 方法

此方法与 SaveNoLock 工作方式相同，但不会将文件名添加到 MRU (最近使用) 列表中。此方法主要用于宏测试。

**用法**

SaveTemp

**参数**

无

**返回值**

### Document.SelectObjects 方法

此方法选择或取消选择 SailWind Logic 数据库对象。

**用法**

```

SelectObjects([type As PlogObjectType = plogObjectTypeAll], [name As String], [select As Boolean = 

      True])

```

**参数**

| 参数    | 描述                                                                      |
|---------|--------------------------------------------------------------------------|
| type    | [可选，见第314页] 要选择/取消选择的 SailWind Logic 数据库对象类型。       |
| name    | [可选] 要选择/取消选择的对象的值或名称。                                  |
| select  | [可选] True 表示选择，False 表示取消选择。                                |

**描述**

此方法的所有参数都是可选的，这意味着可以不带任何参数调用它，也可以使用任意参数组合调用。有关更多信息，请参阅下面的示例。

*Name* 参数支持通配符 ("U\*")、逗号分隔的项目列表 ("U1, U2, R1") 以及由两个对象名称和短划线字符指定的范围 ("U1 - U10, U12, R1 - R20")。短划线必须用空格包围，因为短划线是对象名称中的合法符号。每个名称只允许一个通配符，不能在范围中使用通配符。可以传递 *name* 如 "U\*, R\*, C1 – C100"，但不能传递 *name* 如 "U\*1\*" 或 "C1\* - C10\*"。

如果 *type* 参数不是有效的 SailWind Logic 数据库对象类型，此属性会生成[异常](#page-290-0)。

**示例**

以下示例代码展示了使用此方法的不同方式，使用 [Document.GetObjects 方法](#page-99-0) 和 [Objects.Count 属性](#page-196-0) 属性显示每种方式选择的对象数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Dim objs As Object 

' 示例1: 选择所有类型的所有对象 

ActiveDocument.SelectObjects 

Set objs = ActiveDocument.GetObjects(,,True)

```

MsgBox "示例1: " & objs.Count & " 个选中✅对象(全部)。"

' 示例2: 取消选择所有类型的所有对象

ActiveDocument.SelectObjects(,,False)

Set objs = ActiveDocument.GetObjects(,,True)

MsgBox "示例2: " & objs.Count & " 个选中✅对象(无)。"

' 示例3: 选择所有网络对象

ActiveDocument.SelectObjects(plogObjectTypeNet)

Set objs = ActiveDocument.GetObjects(,,True)

MsgBox "示例3: " & objs.Count & " 个选中✅对象(所有网络)。"

' 示例4: 取消选择网络 VCC

ActiveDocument.SelectObjects(plogObjectTypeNet, "VCC", False)

Set objs = ActiveDocument.GetObjects(,,True)

MsgBox "示例4: " & objs.Count & " 个选中✅对象(除VCC外的所有网络)。"

' 示例5: 仅选择名称以U开头的元件对象

ActiveDocument.SelectObjects(,,False)

ActiveDocument.SelectObjects(plogObjectTypeComponent, "U\*")

Set objs = ActiveDocument.GetObjects(,,True)

MsgBox "示例5: " & objs.Count & " 个选中✅的U\*元件对象。"

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Document.SelectionChange 事件](#page-117-0)

### Document.SetColor 方法

此方法为指定的文档元素设置颜色。

**用法**

SetColor(*colorType* as PlogDocumentColor, *colorIndex* as Integer)

**参数**

| 参数        | 描述                                                      |
|-------------|----------------------------------------------------------|
| colorType   | 指定文档元素。                                            |
| colorIndex  | 调色板中的颜色索引。必须介于0到31之间。                   |

**示例**

此示例还包括 GetColor 方法。

| doc = Application.ActiveDocument                                   |  |  |
|--------------------------------------------------------------------|--|--|
| msg=""                                                             |  |  |
| msg = msg & doc.GetColor(plogDocumentColorBackground) & ", "       |  |  |
| msg = msg & doc.GetColor(plogDocumentColorSelection) & ", "        |  |  |
| msg = msg & doc.GetColor(plogDocumentColorConnection) & ", "       |  |  |
| msg = msg & doc.GetColor(plogDocumentColorBus) & ", "              |  |  |
| msg = msg & doc.GetColor(plogDocumentColorLine) & ", "             |  |  |
| msg = msg & doc.GetColor(plogDocumentColorPart) & ", "             |  |  |
| msg = msg & doc.GetColor(plogDocumentColorHierarchicalComp) & ", " |  |  |
| msg = msg & doc.GetColor(plogDocumentColorText) & ", "             |  |  |
| msg = msg & doc.GetColor(plogDocumentColorTextBox) & ", "          |  |  |
| msg = msg & doc.GetColor(plogDocumentColorRefDes) & ", "           |  |  |

| msg = msg & doc.GetColor(plogDocumentColorRefDesBox) & ", "    |
|----------------------------------------------------------------|
| msg = msg & doc.GetColor(plogDocumentColorPartType) & ", "     |
| msg = msg & doc.GetColor(plogDocumentColorPartTypeBox) & ", "  |
| msg = msg & doc.GetColor(plogDocumentColorPartText) & ", "     |
| msg = msg & doc.GetColor(plogDocumentColorPartTextBox) & ", "  |
| msg = msg & doc.GetColor(plogDocumentColorPinNumber) & ", "    |
| msg = msg & doc.GetColor(plogDocumentColorPinNumberBox) & ", " |
| msg = msg & doc.GetColor(plogDocumentColorNetName) & ", "      |
| msg = msg & doc.GetColor(plogDocumentColorNetNameBox) & ", "   |
| msg = msg & doc.GetColor(plogDocumentColorField) & ", "        |
| msg = msg & doc.GetColor(plogDocumentColorFieldBox)            |
|                                                                |

MsgBox msg

curr\_bkg\_color = doc.GetColor(plogDocumentColorBackground)

doc.SetColor(plogDocumentColorBackground, 2)

MsgBox "按任意键"

doc.SetColor(plogDocumentColorBackground, curr\_bkg\_color)

### Document.Save 事件

此事件在文档保存后发生。

**用法**

Document\_Save ()

**参数**

无

**描述**

此事件在文档保存后发生。

**相关主题**

[Document.Save 方法](#page-105-0)

[Document.SaveAs 方法](#page-106-0)

### Document.SelectionChange 事件

此事件在当前选择发生变化时发生。

**用法**

Document\_SelectionChange ()

**参数**

无

**描述**

此事件在当前选择发生变化时发生。

**相关主题** [Document.SelectObjects 方法](#page-111-0) [Objects.Select 方法](#page-206-0)

## Field 对象

Field 对象包含原理图文件所需的操作系统属性。此对象使您可以检索字段名称、值和父级的信息。

### Field.Application 属性

此属性返回 Application 对象。

**用法**

Application As SailWind Logic.Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为自动化对象。这是 Microsoft 要求的属性。

**示例**

以下示例代码使用集合中的第一个字段检索应用程序的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

MsgBox ActiveDocument.Fields(1).Application.Name

### Field.Name 属性

此属性返回字段对象的名称。

**用法**

Name As String 见第401页

**参数**

无

**描述**

这是 Microsoft 要求的属性。

**示例**

以下示例代码获取字段集合中第一个字段的名称。有关运行此示例的更多信息，请参见第21页的"代码示例"。

MsgBox ActiveDocument.Fields(1).Name

### Field.Parent 属性

此属性返回字段的父对象。父对象是当前 Document 对象。

**用法**

Parent As Object 见第401页

**参数**

无

**描述**

这是 Microsoft 要求的属性。

**示例**

以下示例代码使用集合中的第一个字段获取父对象的名称。有关运行此示例的更多信息，请参见第21页的"代码示例"。

MsgBox ActiveDocument.Fields(1).Parent.Name

### Field.Value 属性

此默认属性设置或返回 Application 字段的值。

**用法**

Value As Variant 见第401页

**参数**

无

**描述**

此属性使您可以查看或设置 Application 字段的值。

![](/logic/scripts/2/_page_122_Picture_9.jpeg)

**提示** 您既不能删除也不能更改系统字段的值。当您对字段执行操作时，它会影响当前原理图中引用该字段的所有字段标签。

**示例**

以下示例代码设置字段的值。

**提示** FieldName 字段应存在于文档中。

ActiveDocument.Fields ("FieldName").Value="FieldValue"

有关运行这些示例的更多信息，请参见第21页的"代码示例"。

以下示例代码使用默认 Value 属性设置字段的值。

```

ActiveDocument.Fields ("FieldName")="FieldValue"

```

## Fields Collection 对象

Fields Collection 对象包含属性和方法，用于向打开的示意图中的集合添加和删除字段。

通常使用 [Application.ActiveDocument 属性](#page-2-0)对象检索此对象。

**提示：**

- 当您对字段执行操作时，它会影响当前原理图中引用该字段的所有字段标签。

- 您既不能删除也不能更改系统字段的值。

### Fields.Application 属性

此属性返回 Application 对象。

**用法**

Application As SailWind Logic.Application 见第401页

**参数**

无

**描述**

此属性将对象标识为 Automation 对象。这是 Microsoft 要求的属性。

**示例**

以下示例代码使用字段集合对象获取应用程序的名称。有关运行此示例的更多信息，请参见第21页的"代码示例"。

MsgBox ActiveDocument.Fields.Application.Name

### Fields.Count 属性

此属性返回当前 Document 中字段集合的字段数量。

**用法**

Count As Long 见第401页

**参数**

无

**描述**

此属性返回当前 Document 中字段集合的字段数量。这是 Microsoft 要求的属性。

**示例**

以下示例代码显示文档中的字段数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

MsgBox ActiveDocument.Fields.Count

### Fields.Item 属性

此默认属性返回字段集合的成员。可以通过位置或名称返回成员。

**用法**

Item as SailWind Logic.Field 见第401页

**参数**

| 参数 | 描述 |
|----------|------------------------------------------------------------------------------------------------------------------|
| Index | 必需的变体，可以有以下两种值 |
| Value | • 字段的字符串名称<br>• 从1到集合 Count 属性值的数字 |

**描述**

此属性返回原理图中字段集合的成员。这是集合对象的 Microsoft 要求属性。

![](/logic/scripts/2/_page_126_Picture_9.jpeg)

**提示** 您既不能删除也不能更改系统字段的值。

**示例**

以下示例代码显示打开的示意图中给定字段的值。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

MsgBox ActiveDocument.Fields.Item

```

("FieldName").Value

以下示例使用默认 item 属性显示给定字段的值。有关运行此示例的更多信息，请参见第21页的"代码示例"。

MsgBox ActiveDocument.Fields("AttrName").Value

### Fields.Parent 属性

此属性返回字段集合对象的父 Application。

**用法**

Parent As Object 见第401页

**参数**

无

**描述**

此属性返回字段集合对象的父对象。父对象是文档。这是 Microsoft 要求的属性。

**示例**

以下示例代码获取字段集合对象父对象的名称。（字段集合对象的父对象是当前文档。）有关运行此示例的更多信息，请参见第21页的"代码示例"。

ActiveDocument.Fields.Parent.Name

### Fields.Add 方法

此方法向当前 Document 的字段集合添加新字段，并返回一个新字段对象。

**用法**

Add(*name* As String, value As Variant) As [SailWind Logic.Field 见第143页](#page-118-0)

**参数**

| 参数 | 描述 |
|----------|--------------------------------------------------|
| Value | 指定字段值的变体。 |

**描述**

如果 *name* 参数是现有字段或不是有效的字段名称，此方法会生成一个[异常](#page-290-0)。

**示例**

以下示例代码向文档添加新字段。有关运行此示例的更多信息，请参见第21页的"代码示例"。

ActiveDocument.Fields.Add "FieldName","StringValue"

**相关主题**

[Fields.Delete 方法](#page-129-0)

### Fields.Delete 方法

此方法按索引从字段集合中删除字段。

**用法**

Delete(*index* As Variant) As [SailWind Logic.Field 见第143页](#page-118-0)

**参数**

| 参数 | 描述 |
|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Index | 必需的变体，可以有以下两种值：<br>• 字段的字符串名称<br>• 从1到集合 Count 属性值的数字 |

**描述**

如果 *index* 参数超出限制、不是有效的索引名称或指向系统字段，此方法会生成一个[异常](#page-290-0)。

**示例**

以下示例代码按名称删除字段。有关运行此示例的更多信息，请参见第21页的"代码示例"。

ActiveDocument.Fields.Delete("FieldName")

**相关主题**

[Fields.Add 方法](#page-128-0)

## Gate 对象

Gate 对象表示存在于打开的示意图中的物理门。门可以是使用中或未使用的。如果门未使用，其 Sheet 属性为 *Nothing*。

以下按钮列出了 Gate 对象中的属性。

### Gate.Application 属性

该属性返回 SailWind Logic 应用对象。

**用法**

Application As Application 参见第401页

**参数**

无

**描述**

此属性标识对象为 SailWind Logic 自动化对象。所有自动化服务器应用都有一个 Application 对象，所有自动化对象都有一个 Application 属性。该属性通常用于处理来自不同源（如不同的自动化服务器应用）的大量对象的大型自动化客户端应用中。使用该属性可快速识别对象所属的应用。

### Gate.Component 属性

该属性返回该门所属的组件。

**用法**

Component As [Component 参见第77页](#page-52-0)

**参数**

无

**示例**

以下示例代码检索门 U6-A 所属的组件（假设它存在于打开的电路图中）。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

  OpenDocument DefaultFilePath & "\preview.sch"

  MsgBox "门 U6-A 属于组件 " & ActiveDocument.Gates("U6-A").Component.Name

### Gate.Name 属性

该属性返回门的名称。

**用法**

Name As String

**参数**

无

**描述**

此属性是 Gate 对象的默认属性。

**示例**

以下示例代码按名称检索活动图纸中的所有门，并在自定义对话框中列出它们。当在列表框中选择一个门时，示例会在 SailWind Logic 中选择该门。

此示例使用 SailWind Logic 中 Sax Basic 引擎的 UserDialog 编辑器。更多信息请参见 Sax Basic 编辑器在线帮助。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Dim ListGates\$(10000) 

Sub Main 

  index = 0 

  for Each nextGate In ActiveDocument.ActiveSheet.Gates 

    ListGates\$(index) = nextGate.Name 
    
    index = index + 1 

  Next nextGate 

  ' 这段代码由 SailWind Logic Basic 对话框编辑器自动生成

  Begin Dialog UserDialog 180,238,"Gates",.CallbackFunc 

  ' %GRID:10,7,1,1 

  ListBox 10,7,160,203,ListGates(),.ListBox1

  OKButton 10,210,160,21

  End Dialog

  Dim dlg As UserDialog

  Dialog dlg

End Sub

' 当对话框中发生某些操作时，系统会自动调用以下函数

' 用于轻松处理用户操作

Function CallbackFunc%(DlgItem\$, Action%, SuppValue%)

  Select Case Action%

    Case 2 ' 值改变或按钮按下
    
      If DlgItem\$="ListBox1" Then
    
        ActiveDocument.SelectObjects(plogObjectTypeAll, , False)
    
        ActiveDocument.SelectObjects(plogObjectTypeGate, ListGates(SuppValue%))
    
      End If

  End Select

End Function

### Gate.Number 属性

该属性返回此门在部件中的索引。

**用法**

Number As Long

**参数**

无

**描述**

U1-A 门的索引号为1；U1-B 为2；以此类推。

**示例**

以下示例代码检索门 U6-A 的索引（假设它存在于打开的电路图中）。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main 

  OpenDocument DefaultFilePath & "\preview.sch" 

  Set U6A = ActiveDocument.Gates("U6-A") 

  MsgBox "门 U6-A 在组件 " & U6A.Component.Name & " 中的门索引为 " & U6A.Number

### Gate.ObjectType 属性

该属性返回对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

此属性返回 [PlogObjectTypeG](#page-286-0)ate。

SailWind Logic 自动化服务器中的所有 SailWind Logic 数据库对象都实现了此属性，以弥补 Basic 中没有 Visual C++ QueryInterface 函数等效功能的问题。

此属性通常用于：

- 在异构的 [Objects Collection](#page-194-0) [Object](#page-194-0) 中识别 SailWind Logic 数据库对象的种类

- 当实现一个依赖于作为参数传递的 SailWind Logic 数据库对象类型的通用例程时。例如：

Sub DoSomething(dbObject As Object)

  Select Case dbObject.ObjectType

    Case plogObjectTypeComponent
    
      ' 对组件对象执行特定操作
    
    Case plogObjectTypeNet
    
      ' 对网络对象执行特定操作
    
    Case plogObjectTypePin
    
      ' 对引脚对象执行特定操作
    
    Case plogObjectTypeGate
    
      ' 对门对象执行特定操作
    
    Case Else 
    
      MsgBox "不是 SailWind Logic 数据库对象" 

  End Select 

End Sub

### Gate.Parent 属性

该属性返回对象的父对象。

**用法**

Parent As Document 参见第419页

**参数**

无

**描述**

无

### Gate.Pins 属性

该属性返回门的所有引脚的集合。

**用法**

Pins As [Objects 参见第219页](#page-194-0)

Pin(*name* As String) As [Pin 参见第246页](#page-221-0)

**参数**

| 参数 | 描述              |
|------|-------------------|
| name | 现有引脚的名称。  |

**描述**

当传递现有引脚 *name* 给此属性时，它返回该 [Pin 对象](#page-221-0)。否则，它返回门的所有引脚的集合，以 [Objects Collection 对象](#page-194-0)形式。

**示例**

以下示例检索门 U6-A 中的引脚总数。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

  OpenDocument DefaultFilePath & "\preview.sch"

  ```

  MsgBox "门 U6A 中的引脚总数为 " & 

  ActiveDocument.Gates("U6-A").Pins.Count

  ```

### Gate.PositionX 属性

该属性返回组件的 x 坐标。

**用法**

PositionX([*unit* As [PlogUnit](#page-287-0)]) As Double

**参数**

| 参数 | 描述                                  |
|------|---------------------------------------|
| unit | [可选 参见第314页] 返回 x 坐标的单位。 |

**描述**

无

**示例**

以下示例代码检索门 U6-A 在当前设计单位中的位置。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

  OpenDocument DefaultFilePath & "\preview.sch"

  ```

  Set U6A = ActiveDocument.Gates("U6-A")

  ```

  MsgBox "U6-A 的位置是 (" & U6A.PositionX & ", " & U6A.PositionY & ")"

### Gate.PositionY 属性

该属性返回组件的 y 坐标。

**用法**

PositionY([*unit* As [PlogUnit](#page-287-0)]) As Double

**参数**

| 参数 | 描述                                  |
|------|---------------------------------------|
| unit | [可选 参见第314页] 返回 y 坐标的单位。 |

**描述**

无

**示例**

以下示例代码检索门 U6-A 在当前设计单位中的位置。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

  OpenDocument DefaultFilePath & "\preview.sch"

  ```

  Set U6A = ActiveDocument.Gates("U6-A")

  ```

  MsgBox "U6-A 的位置是 (" & U6A.PositionX & ", " & U6A.PositionY & ")"

### Gate.ReflectedX 属性

该属性设置镜像或返回门电路是否沿X轴镜像。

**用法**

ReflectedX As Boolean

**参数**

无

**描述**

如果门电路未使用，该属性会生成[异常](#page-290-0)。

**示例**

以下示例旋转并镜像门电路U6-A。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 OpenDocument DefaultFilePath & "\preview.sch"

 Set U6A = ActiveDocument.Gates("U6-A")

 U6A.ReflectedX = True

 U6A.ReflectedY = True

 U6A.Rotated90 = True

```

End Sub

**相关主题**

[Gate.Rotated90 属性](#page-144-0)

[Gate.ReflectedY 属性](#page-143-0)

### Gate.ReflectedY 属性

该属性设置镜像或返回门电路是否沿Y轴镜像。

**用法**

ReflectedY As Boolean

**参数**

无

**描述**

如果门电路未使用，该属性会生成[异常](#page-290-0)。

**示例**

以下示例旋转并镜像门电路U6-A。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main OpenDocument DefaultFilePath & "\preview.sch" Set U6A = ActiveDocument.Gates("U6-A") U6A.ReflectedX = True U6A.ReflectedY = True U6A.Rotated90 = True

End Sub

**相关主题**

[Gate.Rotated90 属性](#page-144-0)

[Gate.ReflectedX 属性](#page-142-0)

### Gate.Rotated90 属性

该属性设置旋转或返回门电路是否旋转。

**用法**

Rotated90 As Boolean

**参数**

无

**描述**

如果门电路未使用，该属性会生成[异常](#page-290-0)。

**示例**

以下示例旋转并镜像门电路U6-A。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 OpenDocument DefaultFilePath & "\preview.sch"

 Set U6A = ActiveDocument.Gates("U6-A")

 U6A.ReflectedX = True

 U6A.ReflectedY = True

 U6A.Rotated90 = True

```

End Sub

**相关主题**

[Gate.ReflectedX 属性](#page-142-0)

[Gate.ReflectedY 属性](#page-143-0)

### Gate.Selected 属性

该属性设置或返回门电路是否被选中✅。

**用法**

Selected As Boolean

**参数**

无

**描述**

您也可以使用[Document.SelectObjects 方法](#page-111-0)或[Objects.Select 方法](#page-206-0)选择SailWind Logic数据库对象。

**示例**

以下示例代码选中✅一个门电路，假设它存在于打开的电路图中。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

ActiveDocument.SelectObjects(,,False)

```

ActiveDocument.Gates("U1-A").Selected = True

End Sub

**相关主题**

[Document.SelectionChange 事件](#page-117-0)

### Gate.Sheet 属性

该属性返回门电路所在的工作表。

**用法**

Sheet As [Sheeton page 265](#page-240-0)

**参数**

无

**描述**

如果门电路未使用，返回值为*Nothing*。您可以使用此属性检查门电路是否未使用。

**示例**

以下示例检索组件U1的第一个门电路，并检索门电路所在工作表的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main Set gate = ActiveDocument.Components("U1").Gate(1) If gate.Sheet Is Nothing then MsgBox "此门电路未使用！" Else MsgBox "部件U1的第一个门电路位于工作表 " & gate.Sheet

End If

### Gate.SwapClass 属性

该属性返回此门电路的交换类。

**用法**

SwapClass As Long

**参数**

无

**描述**

无

**示例**

以下示例代码检索门电路U6-A的交换类，假设它存在于打开的电路图中。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main OpenDocument DefaultFilePath & "\preview.sch" Set U6A = ActiveDocument.Gates("U6-A") MsgBox "门电路U1-A的交换类是 " & U6A.SwapClass

End Sub

**相关主题**

[Gate.Number 属性](#page-135-0)

### Gate.Visibility 属性

该属性设置可见性或返回门电路是否可见。

**用法**

Visibility([*Item* As [PlogGateVisibility](#page-288-2) = plogAttrVisibility], [*AttrName* As String=" "]) As Boolean

PlogGateVisibility 项。可能值为：

plogAttrVisibility = 0

plogAttrNameVisibility = 1

plogRefDesVisibility = 2

plogPartTypeVisibility = 3

plogPinNumberVisibility = 4

plogPinNameVisibility = 5

plogPCBDecalVisibility = 6

plogPCBDecalNameVisibility = 7

**参数**

| 参数 | 描述 |
|----------|----------------------------------------------------------------------------------------------------------------------------------------------|
| Item | [可选on page 314] 可见性项。默认情况下，它指定AttrName中指定的属性的可见性。 |
| AttrName | [可选] 属性名称。可与plogAttrVisibility和plogAttrNameVisibility项一起使用。其他项应省略。 |

**描述**

此函数不会重绘门电路。使用[View.Refresh](#page-280-0)方法更新屏幕。

**示例**

以下示例代码设置门电路U6-A的可见性。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 OpenDocument DefaultFilePath & "\preview.sch"

 Set U6A = ActiveDocument.Gates("U6-A")

```

```

 '显示属性"MFG #1"（如果存在）

 If Not U6A.Component.Attributes("MFG #1") Is Nothing Then

 U6A.Visibility(,"MFG #1") = True

 End If

 U6A.Visibility(plogRefDesVisibility) = False '隐藏参考标识符

```

ActiveDocument.ActiveView.Refresh

End Sub

**相关主题**

[View.Refresh](#page-280-0) 方法

### Gate.Delete 方法

此方法从电路图中删除此门电路。

**用法**

Delete

**参数**

无

**描述**

无

**示例**

以下示例从打开的电路图中删除门电路U6-A。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

OpenDocument DefaultFilePath & "\preview.sch"

ActiveDocument.Gates("U6-A").Delete

### Gate.Move 方法

该方法将此门电路移动到新位置。

**用法**

Move(*x* As Double, *y* As Double, [*unit* As [PlogUnit](#page-287-0)])

**参数**

| 参数  | 描述                                                                 |
|-------|----------------------------------------------------------------------|
| x     | 新位置的X坐标                                                       |
| y     | 新位置的Y坐标                                                       |
| unit  | [可选参数 第314页] 指定x和y坐标的单位<br>                           |

**描述**

如果参数无效，该属性会生成一个[异常](#page-290-0)。

**示例**

以下示例代码将所有门电路在当前设计网格上向左移动。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

  For Each nextGate In ActiveDocument.ActiveSheet.Gate

```

```

  nextGate.Move nextGate.PositionX – ActiveDocument.GridX, 

  nextGate.PositionY

```

Next nextGate

End Sub

**相关主题**

[Gate.PositionX 属性](#page-140-0)

[Gate.PositionY 属性](#page-141-0)

## LibraryItem 对象

LibraryItem 对象表示特定零件库中的项目。

### LibraryItem.Application 属性

该属性返回 Application 对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

这是 Microsoft 要求的属性。

### LibraryItem.Library 属性

该属性返回库项目所属的库。

**用法**

Library As [Objecton page 219](#page-194-0)

**参数**

无

**描述**

无

**示例**

以下示例代码显示项目所属库的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set items = GetLibraryItems()

for Each item In items 

MsgBox "项目 " & item & " 属于 " & item.Library & " 库"

Exit For

Next item 

End Sub

```

### LibraryItem.Name 属性

该默认属性返回库项目的名称。

**用法**

Name As String

**参数**

无

**描述**

这是 Microsoft 要求的属性。

**示例**

以下示例显示消息框，展示可用库中第一个项目的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set items = GetLibraryItems()

for Each item In items

MsgBox "第一个项目名称是 " & item.Name

Exit For

Next item

```

### LibraryItem.ObjectType 属性

该属性返回对象类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

无

**示例**

以下示例测试 ObjectType 属性。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set items = GetLibraryItems()

for Each item In items

If item.ObjectType <> plogObjectTypeLibraryItem Then

MsgBox "测试失败"

End If

Next item 

End Sub

```

### LibraryItem.Parent 属性

该属性返回此对象的父对象。

**用法**

Parent As Applicationon page 401

**参数**

无

**描述**

这是 Microsoft 要求的属性。LibraryItem 对象的 Parent 属性始终返回 Application。

### LibraryItem.Type 属性

该属性返回 LibraryItem 对象类型。

**用法**

Type as [PlogLibraryItemType](#page-289-2)

**参数**

无

**描述**

无

**示例**

以下示例显示库项目的类型。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

| Sub Main                         |
|----------------------------------|
| Set items = GetLibraryItems()    |
| for Each item In items           |
| Select Case item.Type            |
| Case plogLibraryItemTypePartType |
| MsgBox "项目类型: PartType"      |
| Case plogLibraryItemType         |
| MsgBox "项目类型: 封装"          |
| Case plogLibraryItemType         |
| MsgBox "项目类型: LogicDrawing"  |
| Case plogLibraryItemType         |
| MsgBox "项目类型: Drawing"       |
|                                  |

| End Select |
|------------|
|            |
| Exit For   |
|            |
| Next item  |
|            |
| End Sub    |
|            |

## Library 对象

Library 对象表示库列表中包含的库。

### Library.Application 属性

该属性返回 Application 对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

这是 Microsoft 要求的属性。

### Library.FullName 属性

该属性返回库文件的完整名称，包括路径和名称。

**用法**

FullName As String

**参数**

无

**描述**

无

**示例**

以下示例显示消息框，展示第一个可用库的完整名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main for Each lib In Libraries MsgBox "第一个可用库的完整名称是 " & lib.FullName Exit For Next lib End Sub

### Library.Name 属性

该默认属性返回库的名称。

**用法**

Name As String

**参数**

无

**描述**

这是 Microsoft 要求的属性。

**示例**

End Sub

以下示例显示消息框，展示第一个可用库的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main for Each lib In Libraries MsgBox "第一个可用库的名称是 " & lib.Name Exit For Next lib

### Library.ObjectType 属性

此属性返回该对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

无

**示例**

以下示例测试 ObjectType 属性。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main for Each lib In Libraries If lib.ObjectType <> plogObjectTypeLibrary Then MsgBox "Test failed" End If Next lib

### Library.Parent 属性

此属性返回该对象的父对象。

**用法**

Parent As Applicationon page 401

**参数**

无

**描述**

这是一个Microsoft要求的属性。

### Library.Path 属性

此属性返回库的路径。

**用法**

Path As String

**参数**

无

**描述**

无

**示例**

以下示例显示一个消息框，展示第一个可用库的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main for Each lib In Libraries MsgBox "The path to the first available library is " & lib.Path Exit For Next lib

### Library.GetLibraryItems 方法

此方法返回此库中所有项目的对象集合，或指定的项目。

**用法**

GetLibraryItems (*Type* as [PlogLibraryItemType](#page-289-2), *Name* as String) As Collection

**参数**

| 参数 | 描述                                                                                                                               |
|------|-----------------------------------------------------------------------------------------------------------------------------------|
| Name | 要检索的LibraryItem对象的名称。可以包含通配符、列表和范围。<br>可选参数，如果省略，则匹配任何名称。 |
| Type | [可选]。指定要检索对象类型的参数。默认为plogLibraryItemTypeAll。                              |

**描述**

无

**示例**

此示例显示库中的项目数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

for Each lib In Libraries

count = lib.GetLibraryItems(plogLibraryItemTypeDecal, "DIP*").Count

MsgBox "Library " & lib.Name & " has " & count & " DIP decals"

Exit For

Next lib

End Sub

```

### Library.ImportLibraryItems 方法

此方法从PADS格式的ASCII文件中读取库项目，并返回最近导入项目的集合。

**用法**

ImportLibraryItems (*Filename* as String) as Collection

**参数**

| 参数 | 描述                                          |
|------|------------------------------------------------------|
| Name | 从中导入库项目的文件名。 |

**描述**

无

**示例**

此示例从文件导入库项目，并显示导入的项目数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

for Each lib In Libraries

Set coll = lib.ImportLibraryItems("C:\sample")

MsgBox coll.Count & " items are successfully imported"

Exit For

Next lib

End Sub

```

### Library.ImportLibraryItems2 方法

此方法从PowerPCB或SailWind Layout格式的ASCII文件中读取库项目。

**用法**

ImportLibraryItems2 (*Filename* as String, ImportOption as PcbImportLibMode) as Collection

**参数**

|          | 参数     | 描述                                              |
|----------|--------------|----------------------------------------------------------|
| Filename |              | 从中导入库项目的文件名。     |
|          | ImportOption | 指定覆盖现有项目的首选项。 |

**返回值**

返回刚刚导入的项目集合。

**描述**

如果ImportOption是pcbImportLibModePrompt，则Application对象会触发OverwriteLibraryItemPrompt事件。客户端应用程序必须处理此事件以指定是否应覆盖现有库项目。

## Measure 对象

Measure对象提供对SailWind Logic内部单位解析器的访问。

Measure对象可以从Application对象构造（参见Application.Measure方法）或从对象获取（参见Attribute.Measure属性）。从measure对象中，您可以提取有关实际值和单位的信息。

### Measure.Application 属性

此属性返回SailWind Logic Application对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为SailWind Logic自动化对象。所有自动化服务器应用程序都有一个Application对象，所有自动化对象都有一个Application属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的大型自动化客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### Measure.Name 属性

此属性返回measure所表示量的名称。

**用法**

Name As String

**参数**

无

**描述**

此属性返回measure对象表示的量的名称，例如"10pF"的measure返回"Capacitance"（电容），"5V"返回"Voltage"（电压）。

**示例**

以下示例代码显示"10pF"measure的量名称（电容）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox Measure("10pF").Quantity

```

End Sub

**相关主题**

[Measure.Value](#page-179-0) 属性

### Measure.Number 属性

此属性返回与前缀和单位组合的数字。

**用法**

Number As Double

**参数**

无

**描述**

此属性返回measure的左侧数字部分。您可以使用创建客户端时使用的编程语言中的标准方法在输出前格式化此数字。将此属性与[Measure.Prefix 属性](#page-175-0)和[Measure.Unit 属性](#page-178-0)一起使用。

**示例**

以下示例使用标准VB格式化输出格式化后的measure（保留小数点后三位）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set Cap = Measure (1.12345e-12, "F") 

MsgBox Format(Cap.Number, "#.###") & Cap.Prefix & Cap.Unit 

End Sub

```

**相关主题**

[Measure.Prefix 属性](#page-175-0)

[Measure.Unit 属性](#page-178-0)

### Measure.Parent 属性

此属性返回SailWind Logic Application对象。

**用法**

Parent As Applicationon page 401

**参数**

无

**描述**

无

### Measure.Prefix 属性

此属性返回与数字和单位组合的单位前缀。

**用法**

Prefix([Format As [PlogMeasureFormat](#page-289-1)= plogMeasureFormatStandard]) As String

**参数**

| 参数 | 描述                                                                            |
|------|----------------------------------------------------------------------------------------|
| Format   | [可选参数] 指示前缀表示格式的参数。 |

**描述**

此属性返回当前在measure中使用的前缀。将此属性与[Measure.Number](#page-173-0) [属性](#page-173-0)和[Measure.Unit 属性](#page-178-0)一起使用。

可能的Format值：

- plogMeasureFormatStandard—返回标准前缀表示（例如，p表示皮，k表示千）

- plogMeasureFormatCurrent—返回当前在此Measure值中使用的前缀格式

- plogMeasureFormatShort—返回短前缀（例如，p表示皮，k表示千）

- plogMeasureFormatLong—返回短单位名称（例如，皮，千，兆）

**示例**

参见[Measure.Number 属性](#page-173-0)的示例

**相关主题**

[Measure.Number 属性](#page-173-0)

[Measure.Unit 属性](#page-178-0)

### Measure.Text 属性

该属性设置或返回测量值的精确文本值。

**用法**

Text As String

**参数**

无

**描述**

此属性定义测量的自定义格式。文本值由数字和可选的前缀及单位组成。如果属性表示测量值，则文本值始终与 [Attribute.Value](#page-49-0) 属性匹配。

**示例**

以下示例代码演示了 [Measure.Text](#page-176-0) 属性、[Measure.Value](#page-179-0) 属性和 [Attribute.Value](#page-49-0) 属性之间的区别。

有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set C1 = ActiveDocument.Components("C1")

Set Cap = Measure("500pF")

C1.Attributes.Add "Capacitance", Cap 

MsgBox Cap '显示 0.0000000005 (默认Value属性)

MsgBox Cap.Value '显示 0.0000000005

MsgBox Cap.Text '显示 500pF 

MsgBox C1.Attributes("Capacitance") '显示 500pF 

MsgBox C1.Attributes("Capacitance").Value '显示 500pF (完整形式)

End Sub

```

**相关主题**

[Measure.Text](#page-176-0) 属性

[Attribute.Value](#page-49-0) 属性

[Measure.Value](#page-179-0) 属性

### Measure.Unit 属性

该属性返回测量的物理单位名称。

**用法**

Unit([Format As [PlogMeasureFormat](#page-289-1)= plogMeasureFormatStandard]) As String

**参数**

| 参数    | 描述                                                                             |
|---------|---------------------------------------------------------------------------------|
| Format  | 可选参数，表示单位的格式<br>表示方式。参见第314页                                |

**描述**

此属性返回不带前缀的物理单位名称。

可能的Format值：

- plogMeasureFormatStandard - 返回标准单位表示（如电容的F）

- plogMeasureFormatCurrent - 返回当前在此Measure值中使用的单位格式

- plogMeasureFormatShort - 返回短单位名称（如电容的F）

- plogMeasureFormatLong - 返回长单位名称（如电容的Farad）

**示例**

参见 [Measure.Number 属性](#page-173-0) 的示例

**相关主题**

[Measure.Number 属性](#page-173-0)

[Measure.Prefix 属性](#page-175-0)

### Measure.Value 属性

该属性设置或返回测量的实际值。

**用法**

Value As Double

**参数**

无

**描述**

此属性返回考虑单位前缀的测量浮点数。例如，"10K"返回10000，200pF返回2e-10。

这是默认属性，因此在Basic脚本中可以省略。

如果设置新的Value，则Measure会自动标准化（参见Normalize方法）。

**示例**

以下示例演示如何使用测量值。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

'1 - 比较两个现有测量属性（可省略.Value调用）

Set U1_Val = 

 ActiveDocument.Components("U1").Attributes("Value").Measure.Value

Set U2_Val = 

 ActiveDocument.Components("U2").Attributes("Value").Measure.Value

Set U1_Quantity = 

 ActiveDocument.Components("U1").Attributes("Value").Measure.Name

Set U2_Quantity = 

 ActiveDocument.Components("U2").Attributes("Value").Measure.Name

If U1_Quantity <> U2_Quantity then

MsgBox "无法比较具有不同物理单位的值"

ElseIf U1_Val < U2_Val then

```

MsgBox "U1 值小于 U2 值"

ElseIf U1_Val > U2_Val then

MsgBox "U1 值大于 U2 值"

Else

MsgBox "U1 值等于 U2 值"

End If

'2 - 检查电阻值是否在100K到10M范围内

Set R1_Val = ActiveDocument.Components("R1").Attributes("Value").Measure

If R1_Val >= Measure("100k") And R1_Val <= Measure("10M") And R1_Val.Name="Resistence" Then

MsgBox "电阻在 [100k, 10M] 范围内"

End If

'3 - 计算所有部件的总热耗散

'确保所有部件都有该属性

For Each part In ActiveDocument.Components

If part.Attributes("Thermal.Dissipation") Is Nothing Then

Part.Attributes.Add "Thermal.Dissipation", Measure("10mW")

End If

Next

Dim Total As Measure '显式声明Total为Measure对象!

Set Total = Measure("0mW") '创建Measure来累加总值

For Each part In ActiveDocument.Components Total = Total + part.Attributes("Thermal.Dissipation").Measure

Next

MsgBox "总热耗散=" & Total.Text

End Sub

**相关主题**

[Measure.Text](#page-176-0) 属性

[Measure.Name 属性](#page-172-0)

### Measure.Normalize 方法

该方法标准化测量的文本值并返回新文本。

**用法**

Normalize As String

**参数**

无

**描述**

此方法选择适当的单位前缀，并在缺少单位时追加单位。例如，它将时间测量的文本从"5e-9"转换为"5ns"。

**示例**

以下示例标准化电容测量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set Cap = Measure(2e-10, "F")

MsgBox Cap.Normalize ' 显示 200pF

End Sub

```

**相关主题**

[Measure.Text](#page-176-0) 属性

## Net 对象

Net对象表示存在于打开的原理图中的物理网络。

### Net.Application 属性

该属性返回SailWind Logic应用程序对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为SailWind Logic自动化对象。所有自动化服务器应用程序都有一个Application对象，所有自动化对象都有一个Application属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的自动化客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### Net.Attributes 属性

该属性返回网络的所有属性集合。

**用法**

Attributes As Attributes

Attributes(*name* As String) As [Attributeon page 69](#page-44-0)

**参数**

| 参数  | 描述                          |
|-------|------------------------------|
| name  | 现有组件属性的名称。          |

**描述**

当传递现有属性*name*给此属性时，它将返回该组件 [Attributeon page 69](#page-44-0) 对象。否则，它将返回 [Attributes Collection Object Property](#page-34-0) 中所有网络属性的集合。

**示例**

以下示例代码使用 [Attributes.Count 属性](#page-36-0) 属性检索网络GND的属性数量（假设它存在于打开的原理图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set attrs = ActiveDocument.Nets("GND").Attributes 

MsgBox "网络GND中有 " & attrs.Count & " 个属性。"

```

### Net.Name 属性

该属性返回网络的名称。

**用法**

Name As String

**参数**

无

**描述**

例如，此属性为网络GND返回字符串"GND"。

此属性是Net对象的默认属性。

**示例**

以下示例代码检索打开的原理图中的所有网络，并将该列表放入自定义对话框中。当在列表框中选择网络时，示例在SailWind Logic中选择该网络。

此示例使用SailWind Logic中Sax Basic引擎的UserDialog编辑器。有关更多信息，请参阅Sax Basic编辑器在线帮助。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Dim ListNets$(10000)

Sub Main

index = 0 

for Each nextNet In ActiveDocument.Nets 

ListNets$(index) = nextNet.Name 

index = index + 1 

Next nextNet 

' 这段代码由SailWind Logic Basic对话框编辑器自动生成。 

Begin Dialog UserDialog 180,238,"网络",.CallbackFunc ' %GRID:10,7,1,1 

ListBox 10,7,160,203,ListNets(),.ListBox1

```

OKButton 10,210,160,21

End Dialog

Dim dlg As UserDialog

Dialog dlg

End Sub

' 当对话框中发生某些事情时，系统会自动调用以下函数

' 用于轻松处理用户操作。

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)

Select Case Action%

Case 2 ' 值改变或按钮按下

If DlgItem$="ListBox1" Then

ActiveDocument.SelectObjects(plogObjectTypeAll, , False)

ActiveDocument.SelectObjects(plogObjectTypeNet, ListNets(SuppValue%))

End If

End Select

End Function

### Net.ObjectType 属性

该属性返回对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

该属性返回 [PlogObjectTypeN](#page-286-0)et。

SailWind Logic 自动化服务器中的所有 SailWind Logic 数据库对象都实现了此属性，以弥补 Basic 语言中没有 Visual C++ QueryInterface 函数等效功能的问题。

该属性通常用于：

- 在异构的 [Objects Collection](#page-194-0) [Object](#page-194-0) 中识别 SailWind Logic 数据库对象的类型

- 当实现一个依赖于作为参数传递的 SailWind Logic 数据库对象类型的通用例程时。例如：

Sub DoSomething(dbObject As Object)

Select Case dbObject.ObjectType

Case plogObjectTypeComponent

' 对组件对象执行特定操作

Case plogObjectTypeNet

' 对网络对象执行特定操作

Case plogObjectTypePin

' 对引脚对象执行特定操作

Case plogObjectTypeGate

' 对门对象执行特定操作

Case Else MsgBox "不是 SailWind Logic 数据库对象" End Select End Sub

### Net.Parent 属性

该属性返回对象的父对象。

**用法**

Parent As Documenton page 419

**参数**

无

**描述**

无

### Net.Pins 属性

该属性返回连接到网络的所有引脚的集合。

**用法**

Pins As [Objectson page 219](#page-194-0)

Pins(*name* As String) As [Pinon page 246](#page-221-0)

**参数**

| 参数 | 描述              |
|------|-------------------|
| name | 现有引脚的名称。 |

**描述**

当传递现有引脚名称给此属性时，它将返回该 [Pin Object](#page-221-0)。否则，它将返回连接到网络的所有引脚的集合，以 [Objects Collection Object](#page-194-0) 形式返回。

**示例**

以下示例代码检索连接到网络 GND 的引脚数量，假设它在打开的示意图中存在。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main MsgBox "GND 连接了 " & ActiveDocument.Nets("GND").Pins.Count & " 个引脚。"

### Net.Selected 属性

该属性设置或返回网络是否被选中✅。

**用法**

Selected As Boolean

**参数**

无

**描述**

您也可以使用 [Document.SelectObjects Method](#page-111-0) 或 [Objects.Select Method](#page-206-0) 方法来选择 SailWind Logic 数据库对象。

**示例**

以下示例代码仅选择网络 GND，假设它在打开的示意图中存在。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

ActiveDocument.SelectObjects(,,False)

```

ActiveDocument.Nets("GND").Selected = True

End Sub

**相关主题**

[Document.SelectionChange Event](#page-117-0)

### Net.Width 属性

该属性返回指定网络的宽度。

**用法**

Width([*Unit* As [PlogUnit](#page-287-0)]) As Double

**参数**

| 参数 | 描述                                                |
|------|-----------------------------------------------------|
| Unit | [Optionaon page 314l] 返回宽度时使用的单位。 |

**描述**

无

**示例**

以下示例代码检索网络 GND 的宽度，假设它在打开的示意图中存在。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

MsgBox "GND 网络的宽度是 " & ActiveDocument.Nets("GND").Width

## Objects Collection 对象

Objects 集合对象是打开的示意图中同质或异质的 SailWind Logic 数据库对象的集合，如 Component 对象、Gate 对象、Net 对象和 Pin 对象。

通常使用 [Document.GetObjects Method](#page-99-0) 方法或使用数据库对象特定属性（如 [Document.Components Property](#page-77-0)、[Document.Fields Property](#page-79-0)、[Document.Nets](#page-85-0) [Property,](#page-85-0) [Document.Pins Property,](#page-90-0) 和 [Document.Gates Property](#page-81-0)）检索此对象。

### Objects.Application 属性

该属性返回 SailWind Logic Application 对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的大型自动化客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### Objects.Count 属性

该属性返回集合中的对象数量。

**用法**

Count As Long

**参数**

无

**描述**

无

**示例**

以下示例代码检索打开的示意图中选定项的数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main Set selectedObjects = ActiveDocument.GetObjects(,,True) MsgBox "有 " & selectedObjects.Count & " 个选定的对象。" End Sub

### Objects.Item 属性

该属性根据索引或名称返回对象。

**用法**

Item(*index* As Long) As Object

Item(*name* As String) As Object

**参数**

| 参数  | 描述                                          |
|-------|-----------------------------------------------|
| index | 要检索的对象在集合中的索引。                 |
| name  | 要检索的对象的名称。                         |

**描述**

这是 Objects 集合对象的默认成员。

如果 index 或 name 参数无效，此属性会生成一个 [Exception](#page-290-0)。

**示例**

以下示例展示了两种不同的方法，使用 [Component Object.](#page-52-0) 遍历打开的示意图中包含的所有 SailWind Logic 数据库对象。第二种方法通常更受青睐，因为它更简洁、更快。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

' 方法1：使用 Object.Item 属性

Sub Main

Set comps = ActiveDocument.Components 

for I=1 To comps.Count 

Set thisComp = comps.Item(I) 

' 对此组件 thisComp 执行操作 

Next I 

End Sub

```

' 方法2：不使用 Object.Item 属性（首选方法）

Sub Main

for Each nextComp in ActiveDocument.Components

' 对此组件 nextComp 执行操作

Next nextComp

End Sub

**相关主题**

[Objects.ItemType](#page-199-0) Property

### Objects.ItemType 属性

该属性根据索引 *index* 返回对象的类型。

**用法**

ItemType(*index* As Long) As [PlogObjectType](#page-286-0)

ItemType(*name* As String) As PlogObjectType

**参数**

| 参数  | 描述                                     |
|-------|------------------------------------------|
| index | 要查询的对象在集合中的索引。            |
| name  | 要检索的对象的名称。                    |

**描述**

如果 *index* 或 *name* 参数无效，此属性会生成一个 [Exception](#page-290-0)。

**相关主题**

[Objects.Item Property](#page-197-0)

### Objects.Next 属性

此属性返回索引 *index* 之后类型为 *type* 的下一个对象的索引。

**用法**

Next(*index* As long, *type* As [PlogObjectType](#page-286-0)) As Long

**参数**

| 参数   | 描述                                     |
|--------|------------------------------------------|
| index  | 要查询的对象在集合中的索引。             |
| type   | 要查询的对象类型。                       |

**描述**

如果 *index* 参数无效，此属性会生成一个[异常](#page-290-0)。

如果 Index = 0，此函数返回给定类型的第一个项目的索引。如果未找到项目，返回值为 0。

### Objects.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Documenton page 419

**参数**

无

**描述**

无

### Objects.Add 方法

此方法向集合添加一个对象。

**用法**

Add(*object* As object)

**参数**

| 参数   | 描述                                                                                                                      |
|--------|--------------------------------------------------------------------------------------------------------------------------|
| object | 要添加到集合的对象。必须是 SailWind Logic 数据对象，如第 77 页的 Component、第 208 页的 Net、第 246 页的 Pin 或第 155 页的 Gate。 |

**描述**

如果 *object* 参数不是 SailWind Logic 数据库对象，此属性会生成一个[异常](#page-290-0)。

**示例**

以下示例代码创建了打开原理图中所有 U* 部件的集合。有关运行此示例的更多信息，请参阅第 21 页的"代码示例"。

Sub Main '创建空集合 Set UComps = ActiveDocument.GetObjects(plogObjectTypeUnknown) for Each nextComp In ActiveDocument.Components If Left(nextComp.Name, 1)="U" Then UComps.Add(nextComp) Next nextComp MsgBox "There are " & UComps.Count & " U\* parts" End Sub

**相关主题**

[Objects.Remove 方法](#page-204-0)

### Objects.Merge 方法

此方法合并两个对象集合。

**用法**

Merge(objects As [Objectson page 219\)](#page-194-0)

**参数**

| 参数    | 描述                                      |
|---------|-------------------------------------------|
| objects | 要与当前对象集合合并的对象集合。          |

**描述**

此方法将 *objects* 集合对象中的所有对象添加到当前对象集合中。

### Objects.Remove 方法

此方法从集合中移除一个对象。

**用法**

Remove(*index* As Long)

Remove(*name* As String)

**参数**

| 参数   | 描述                    |
|--------|-------------------------|
| index  | 要移除的对象的索引。    |
| name   | 要移除的对象的名称。    |

**描述**

如果 *index* 或 *name* 参数无效，此属性会生成一个[异常](#page-290-0)。

**相关主题**

[Objects.Add 方法](#page-202-0)

### Objects.Reset 方法

此方法重置对象集合。

**用法**

Reset()

**参数**

无

**描述**

无

**相关主题**

[Objects.Remove 方法](#page-204-0)

### Objects.Select 方法

此方法选择或取消选择集合中的所有对象。

**用法**

Select([*bselect* As Boolean = True])

**参数**

| 参数    | 描述                                                 |
|---------|------------------------------------------------------|
| bSelect | [Optionalon page 314] True 表示选择。False 表示取消选择。 |

**描述**

无

**示例**

以下示例代码快速选择打开原理图中的所有门。有关运行此示例的更多信息，请参阅第 21 页的"代码示例"。

Sub Main

ActiveDocument.Gates.Select

End Sub

**相关主题**

[Document.SelectionChange 事件](#page-117-0)

### Objects.Sort 方法

此方法按对象名称对集合中的对象进行排序。

**用法**

Sort ()

**参数**

无

**描述**

无

**示例**

以下示例创建引脚集合，按名称排序，并在列表框中显示它们。有关运行此示例的更多信息，请参阅第 21 页的"代码示例"。

```

Dim ListPins$(10000)

Sub Main

 Set Objs = ActiveDocument.Pins

 Objs.Sort 

index = 0

for Each nextPin In Objs

ListPins$(index) = nextPin.Name

index = index + 1

Next nextPin

' 这段代码由 SailWind Logic Basic 对话框编辑器自动生成。

 Begin Dialog UserDialog 180,238,"All Pins Sorted " ' %GRID:10,7,1,1

```

ListBox 10,7,160,203,ListPins(),.ListBox1

OKButton 10,210,160,21

End Dialog

Dim dlg As UserDialog

Dialog dlg

## PartType 对象

PartType 对象表示打开的原理图中存在的物理部件的部件类型（封装）。

### PartType.Application 属性

此属性返回 SailWind Logic Application 对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的大型自动化客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### PartType.Components 属性

此属性返回此部件类型的所有组件的对象集合。

**用法**

Components As [Objectson page 219](#page-194-0)

Components (*name* As String) As [Componenton page 77](#page-52-0)

**参数**

| 参数   | 描述                    |
|--------|-------------------------|
| name   | 现有组件的名称。        |

**描述**

当传递现有引脚 *name* 给此属性时，它返回该[组件对象](#page-52-0)。否则，它返回部件类型的所有组件的集合作为[对象集合对象](#page-194-0)。

**示例**

以下示例显示 7400 部件的总数。

Sub Main

MsgBox Str(ActiveDocument.PartTypes("7400").Components.Count)

### PartType.ECORegistered 属性

此属性设置或返回 PartType 的 ECO 注册状态。

**用法**

Selected As Boolean

**参数**

无

**描述**

无

**示例**

以下示例代码将部件类型"87C256"的 ECO 注册状态设置为未注册（如果存在），然后检查它是否已注册，并显示一个消息框指示它是否已 ECO 注册。

Sub Main

ActiveDocument.PartTypes("87C256").ECORegistered = False

If ActiveDocument.PartTypes("87C256").ECORegistered Then

MsgBox "部件 87C256 已 ECO 注册。"

Else

MsgBox "部件 87C256 未 ECO 注册。"

End If

### PartType.Logic 属性

此属性返回元件类型的逻辑系列。

**用法**

Logic As String

**参数**

无

**描述**

无

**示例**

以下示例代码获取元件类型 7400 的逻辑系列，假设它存在于打开的电路图中。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

Sub Main

```

MsgBox "7400 的逻辑系列是 " & 

 ActiveDocument.PartTypes("7400").Logic

```

### PartType.Name 属性

此属性返回元件类型的名称。

**用法**

Name As String

**参数**

无

**描述**

例如，对于元件类型 7402，此属性返回字符串 "7402"。

此属性是 Component 对象的默认属性。

**示例**

以下示例代码获取打开电路图中的所有元件类型，并将该列表放入自定义对话框列表框中。当在列表框中选择一个元件类型时，示例将在 SailWind Logic 中选择该元件类型的所有元件。此示例使用 SailWind Logic 中 Sax Basic 引擎的 UserDialog 编辑器。更多信息请参见 Sax Basic 编辑器在线帮助。

有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

```

Dim ListPkgs$(10000)

Sub Main

index = 0 

for Each nextPkg In ActiveDocument.PartTypes

ListPkgs$(index) = nextPkg.Name

index = index + 1

Next nextPkg

' 这段代码由 SailWind Logic Basic 对话框编辑器自动生成

Begin Dialog UserDialog 180,238,"Part Types",.CallbackFunc ' 

 %GRID:10,7,1,1

```

ListBox 10,7,160,203,ListPkgs(),.ListBox1

OKButton 10,210,160,21

End Dialog

Dim dlg As UserDialog

Dialog dlg

End Sub

```

' 当对话框中发生某些事件时，系统会自动调用以下函数

```

' 用于轻松处理用户操作

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)

Select Case Action%

Case 2 ' 值改变或按钮按下

If DlgItem$="ListBox1" Then

ActiveDocument.SelectObjects(plogObjectTypeAll, , False)

'按名称获取元件

Set pkg = ActiveDocument.PartTypes(ListPkgs(SuppValue%))

'选择元件

pkg.Selected = True

'激活元件第一个门所在的图纸

pkg.Components(1).Gates(1).Sheet.Activate

End If

End Select

End Function

### PartType.ObjectType 属性

此属性返回对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

此属性返回 [PlogObjectTypeP](#page-286-0)artType。

SailWind Logic 自动化服务器中的所有 SailWind Logic 数据库对象都实现了此属性，以弥补 Basic 中没有 Visual C++ QueryInterface 函数等效功能的问题。

此属性通常用于：

- 在异构的 [Objects Collection](#page-194-0) [Object](#page-194-0) 中识别 SailWind Logic 数据库对象的种类

- 当实现一个依赖于作为参数传递的 SailWind Logic 数据库对象类型的通用例程时。例如：

Sub DoSomething(dbObject As Object)

Select Case dbObject.ObjectType

Case plogObjectTypeComponent

' 对元件对象执行特定操作

Case plogObjectTypeNet

' 对网络对象执行特定操作

Case plogObjectTypePin

' 对引脚对象执行特定操作

Case plogObjectTypeGate

' 对门对象执行特定操作

Case plogObjectTypePartType

' 对元件类型对象执行特定操作

Case Else

MsgBox "不是 SailWind Logic 数据库对象"

End Select

### PartType.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Documenton page 419

**参数**

无

**描述**

无

### PartType.Selected 属性

此属性设置或返回元件类型的元件是否被选中✅。

**用法**

Selected As Boolean

**参数**

无

**描述**

当此元件类型的一个或多个元件被选中✅时，该元件类型被视为已选中✅。您也可以使用 [Document.SelectObjects Method](#page-111-0) 或 [Objects.Select Method](#page-206-0) 方法选择 SailWind Logic 数据库对象。

**示例**

以下示例代码仅选择元件类型 7400，假设它存在于打开的电路图中，并激活其所在的图纸。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

```

Sub Main

ActiveDocument.SelectObjects(,,False)

 ActiveDocument.PartTypes("7400").Selected = True

```

```

 ActiveDocument.PartTypes("7400").Components(1).Gates(1).Sheet.Activate

```

End Sub

**相关主题**

[Document.SelectionChange Event](#page-117-0)

## Pin 对象

Pin 对象表示存在于打开电路图中的物理引脚。

### Pin.AlphaNumber 属性

此属性返回引脚的字母数字编号。

**用法**

Alpha Number As String

**参数**

无

**描述**

无

**示例**

以下示例代码获取引脚 U1.1 的字母数字引脚编号，假设它存在于打开的电路图中。有关运行此示例的更多信息，请参见第 21 页的"代码示例"。

```

Sub Main

 Set U1_1 = ActiveDocument.Pins("U1.1")

MsgBox "引脚 U1.1 在元件 " &

```

U1\_1.Component.Name 中的引脚编号为 " & U1_1.AlphaNumber

### Pin.Application 属性

此属性返回 SailWind Logic Application 对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的大型自动化客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### Pin.Component 属性

此属性返回引脚所属的元件。

**用法**

Component As [Component 第77页](#page-52-0)

**参数**

无

**描述**

无

**示例**

以下示例代码显示引脚 U1.1 所属的元件（假设该引脚存在于打开的电路图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox "引脚 U1.1 属于元件 " & 

 ActiveDocument.Pins("U1.1").Component.Name

```

### Pin.ElectricalType 属性

此属性返回引脚的电气类型。

**用法**

ElectricalType As [PlogPinElectricalType](#page-287-2)

**参数**

无

**描述**

无

**示例**

以下示例代码获取引脚 U1.1 的电气类型（假设该引脚存在于打开的电路图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox "引脚 U1.1 的类型是 " & 

 ElectricalTypeName(ActiveDocument.Pins("U1.1").ElectricalType)

```

End Sub

Function ElectricalTypeName(theType As Long) As String

Select Case theType

Case plogElectricalTypeUnknown

ElectricalTypeName="unknown"

Case plogElectricalTypeSource

ElectricalTypeName="source"

Case plogElectricalTypeBidirectional

ElectricalTypeName="bidirectional"

Case plogElectricalTypeOpenCollector

ElectricalTypeName="open collector"

Case plogElectricalTypeOrTieableSource

ElectricalTypeName="tiable source"

Case plogElectricalTypeTristate

ElectricalTypeName="tristate"

Case plogElectricalTypeLoad

ElectricalTypeName="load"

Case plogElectricalTypeTerminator

ElectricalTypeName="terminator"

Case plogElectricalTypePower

ElectricalTypeName="power"

Case plogElectricalTypeGround

ElectricalTypeName="ground"

Case Else

ElectricalTypeName="unknown"

End Select

End Function

### Pin.FunctionName 属性

此属性返回引脚的功能名称。

**用法**

FunctionName As String

**参数**

无

**描述**

无

**示例**

以下示例代码获取引脚 U1.1 的功能名称（假设该引脚存在于打开的电路图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox "引脚 U1.1 的功能名称是 " & 

 ActiveDocument.Pins("U1.1").FunctionName

```

### Pin.Gate 属性

此属性返回引脚所属的门电路。

**用法**

Gate As [Gate 第155页](#page-130-0)

**参数**

无

**描述**

如果引脚不属于任何门电路，此属性返回 Nothing。

**示例**

以下示例代码获取引脚 U1.1 所属的门电路（假设该引脚存在于打开的电路图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 Set U1_1 = ActiveDocument.Pins("U1.1")

 If U1_1.Gate Is Nothing Then

 MsgBox "引脚 U1.1 不属于任何门电路" 

 Else

 MsgBox "引脚 U1.1 属于门电路 " & U1_1.Gate.Name

 End If

```

### Pin.GatePinName 属性

此属性返回通过其门电路的引脚名称。

**用法**

GatePinName As String

**参数**

无

**描述**

此属性返回由门电路名称和本地引脚名称组成的名称。例如：U1-B.1、U2-B.A。

如果引脚不属于任何门电路，返回值与 [Pin.Name 属性](#page-230-0) 相同。

**示例**

以下示例代码获取引脚 U1.1 的门电路引脚名称（假设该引脚存在于打开的电路图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

MsgBox "引脚 U1.1 的门电路引脚名称是 " & ActiveDocument.Pins("U1.1").GatePinName

### Pin.Name 属性

此属性返回引脚的名称。

**用法**

Name As String

**参数**

无

**描述**

例如，对于引脚 U1.1，此属性返回字符串 "U1.1"。

此属性是 Pin 对象的默认属性。

**示例**

以下示例代码在自定义对话框中按名称列出所有引脚。当在列表框中选择一个引脚时，该示例也会在 SailWind Logic 中选择该引脚。

此示例使用 SailWind Logic 中 Sax Basic 引擎的 UserDialog 编辑器。更多信息请参阅 Sax Basic 编辑器在线帮助。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Dim ListPins$(10000)

Sub Main

index = 0 

for Each nextPin In ActiveDocument.Pins 

ListPins$(index) = nextPin.Name 

index = index + 1 

Next nextPin 

' 这段代码由 SailWind Logic Basic 对话框编辑器自动生成

Begin Dialog UserDialog 180,238,"Pins",.CallbackFunc ' %GRID:10,7,1,1 

ListBox 10,7,160,203,ListPins(),.ListBox1

```

OKButton 10,210,160,21

End Dialog

Dim dlg As UserDialog

Dialog dlg

End Sub

' 当对话框中发生某些事件时，系统会自动调用以下函数

' 用于轻松处理用户操作

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)

Select Case Action%

Case 2 ' 值改变或按钮按下

If DlgItem$="ListBox1" Then

ActiveDocument.SelectObjects(plogObjectTypeAll, , False)

ActiveDocument.SelectObjects(plogObjectTypePin, ListPins(SuppValue%))

End If

End Select

End Function

### Pin.Net 属性

此属性返回与引脚连接的网络。

**用法**

Net As [Net 第208页](#page-183-0)

**参数**

无

**描述**

如果引脚未连接，此属性返回 Nothing。

**示例**

以下示例代码获取引脚 U1.1 连接的网络，假设它在打开的电路图中存在。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

```

MsgBox "引脚 U1.1 连接到网络 " & 

 ActiveDocument.Pins("U1.1").Net.Name

```

### Pin.ObjectType 属性

此属性返回对象的类型。

**用法**

ObjectType As [PlogObjectType](#page-286-0)

**参数**

无

**描述**

此属性返回 [PlogObjectTypeP](#page-286-0)in。

SailWind Logic 自动化服务器中的所有 SailWind Logic 数据库对象都实现了此属性，以弥补 Basic 中没有 Visual C++ QueryInterface 函数等效项的问题。

此属性通常用于：

- 在异构的 [Objects 集合](#page-194-0) [对象](#page-194-0)中识别 SailWind Logic 数据库对象的类型

- 当实现依赖于作为参数传递的 SailWind Logic 数据库对象类型的通用例程时。例如：

Sub DoSomething(dbObject As Object)

Select Case dbObject.ObjectType

Case plogObjectTypeComponent

' 对组件对象执行特定操作

Case plogObjectTypeNet

' 对网络对象执行特定操作

Case plogObjectTypePin

' 对引脚对象执行特定操作

Case plogObjectTypeGate

' 对门对象执行特定操作

Case Else

MsgBox "不是 SailWind Logic 数据库对象"

End Select

### Pin.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Document 第419页

**参数**

无

**描述**

无

### Pin.PositionX 属性

此属性返回引脚的 x 坐标。

**用法**

PositionX([*unit* As [PlogUnit](#page-287-0)]) As Double

**参数**

| 参数 | 描述 |
|------|------|
| unit | [可选 第314页] 返回 x 坐标的单位 |

**描述**

如果引脚未使用或是信号引脚，此属性会生成 [异常](#page-290-0)。

**示例**

以下示例代码获取引脚 U1.1 的坐标，假设它在打开的电路图中存在，使用当前设计单位。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main Set pinU1\_1 = ActiveDocument.Pins("U1.1") MsgBox "U1.1 位置是 (" & pinU1\_1.PositionX & ", " & pinU1\_1.PositionY & ")"

End Sub

**相关主题**

[Pin.PositionY 属性](#page-237-0)

### Pin.PositionY 属性

此属性返回引脚的 y 坐标。

**用法**

PositionY([*unit* As [PlogUnit](#page-287-0)]) As Double

**参数**

| 参数 | 描述 |
|------|------|
| unit | [可选 第314页] 返回 y 坐标的单位 |

**描述**

如果引脚未使用或是信号引脚，此属性会生成 [异常](#page-290-0)。

**示例**

以下示例代码获取引脚 U1.1 的坐标，假设它在打开的电路图中存在，使用当前设计单位。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main Set pinU1\_1 = ActiveDocument.Pins("U1.1") MsgBox "U1.1 位置是 (" & pinU1\_1.PositionX & ", " & pinU1\_1.PositionY & ")"

End Sub

**相关主题**

[Pin.PositionX 属性](#page-236-0)

### Pin.Selected 属性

此属性设置或返回引脚是否被选中✅。

**用法**

Selected As Boolean

**参数**

无

**描述**

您也可以使用 [Document.SelectObjects 方法](#page-111-0) 或 [Objects.Select 方法](#page-206-0) 来选择 SailWind Logic 数据库对象。

**示例**

以下示例代码仅选择引脚 U1.1，假设它在打开的电路图中存在。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

ActiveDocument.SelectObjects(,,False)

```

ActiveDocument.Pins("U1.1").Selected = True

End Sub

**相关主题**

[Document.SelectionChange 事件](#page-117-0)

### Pin.SwapClass 属性

此属性返回此引脚的交换类。

**用法**

SwapClass As Long

**参数**

无

**描述**

无

**示例**

以下示例代码获取引脚 U1.1 的交换类，假设它在打开的电路图中存在。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

 Set U1_1 = ActiveDocument.Pins("U1.1")

MsgBox "引脚 U1.1 的交换类是 " & U1_1.SwapClass

```

End Sub

**相关主题**

[Gate.SwapClass 属性](#page-147-0)

## Sheet 对象

Sheet 对象表示单个 SailWind Logic 原理图图纸。

此对象通常从 [Sheets 集合对象](#page-261-0) 中获取。

### Sheet.Application 属性

此属性返回 SailWind Logic 应用程序对象。

**用法**

Application As Application 第401页

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的大型自动化客户端应用程序中，以快速识别对象所属的应用程序。

### Sheet.ChildSheets 属性

此属性返回此图纸中的子图纸集合。

**用法**

ChildSheets As [Sheets 第265页](#page-240-0)

ChildSheets(*name* As String) As [Sheet 第265页](#page-240-0)

**参数**

| 参数 | 描述 |
|------|------|
| name | 现有图纸的名称 |

**描述**

当传递现有图纸名称时，此属性返回该图纸包装为 [Sheet](#page-240-0) [对象](#page-240-0)。否则，它返回所有子图纸的集合，包装为 [Sheets 集合对象](#page-261-0)。

**示例**

以下示例代码使用 [Sheets.Count 属性](#page-263-0) 属性获取活动图纸的子图纸数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set shts = ActiveDocument.ActiveSheet.ChildSheets 

MsgBox "在 " & 

 ActiveDocument.ActiveSheet.Name & " 中有 " & shts.Count & " 个子图纸"
```
### Sheet.Components 属性

此属性返回当前图纸中所有组件的集合。

**用法**

Components As [Objectson page 219](#page-194-0)

Components(*name* As String) As [Componenton page 77](#page-52-0)

**参数**

| 参数 | 描述                    |
|----------|--------------------------------|
| name     | 现有组件的名称 |

**描述**

当传入现有组件名称 *name* 时，此属性返回该[组件对象](#page-52-0)。否则返回[对象集合](#page-194-0)中所有组件的集合。

**示例**

以下示例代码使用[Objects.Count](#page-196-0) [属性](#page-196-0)获取活动图纸中的组件数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set comps = ActiveDocument.ActiveSheet.Components 

MsgBox "当前图纸 " & ActiveDocument.ActiveSheet.Name & " 中共有 " & comps.Count & " 个组件"

```

以下示例代码使用[Component.Pins 属性](#page-65-0)和[Objects.Count 属性](#page-196-0)获取组件U1的引脚数量（假设U1存在于活动图纸中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set compU1 = ActiveDocument.ActiveSheet.Components("U1")

MsgBox "组件 " & compU1.Name & " 共有 " & compU1.Pins.Count & " 个引脚"

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Sheet.GetObjects 方法](#page-258-0)

[Document.Components 属性](#page-77-0)

### Sheet.Gates 属性

此属性返回当前图纸中所有门电路的集合。

**用法**

Gates As [Objectson page 219](#page-194-0)

Gates(*name* As String) As [Gateon page 155](#page-130-0)

**参数**

| 参数 | 描述               |
|----------|---------------------------|
| name     | 现有门电路的名称 |

**描述**

当传入现有门电路名称时，此属性返回该[门电路对象](#page-130-0)。否则返回[对象集合](#page-194-0)中所有门电路的集合。

**示例**

以下示例代码使用[Objects.Count](#page-196-0) [属性](#page-196-0)获取活动图纸中的门电路数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set gates = ActiveDocument.ActiveSheet.Gates

MsgBox "当前图纸 " & ActiveDocument.ActiveSheet.Name & " 中共有 " & gates.Count & " 个门电路"

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

[Sheet.GetObjects 方法](#page-258-0)

[Document.Gates 属性](#page-81-0)

### Sheet.Name 属性

此属性获取或设置图纸名称。

**用法**

Name As String

**参数**

无

**描述**

此属性是Sheet对象的默认属性。

**示例**

以下示例代码获取活动图纸的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main MsgBox "您好，您正在使用 " & ActiveDocument.ActiveSheet.Name

### Sheet.Nets 属性

此属性返回当前图纸中所有网络的集合。

**用法**

Nets As [Objectson page 219](#page-194-0)

Nets(*name* As String) As [Neton page 208](#page-183-0)

**参数**

| 参数 | 描述              |
|----------|--------------------------|
| name     | 现有网络的名称 |

**描述**

当传入现有网络名称 *name* 时，此属性返回该[网络对象](#page-183-0)。否则返回[对象集合](#page-194-0)中所有网络的集合。

**示例**

以下示例代码使用[Objects.Count](#page-196-0) [属性](#page-196-0)获取活动图纸中的网络数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set nets = ActiveDocument.ActiveSheet.Nets 

MsgBox "当前图纸 " & ActiveDocument.ActiveSheet.Name & " 中共有 " & nets.Count & " 个网络"

```

以下示例代码使用[Net.Pins 属性](#page-191-0)和[Objects.Count 属性](#page-196-0)获取网络VCC连接的引脚数量（假设VCC存在于活动图纸中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set netVCC = ActiveDocument.ActiveSheet.Nets("VCC") 

MsgBox "网络 " & netVCC.Name & " 共连接了 " & netVCC.Pins.Count & " 个引脚"

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

### Sheet.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Applicationon page 401

**参数**

无

**描述**

无

### Sheet.ParentSheet 属性

此属性返回当前图纸的父图纸。

**用法**

Parent As Applicationon page 401

**参数**

无

**描述**

如果当前图纸是祖先图纸，则返回值为Nothing。

**示例**

以下示例获取活动图纸的父图纸名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

If Not ActiveDocument.ActiveSheet.ParentSheet Is Nothing Then

MsgBox ActiveDocument.ActiveSheet.ParentSheet

Else

MsgBox "图纸 " & ActiveDocument.ActiveSheet.Name & " 是祖先图纸"

End If

### Sheet.PartTypes 属性

此属性返回当前图纸中所有零件类型的集合。

**用法**

PartTypes As [Objectson page 219](#page-194-0)

PartTypes(*name* As String) As [PartType on page 234](#page-209-0)

**参数**

| 参数 | 描述                    |
|----------|--------------------------------|
| name     | 现有零件类型的名称 |

**描述**

当传入现有零件类型名称 *name* 时，此属性返回该[零件类型](#page-209-0)对象。否则返回[对象集合](#page-194-0)中所有零件类型的集合。

**示例**

以下示例代码使用[Objects.Count 属性](#page-196-0)获取打开的原理图中当前图纸的零件类型数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set pkgs = ActiveDocument.ActiveSheet.PartTypes 

MsgBox "当前原理图 " & ActiveDocument.Name & " 的活动图纸中共有 " & pkgs.Count & " 种零件类型"

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

### Sheet.Pins 属性

此属性返回当前图纸中所有引脚的集合。

**用法**

Pins As [Objectson page 219](#page-194-0)

Pins(*name* As String) As [Pinon page 246](#page-221-0)

**参数**

| 参数 | 描述              |
|----------|--------------------------|
| name     | 现有引脚的名称 |

**描述**

当传入现有引脚名称 *name* 时，此属性返回该[引脚对象](#page-221-0)。否则返回[对象集合](#page-194-0)中所有引脚的集合。

**示例**

以下示例代码使用[Objects.Count](#page-196-0) [属性](#page-196-0)获取活动图纸中的引脚数量。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

Set pins = ActiveDocument.ActiveSheet.Pins 

MsgBox "当前图纸 " & ActiveDocument.ActiveSheet.Name & " 中共有 " & Pins.Count & " 个引脚"

```

**相关主题**

[Document.GetObjects 方法](#page-99-0)

### Sheet.Activate 方法

此方法用于激活指定图纸。

**用法**

Activate

**参数**

无

**描述**

无

**示例**

以下示例激活活动文档中的第一张图纸。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main

```

 ActiveDocument.Sheets(1).Activate

```

### Sheet.AddComponent 方法

此方法向活动图纸添加新部件（包括其所有门电路），并返回新部件对象。

**用法**

AddComponent (*PartType* As String, [*RefDes* As String], [*PositionX* As Double = [PlogDefaultPosition](#page-289-0)X], [*PositionY* As Double = plogDefaultPositionX], [*DeltaX* As Double = plogDefaultPositionX], [*DeltaY* As Double = plogDefaultPositionX], [*Unit* As [PlogUnit\]](#page-287-0)) As [Componenton page 77](#page-52-0)

**参数**

| 参数       | 描述                                                                                                                              |
|------------|------------------------------------------------------------------------------------------------------------------------------------------|
| PartType:  | 必需的部件类型名称。                                                                                                                 |
| RefDes:    | [可选 第314页] 此门电路所属部件的名称。如果未指定RefDes，<br>则使用下一个参考标识符。 |
| PositionX: | [可选] 新门电路的x坐标。如果未指定PositionX，则会在当前图纸视图中<br>选择一个x坐标。    |
| PositionY: | [可选] 新门电路的y坐标。如果未指定PositionY，则会在当前图纸视图中<br>选择一个y坐标。     |
| DeltaX:    | [可选] 门电路之间的x坐标增量。如果未指定DeltaX，则自动<br>计算。                          |
| DeltaY:    | [可选] 门电路之间的y坐标增量。如果未指定DeltaY，则自动<br>计算。                          |
| Unit:      | [可选] 指定位置的单位。如果未指定Unit，则使用当前单位。<br>                                   |

**返回值**

新部件对象，封装为[Component 第77页](#page-52-0)。

**描述**

如果参数无效，此属性会生成[异常](#page-290-0)。

**示例**

以下示例向活动图纸添加一个新部件并选中✅它。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

 Set comp = ActiveDocument.ActiveSheet.AddComponent("7402", "U2")

```

Comp.Selected = True

End Sub

**相关主题**

[Sheet.AddGate 方法](#page-256-0)

[Component.Delete 方法](#page-70-0)

### Sheet.AddGate 方法

此方法向活动图纸添加新门电路，并返回新门电路对象。

**用法**

AddGate(*PartType* As String, [*RefDes* As String], [*GateIndex* As Long], [*PositionX* As Double = [PlogDefaultPosition](#page-289-0)X], [*PositionY* As Double = plogDefaultPositionY], [*Unit* As PlogUnit]) As [Gate](#page-130-0)  [第155页](#page-130-0)

**参数**

| 参数      | 描述                                                                                                                                       |
|-----------|---------------------------------------------------------------------------------------------------------------------------------------------------|
| PartType  | 必需的部件类型名称。                                                                                                                          |
| RefDes    | [可选 第314页] 此门电路所属部件的名称。如果未指定RefDes，<br>则使用下一个可用的参考标识符 |
| GateIndex | [可选] 部件中的门电路索引。如果未指定GateIndex，则使用下一个可用的门电路<br>索引。                                       |
| PositionX | [可选] 新门电路的x坐标。如果未指定PositionX，则会在当前图纸视图中<br>选择一个x坐标。             |
| PositionY | [可选] 新门电路的y坐标。如果未指定PositionY，则会在当前图纸视图中<br>选择一个y坐标。              |
| Unit      | [可选] 指定位置的单位。如果未指定Unit，则使用当前单位。<br>                                            |

**返回值**

新门电路对象，封装为[Gate 第155页](#page-130-0)。

**描述**

如果参数无效，此属性会生成[异常](#page-290-0)。

**示例**

以下示例将部件的第一个未使用门电路添加到活动图纸并选中✅它。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

 Set part = ActiveDocument.Components("U1")

 Set sht = ActiveDocument.ActiveSheet

```

```

 For Each gt in part.UnusedGates

 Set newGate = sht.AddGate(part.PartType, part.Name, gt.Number)

 newGate.Selected = True

 Exit For

 Next

End Sub

```

**相关主题**

[Gate.Delete 方法](#page-150-0)

[Sheet.AddComponent 方法](#page-254-0)

### Sheet.GetObjects 方法

此方法返回此图纸中SailWind Logic数据库对象的集合。

**用法**

GetObjects([*type* As [PlogObjectType](#page-286-0) = plogObjectTypeAll], [*name* As String], [*selected* As Boolean = False]) As [Objects 第219页](#page-194-0)

**参数**

| 参数     | 描述                                                             |
|----------|-------------------------------------------------------------------------|
| type     | [可选 第314页] 要获取的SailWind Logic数据库对象类型。    |
| name     | [可选] 要获取对象的值或名称。                       |
| selected | [可选] True表示仅获取选中✅的对象。False表示获取所有对象。 |

**返回值**

返回的对象是一个[Objects 集合对象](#page-194-0)。如果没有对象满足请求，则返回空集合。

**描述**

此方法的所有参数都是可选的，这意味着可以不带任何参数调用它，也可以使用任意参数组合调用它。有关更多信息，请参见下面的示例。

*Name* 参数支持通配符("U*")、逗号分隔的项目列表("U1, U2, R1")、由两个对象名称和短划线字符指定的范围("U1 - U10, U12, R1 - R20")。短划线必须用空格包围，因为短划线是对象名称中的合法符号。每个名称只允许一个通配符，不能在范围中使用通配符。可以传递*name*如"U*, R*, C1 – C100"，但不能传递"U*1*"或"C1* - C10*"。

要获取相同类型的所有对象，请使用对象Sheet属性。例如，要获取图纸中的所有门电路，请使用[Sheet.Gates 属性](#page-245-0)而不是[Sheet.GetObjects 方法](#page-258-0)([PlogObjectType](#page-286-0))。

如果type参数不是有效的SailWind Logic数据库对象类型，此属性会生成[异常](#page-290-0)。

**示例**

以下示例代码展示了使用此方法的不同方式，使用[Objects.Count 属性](#page-196-0)属性检索每种方式的对象数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

Sub Main

Dim objs As Object

```

```

' 示例1: 获取所有类型的对象

Set objs = ActiveDocument.ActiveSheet.GetObjects 

MsgBox "示例1: " & objs.Count & " 个对象。" 

' 示例2: 获取所有选中✅的对象

Set objs = ActiveDocument.ActiveSheet.GetObjects(,,True) 

MsgBox "示例2: " & objs.Count & " 个选中✅的对象。" 

' 示例3: 获取所有网络对象

Set objs = ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeNet) 

MsgBox "示例3: " & objs.Count & " 个网络对象。" 

' 示例4: 获取名称为"VCC"的所有网络对象(至少有1个)

Set objs = ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeNet, "VCC") 

MsgBox "示例4: " & objs.Count & " 个VCC网络对象。" 

' 示例5: 获取名称以U开头的所有部件对象

Set objs = ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeComponent, 

 "U*") 

MsgBox "示例3: " & objs.Count & " 个U*部件对象。"


End Sub
```

**相关主题**

[Document.SelectObjects 方法](#page-111-0)

[Document.Components 属性](#page-77-0)

[Document.Gates 属性](#page-81-0)

[Document.Nets 属性](#page-85-0)

[Document.Pins 属性](#page-90-0)

[Document.PartTypes](#page-88-0) 属性

[Document.GetObjects 方法](#page-99-0)

[Sheet.Components 属性](#page-243-0)

[Sheet.Gates 属性](#page-245-0)

[Sheet.Nets 属性](#page-247-0)

[Sheet.Pins 属性](#page-252-0)

[Sheet.PartTypes](#page-251-0) 属性

## Sheets 集合对象

Sheets集合对象是打开的原理图中所有图纸的集合。

通常使用[Document.Sheets 属性](#page-92-0)属性获取此对象。

以下列表标识了Sheets属性和方法：

[Sheets.Application 属性](#page-262-0) [Sheets.Count 属性](#page-263-0) [Sheets.Item 属性](#page-264-0) [Sheets.Parent 属性](#page-266-0) [Sheets.Add 方法](#page-267-0) [Sheets.Delete 方法](#page-268-0)

### Sheets.Application 属性

此属性返回SailWind Logic Application对象。

**用法**

Application As Applicationon page 401

**参数**

无

**描述**

此属性将对象标识为SailWind Logic Automation对象。所有Automation服务器应用程序都有一个Application对象，所有Automation对象都有一个Application属性。此属性通常用于处理来自不同源(如不同的Automation服务器应用程序)的大量对象的大型Automation客户端应用程序中。使用该属性可以快速识别对象所属的应用程序。

### Sheets.Count 属性

此属性返回图纸数量。

**用法**

Count As Long

**参数**

无

**描述**

无

**示例**

以下示例代码检索打开的原理图中的图纸数量。有关运行此示例的更多信息，请参见第21页的"代码示例"。

Sub Main Set shts = ActiveDocument.Sheets MsgBox "在 " & ActiveDocument.Name & " 中有 " & shts.Count & " 张图纸" End Sub

### Sheets.Item 属性

此属性根据索引或名称返回图纸。

**用法**

Item(*index* As Long) As [Sheet 第265页](#page-240-0)

Item(*name* As String) As [Sheet 第265页](#page-240-0)

**参数**

| 参数    | 描述                                         |
|---------|-----------------------------------------------------|
| index   | 要检索的图纸在集合中的索引。 |
| name    | 要检索的图纸名称。                      |

**描述**

这是Sheets集合对象的默认成员。

如果index或name参数无效，此属性会生成[异常](#page-290-0)。

**示例**

以下示例代码展示了两种遍历打开的原理图中所有图纸的不同方法。通常首选第二种方法，因为它更简洁、更快。有关运行此示例的更多信息，请参见第21页的"代码示例"。

```

' 方法1: 使用Object.Item属性

Sub Main

Set shts = ActiveDocument.Sheets

for I=1 To shts.Count

Set thisSht = shts.Item(I)

' 对图纸执行某些操作

Next I

End Sub

```

' 方法2: 不使用Object.Item属性(首选方法)

Sub Main

for Each nextSht in ActiveDocument.Sheets

' 对图纸执行某些操作

Next nextSht

### Sheets.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Document 第419页

**参数**

无

**描述**

无

### Sheets.Add 方法

此方法向原理图添加新图纸。

**用法**

Add As [Sheet 第265页](#page-240-0)

**参数**

无

**返回值**

新图纸打包为 [Sheet 对象](#page-240-0)

**相关主题**

[Sheets.Delete 方法](#page-268-0)

### Sheets.Delete 方法

此方法删除图纸。

**用法**

Delete(*index* As Long)

Delete(*name* As String)

**参数**

| 参数  | 描述                                       |
|-------|-------------------------------------------|
| index | 要删除的图纸在集合中的索引                 |
| name  | 要删除的图纸名称                           |

**描述**

此属性在以下情况下会生成 [异常](#page-290-0)：

- 如果 *name* 参数不是现有图纸

- 如果 *name* 参数不是有效的可删除图纸

- 如果 *index* 参数大于现有图纸数量

**示例**

以下示例代码从打开的原理图中删除名为 "MYSHEET" 的图纸。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

ActiveDocument.Sheets.Delete("MYSHEET")

End Sub

**相关主题**

[Sheets.Add 方法](#page-267-0)

## View 对象

View 对象表示当前 SailWind Logic 视图窗口，显示打开的原理图中的活动图纸。

此对象通常通过 [Document.ActiveView](#page-74-0) 属性获取。

以下是 View 属性、方法和事件的列表：

[View.Application](#page-270-0) 属性 [View.BottomRightX](#page-271-0) 属性 [View.BottomRightY](#page-272-0) 属性 [View.Name](#page-273-0) 属性 [View.Parent](#page-274-0) 属性 [View.PointerX](#page-275-0) 属性 [View.PointerY](#page-276-0) 属性 [View.TopLeftX](#page-277-0) 属性 [View.TopLeftY](#page-278-0) 属性 [View.Pan](#page-279-0) 方法 [View.Refresh](#page-280-0) 方法 [View.SetExtents](#page-281-0) 方法 [View.SetExtentsToAll](#page-283-0) 方法 [View.SetExtentsToSheet](#page-284-0) 方法 [View.Change](#page-285-0) 事件

### View.Application 属性

此属性返回 SailWind Logic 应用程序对象。

**用法**

Application As Application 第401页

**参数**

无

**描述**

此属性将对象标识为 SailWind Logic 自动化对象。所有自动化服务器应用程序都有一个 Application 对象，所有自动化对象都有一个 Application 属性。此属性通常用于处理来自不同源（如不同的自动化服务器应用程序）的大量对象的自动化客户端应用程序中。使用该属性可快速识别对象所属的应用程序。

### View.BottomRightX 属性

此属性返回视图右下角的 X 值。

**用法**

BottomRightX([*unit* As [PlogUnit\]](#page-287-0)) As Double

**参数**

| 参数  | 描述                                                              |
|-------|------------------------------------------------------------------|
| unit  | [可选 第314页] 返回右下角 X 值的单位                             |

**描述**

无

**示例**

以下示例代码获取打开的原理图中当前视图的坐标。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

x0 = Format$(ActiveDocument.ActiveView.TopLeftX, "Fixed")

y0 = Format$(ActiveDocument.ActiveView.TopLeftY, "Fixed")

x1 = Format$(ActiveDocument.ActiveView.BottomRightX, "Fixed")

y1 = Format$(ActiveDocument.ActiveView.BottomRightY, "Fixed")

MsgBox "View is (" & x0 & ", " & y0 & ") - (" & x1 & ", " & y1 & ")


End Sub
```

**相关主题**

[View.BottomRightY](#page-272-0) 属性

[View.TopLeftX](#page-277-0) 属性

[View.TopLeftY](#page-278-0) 属性

### View.BottomRightY 属性

此属性返回视图右下角的 Y 值。

**用法**

BottomRightY([*unit* As [PlogUnit\]](#page-287-0)) As Double

**参数**

| 参数  | 描述                                                              |
|-------|------------------------------------------------------------------|
| unit  | [可选 第314页] 返回右下角 Y 值的单位                             |

**描述**

无

**示例**

以下示例代码获取打开的原理图中当前视图的坐标。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

x0 = Format$(ActiveDocument.ActiveView.TopLeftX, "Fixed") 

y0 = Format$(ActiveDocument.ActiveView.TopLeftY, "Fixed") 

x1 = Format$(ActiveDocument.ActiveView.BottomRightX, "Fixed") 

y1 = Format$(ActiveDocument.ActiveView.BottomRightY, "Fixed") 

MsgBox "View is (" & x0 & ", " & y0 & ") - (" & x1 & ", " & y1 & ")

```

End Sub

**相关主题**

[View.BottomRightX](#page-271-0) 属性

[View.TopLeftX](#page-277-0) 属性

[View.TopLeftY](#page-278-0) 属性

### View.Name 属性

此属性返回视图名称。

**用法**

Name As String

**参数**

无

**描述**

例如，在 SailWind Logic 中此函数返回字符串 "Current View"。

此属性是 View 对象的默认属性。

**示例**

以下示例代码获取当前视图的名称。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

MsgBox ActiveDocument.ActiveView.Name

### View.Parent 属性

此属性返回对象的父对象。

**用法**

Parent As Document 第419页

**参数**

无

**描述**

无

### View.PointerX 属性

此属性返回指针的 X 位置。

**用法**

PointerX (*unit* as PlogUnit) as Double

**参数**

| 参数  | 描述                                                                                                         |
|-------|-------------------------------------------------------------------------------------------------------------|
| unit  | [可选 第314页] 返回值的单位。此可选参数默认为 plogUnitCurrent                                               |

**示例**

此示例也包含 PointerY 属性。

```

' 加载 preview.sch

```

Application.ModelessCommand("s")

DlgModelessCmd.Command="s U5-A"

DlgModelessCmd.OnOk()

doc = Application.ActiveDocument

view = doc.ActiveView

MsgBox view.PointerX & ", " & view.PointerY

### View.PointerY 属性

此属性返回指针的 Y 位置。

**用法**

PointerY (*unit* as PlogUnit) as Double

**参数**

| 参数  | 描述                                                                                                         |
|-------|-------------------------------------------------------------------------------------------------------------|
| unit  | [可选 第314页] 返回值的单位。此可选参数默认为 plogUnitCurrent                                               |

**示例**

此示例也包含 PointerX 属性。

```

' 加载 preview.sch

```

Application.ModelessCommand("s")

DlgModelessCmd.Command="s U5-A"

DlgModelessCmd.OnOk()

doc = Application.ActiveDocument

view = doc.ActiveView

MsgBox view.PointerX & ", " & view.PointerY

### View.TopLeftX 属性

该属性返回视图左上角的 x 坐标。

**用法**

TopLeftX([*unit* As [PlogUnit\]](#page-287-0)) As Double

**参数**

| 参数  | 描述                                                       |
|-------|-----------------------------------------------------------|
| unit  | [可选参数 314页] 返回 x 坐标的单位。                       |

**描述**

无

**示例**

以下示例代码获取打开的原理图中当前视图的坐标。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

x0 = Format$(ActiveDocument.ActiveView.TopLeftX, "Fixed") 

y0 = Format$(ActiveDocument.ActiveView.TopLeftY, "Fixed") 

x1 = Format$(ActiveDocument.ActiveView.BottomRightX, "Fixed") 

y1 = Format$(ActiveDocument.ActiveView.BottomRightY, "Fixed") 

MsgBox "View is (" & x0 & ", " & y0 & ") - (" & x1 & ", " & y1 & ")

```

End Sub

**相关主题**

[View.BottomRightX](#page-271-0) 属性

[View.BottomRightY](#page-272-0) 属性

[View.TopLeftY](#page-278-0) 属性

### View.TopLeftY 属性

该属性返回视图左上角的 y 坐标。

**用法**

TopLeftY([*unit* As [PlogUnit\]](#page-287-0)) As Double

**参数**

| 参数  | 描述                                                             |
|-------|-----------------------------------------------------------------|
| unit  | [可选参数 314页] 返回左上角 Y 值的单位。                         |

**描述**

无

**示例**

以下示例代码获取打开的原理图中当前视图的坐标。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

x0 = Format$(ActiveDocument.ActiveView.TopLeftX, "Fixed")

y0 = Format$(ActiveDocument.ActiveView.TopLeftY, "Fixed")

x1 = Format$(ActiveDocument.ActiveView.BottomRightX, "Fixed")

y1 = Format$(ActiveDocument.ActiveView.BottomRightY, "Fixed")

MsgBox "View is (" & x0 & ", " & y0 & ") - (" & x1 & ", " & y1 & ")

```

End Sub

**相关主题**

[View.BottomRightX](#page-271-0) 属性

[View.BottomRightY](#page-272-0) 属性

[View.TopLeftX](#page-277-0) 属性

### View.Pan 方法

此方法将视图平移至指定位置。

**用法**

Pan(*x* As Double, *y* As Double, [*unit* As [PlogUnit\]](#page-287-0))

**参数**

| 参数  | 描述                                                       |
|-------|-----------------------------------------------------------|
| x     | 要平移到的点的 x 坐标。                                    |
| y     | 要平移到的点的 y 坐标。                                    |
| unit  | [可选参数 314页] X 和 Y 值的单位。                         |

**描述**

**无**

**示例**

以下示例代码将视图居中到组件 U1 的第一个门的位置（假设它存在于打开的原理图中）。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

x = ActiveDocument.Components("U1").Gate(1).PositionX

y = ActiveDocument.Components("U1").Gate(1).PositionY

```

ActiveDocument.ActiveView.Pan(x,y)

End Sub

**相关主题**

### View.Refresh 方法

此方法刷新视图。

**用法**

Refresh

**参数**

无

**描述**

无

**示例**

以下示例代码重绘 SailWind Logic 工作区。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

ActiveDocument.ActiveView.Refresh

### View.SetExtents 方法

此方法设置视图范围。

**用法**

SetExtents(*tlx* As Double, *tly* As Double, *brx* As Double, *bry* As Double, [*unit* As [PlogUnit\]](#page-287-0))

**参数**

| 参数  | 描述                                                                      |
|-------|--------------------------------------------------------------------------|
| tlx   | 新视图左上角的 x 坐标。                                                   |
| tly   | 新视图左上角的 y 坐标                                                     |
| brx   | 新视图右下角的 x 坐标。                                                   |
| bry   | 新视图右下角的 y 坐标                                                     |
| unit  | [可选参数 314页] tlx、tly、brx 和 bry 值的单位。                          |

**描述**

无

**示例**

以下示例代码将视图设置为活动图纸中连接到第一个网络的所有引脚的范围。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

```

Sub Main

xMin = 100000000.0 

yMin = 100000000.0 

xMax = -100000000.0 

yMax = -100000000.0 

ActiveDocument.ActiveSheet.Net(1).Selected =True 

for Each nextPin In ActiveDocument.ActiveSheet.Net(1).Pin 

 'ensure pin is visible i.e. belongs to gate

```

If Not nextPin.Gate Is Nothing Then

'ensure gate is used

If Not nextPin.Gate.Sheet Is Nothing Then

If nextPin.PositionX < xMin Then xMin = nextPin.PositionX

If nextPin.PositionX > xMax Then xMax = nextPin.PositionX

If nextPin.PositionY < yMin Then yMin = nextPin.PositionY

If nextPin.PositionY > yMax Then yMax = nextPin.PositionY

End If

End If

Next nextPin

ActiveDocument.ActiveView.SetExtents(xMin, yMin, xMax, yMax)

End Sub

**相关主题**

[View.SetExtentsToAll](#page-283-0) 方法

[View.SetExtentsToSheet](#page-284-0) 方法

### View.SetExtentsToAll 方法

此方法将视图范围设置为原理图中的所有对象。

**用法**

SetExtentsToAll ()

**参数**

无

**描述**

无

**示例**

以下示例代码将视图范围设置为打开的原理图中的所有对象。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

ActiveDocument.ActiveView.SetExtentsToAll

End Sub

**相关主题**

[View.SetExtents](#page-281-0) 方法

[View.SetExtentsToSheet](#page-284-0) 方法

### View.SetExtentsToSheet 方法

此方法将视图范围设置为活动图纸的范围。

**用法**

SetExtentsToSheet ()

**参数**

无

**描述**

无

**示例**

以下示例代码将视图范围设置为活动图纸。有关运行此示例的更多信息，请参阅第21页的"代码示例"。

Sub Main

ActiveDocument.ActiveView.SetExtentsToSheet

End Sub

**相关主题**

[View.SetExtents](#page-281-0) 方法

[View.SetExtentsToAll](#page-283-0) 方法

### View.Change 事件

当当前视图发生变化时触发此事件。

**用法**

View_Change()

**参数**

无

**描述**

此事件在当前视图或活动图纸变更后触发。

**相关主题**

[View.Pan](#page-279-0) 方法

[View.SetExtents](#page-281-0) 方法

[View.SetExtentsToSheet](#page-284-0) 方法

[View.SetExtentsToAll](#page-283-0) 方法

## 常量

以下是可用的常量。

**PlogDocumentColor**

可能取值为：

| plogDocumentColorBackground       | = 0, |
|-----------------------------------|------|
| plogDocumentColorSelection        | = 1, |
| plogDocumentColorConnection       | = 2, |
| plogDocumentColorBus              | = 3, |
| plogDocumentColorLine             | = 4, |
| plogDocumentColorPart             | = 5, |
| plogDocumentColorHierarchicalComp | = 6, |
| plogDocumentColorText             | = 7, |
| plogDocumentColorTextBox          | = 8, |

| plogDocumentColorRefDes       | = 9,  |
|-------------------------------|-------|
| plogDocumentColorRefDesBox    | = 10, |
| plogDocumentColorPartType     | = 11, |
| plogDocumentColorPartTypeBox  | = 12, |
| plogDocumentColorPartText     | = 13, |
| plogDocumentColorPartTextBox  | = 14, |
| plogDocumentColorPinNumber    | = 15, |
| plogDocumentColorPinNumberBox | = 16, |
| plogDocumentColorNetName      | = 17, |
| plogDocumentColorNetNameBox   | = 18, |
| plogDocumentColorField        | = 19, |
| plogDocumentColorFieldBox     | = 20, |

**PlogObjectType**

SailWind Logic 数据库对象类型。可能取值为：

• plogObjectTypeUnknown = 0 -

服务器可能返回此值表示无效对象。客户端在处理空对象集合时可能使用此值。

- plogObjectTypeComponent = 1

- plogObjectTypeNet = 2

- plogObjectTypePin = 3

- plogObjectTypeGate = 4

- plogObjectTypePartType = 8

- plogObjectTypeLibrary = 18

- plogObjectTypeLibraryItem = 19

- plogObjectTypeApplication = 20

• plogObjectTypeAll = 9999 -

包含所有 SailWind Logic Automation 数据库对象类型，包括 Component、Net、Pin、Gate 和 PartType 对象。

**PlogUnit**

SailWind Logic 单位类型。可能取值为：

| 单位类型                | 描述                                                   |  |
|-----------------------|-------------------------------------------------------|--|
| plogUnitCurrent = 0   | SailWind Logic 当前使用的单位。                    |  |
| plogUnitDatabase = 1  | SailWind Logic 内部数据库单位。                        |  |
| plogUnitMils = 2      | 密尔单位 (1/1000英寸)。                              |  |
| plogUnitInch = 3      | 英寸单位。                                                    |  |
| plogUnitMetric = 4    | 公制单位 (1/1000米)。                            |  |
| plogUnitDrawArea = 99 | SailWind Logic 工作区内部单位。请勿使用此值。 |  |

**PlogGridType**

SailWind Logic 网格类型。可能取值为：

| 网格类型          | 描述                  |
|---------------------|------------------------------|
| plogGridNone = 0    | 无网格。                        |
| plogGridDesign = 1  | SailWind Logic 设计网格。  |
| plogGridDisplay = 3 | SailWind Logic 显示网格。 |
| plogGridAll = 9999  | 所有 SailWind Logic 网格。    |

**PlogPinElectricalType**

SailWind Logic 门电路电气类型。可能取值为：

- plogElectricalTypeUnknown = 0

- plogElectricalTypeSource = 1

- plogElectricalTypeBidirectional = 2

- plogElectricalTypeOpenCollector = 3

- plogElectricalTypeOrTieableSource = 4

- plogElectricalTypeTristate = 5

- plogElectricalTypeLoad = 6

- plogElectricalTypeTerminator = 7

- plogElectricalTypePower = 8

- plogElectricalTypeGround = 9

**PlogASCIIVersion**

SailWind Logic ASCII 版本。可能取值为：

- plogASCIIVerCurrent = 0

- plogASCIIVer1_2 = 1

- plogASCIIVer3_0 = 2

- plogASCIIVer3_5 = 3

**PlogNetListVersion**

SailWind Logic/SailWind Layout 网表版本。可能取值为：

- plogPowerPCBNetListVerCurrent = 0

- plogPowerPCBNetListVer2_1 = 2

- plogPowerPCBNetListVer3_0 = 3

- plogPowerPCBNetListVer3_5 = 4

**PlogGateVisibility**

SailWind Logic 门电路可见性项。可能取值为：

- plogAttrVisibility = 0

- plogAttrNameVisibility = 1

- plogRefDesVisibility = 2

- plogPartTypeVisibility = 3

- plogPinNumberVisibility = 4

- plogPinNameVisibility = 5

- plogPCBDecalVisibility = 6

- plogPCBDecalNameVisibility = 7

**PlogDefaultPosition**

SailWind Logic 默认位置坐标。可能取值为：

plogDefaultPositionX = &H80000000

plogDefaultPositionY = &H80000000

**PlogMeasureFormat**

SailWind Logic 默认位置坐标。可能取值为：

plogMeasureFormatStandard = 0

plogMeasureFormatCurrent = 1

plogMeasureFormatShort = 2

plogMeasureFormatLong = 3

**PlogLibraryItemType**

SailWind Logic 库项目类型。可能取值为：

- plogLibraryItemTypePartType = 0

- plogLibraryItemTypeDecal = 1

- plogLibraryItemTypeLogicDrawing = 2

- plogLibraryItemTypeDrawing = 3

- plogLibraryItemTypeAll = 9999

## 可选参数

可选参数是 SailWind Logic 中属性或方法的参数，可以省略而使用默认值。默认值是统计上最常用的值或最能代表该参数的值。

例如，如果方法 M([arg1], [arg2]) 的两个参数 arg1 和 arg2 都是可选的，您可以用以下四种方式调用该方法：

| 参数示例                        | 描述                                                            |
|------------------------------------------|------------------------------------------------------------------------|
| M()                                      | 传递 arg1 和 arg2 的默认值。                   |
| M( <value1>)                    | 传递 arg1 的 <value1> 和 arg2 的默认值。          |
| M(, <value2>)                   | 传递 arg1 的默认值和 arg2 的 <value2>。          |
| M( <value1>, <value2>) | 同时传递 arg1 的 <value1> 和 arg2 的 <value2>。 |

**表 1. 可选参数示例**

## 异常

Automation 异常是来自 SailWind Logic 服务器的特殊通知，用于向 Automation 客户端发出错误信号。例如，如果 Basic 脚本尝试删除不存在的属性，SailWind Logic 会生成异常。

当脚本接收到异常时，会发生以下情况：

- 如果使用 Basic On Error 语句实现了异常处理程序，客户端代码的执行流程将重定向到异常处理程序。

- 如果未实现异常处理程序，则调用客户端的默认处理程序。在所有 Basic 解释器中，默认处理程序是在生成异常的行设置断点。

## 变体

变体是一种可以包含或表示任何类型数据的数据类型，例如布尔值、整数值、长整型值、双精度值、字符串值或数组。

在 SailWind Logic Automation 中，变体数据类型在两种不同情况下使用：

- 当属性或方法的参数是可以表示为不同值和值类型的数据时。例如，集合中的特定对象可以通过其在集合中的索引（长整型值）或其名称（字符串值）来引用。

- 当属性或方法的参数或返回值是复杂类型，而不是由 SailWind Logic Basic 对象明确定义的类型时。例如，点数组表示为变体。

```