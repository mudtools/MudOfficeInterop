# MudTools.OfficeInterop.Vbe

VBE (Visual Basic Editor) 操作模块，提供对 Office 应用程序中 VBA 项目和代码的编程访问。

## 项目概述

MudTools.OfficeInterop.Vbe 是专门用于操作 Microsoft Visual Basic Editor (VBE) 的 .NET 封装库。该模块提供了对 VBE 对象模型的完整封装，包括 VBProject、VBComponent、CodeModule 等核心对象的管理。

通过使用本模块，开发者可以编程方式访问和操作 Office 应用程序中的 VBA 项目、组件和代码模块，实现自动化代码生成、宏管理等功能。

## 主要功能

- VBE 应用程序对象操作接口
- VBProject 项目管理功能
- VBComponent 组件管理功能
- CodeModule 代码模块操作功能
- VBA 引用管理功能

## 支持的框架

- .NET Framework 4.6.2
- .NET Framework 4.7
- .NET Framework 4.8
- .NET Standard 2.1

## 安装

```xml
<PackageReference Include="MudTools.OfficeInterop.Vbe" Version="1.1.0" />
```

## 核心组件

### VBE 对象模型接口

VBE 模块提供以下核心接口：

- [IVbeApplication](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeApplication.cs#L12-L101) - VBE 应用程序对象接口
- [IVbeVBProject](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeVBProject.cs#L12-L174) - VB 项目对象接口
- [IVbeVBProjects](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeVBProjects.cs#L12-L172) - VB 项目集合接口
- [IVbeVBComponent](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeVBComponent.cs#L12-L124) - VB 组件对象接口
- [IVbeVBComponents](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeVBComponents.cs#L12-L168) - VB 组件集合接口
- [IVbeCodeModule](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeCodeModule.cs#L12-L162) - 代码模块对象接口
- [IVbeReference](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeReference.cs#L12-L110) - 引用对象接口
- [IVbeReferences](file:///D:/Repos/MudTools.OfficeInterop/MudTools.OfficeInterop.Vbe/IVbeReferences.cs#L12-L166) - 引用集合接口

## 使用示例

### 访问 VBA 项目

```csharp
// 通过 Excel 应用程序访问 VBE
using var excelApp = ExcelFactory.CreateApplication();
var vbeApp = excelApp.Vbe;

// 获取所有 VBA 项目
var vbProjects = vbeApp.VBProjects;
foreach (var project in vbProjects)
{
    Console.WriteLine($"Project: {project.Name}");
    Console.WriteLine($"Type: {project.Type}");
    Console.WriteLine($"Description: {project.Description}");
}
```

### 操作 VBA 组件

```csharp
// 获取活动的 VBA 项目
var activeProject = vbeApp.ActiveVBProject;

// 遍历项目中的所有组件
var components = activeProject.VBComponents;
foreach (var component in components)
{
    Console.WriteLine($"Component: {component.Name}");
    Console.WriteLine($"Type: {component.Type}");
    
    // 获取组件的代码模块
    var codeModule = component.CodeModule;
    Console.WriteLine($"Lines of code: {codeModule.CountOfLines}");
}
```

### 添加和修改代码

```csharp
// 获取或添加一个标准模块
var components = activeProject.VBComponents;
var module = components.Add(vbext_ComponentType.vbext_ct_StdModule);
module.Name = "GeneratedModule";

// 在模块中添加代码
var codeModule = module.CodeModule;
codeModule.AddFromString(@"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
");
```

### 管理引用

```csharp
// 获取项目引用
var references = activeProject.References;

// 添加新引用
try 
{
    references.AddFromGuid("{000204EF-0000-0000-C000-000000000046}", 4, 1); // VBA 扩展库
    Console.WriteLine("Reference added successfully");
}
catch (Exception ex)
{
    Console.WriteLine($"Failed to add reference: {ex.Message}");
}

// 列出所有引用
foreach (var reference in references)
{
    Console.WriteLine($"Reference: {reference.Name}");
    Console.WriteLine($"Description: {reference.Description}");
    Console.WriteLine($"FullPath: {reference.FullPath}");
}
```

## 常见应用场景

1. **自动化代码生成** - 动态生成 VBA 代码以实现特定功能
2. **宏管理** - 管理和维护 Office 文档中的宏代码
3. **插件开发** - 为 Office 应用程序开发 VBA 插件
4. **代码分析** - 分析现有 VBA 项目的代码结构和内容
5. **批量操作** - 批量修改多个 Office 文档中的 VBA 代码

## 注意事项

1. 使用 VBE 功能需要启用 Office 应用程序的程序化访问权限
2. 某些操作可能需要管理员权限
3. 操作 VBA 代码时应格外小心，避免破坏现有功能
4. 建议在操作前备份相关文件

## 许可证

本项目采用双重许可证模式：

- [MIT 许可证](../../LICENSE-MIT)
- [Apache 许可证 2.0](../../LICENSE-APACHE)

## 免责声明

本项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。

不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任。