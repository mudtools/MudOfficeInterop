# MudTools.OfficeInterop

核心模块，提供 Office 应用程序的基础接口和通用功能。

## 项目概述

MudTools.OfficeInterop 是整个 Office 互操作库的核心模块，提供了 Office 应用程序的基础接口和通用功能。该模块封装了 Office 核心组件的常用操作，为其他 Office 应用程序模块（Excel、Word、PowerPoint）提供基础支撑。

此外，该模块还提供了 Office UI 相关组件的封装，包括功能区(Ribbon)和自定义任务窗格(CTP)，方便开发者创建 Office 插件时使用。

## 主要功能

- 提供 Office 应用程序的基础接口和通用功能
- 封装 Office 核心组件的常用操作
- 提供 Office UI 相关组件的封装，包括功能区(Ribbon)和自定义任务窗格(CTP)
- 为 Excel、Word、PowerPoint 模块提供基础接口支持

## 支持的框架

- .NET Framework 4.6.2
- .NET Framework 4.7
- .NET Framework 4.8
- .NET Standard 2.1
- .NET 6.0-windows
- .NET 7.0-windows
- .NET 8.0-windows
- .NET 9.0-windows

## 安装

```xml
<PackageReference Include="MudTools.OfficeInterop" Version="1.1.8" />
```

## 核心组件

### OfficeUIFactory

[OfficeUIFactory](https://gitee.com/mudtools/OfficeInterop/tree/master/MudTools.OfficeInterop/OfficeUIFactory.cs#L16-L51) 是用于创建 Office UI 相关组件的工厂类，提供以下方法：

- `CreateCTPFactory` - 创建自定义任务窗格工厂的包装器实例
- `CreateRibbonUI` - 创建功能区 UI 的包装器实例
- `CreateRibbonControl` - 创建功能区控件的包装器实例

### 使用示例

#### 使用自定义任务窗格

```csharp
// 创建自定义任务窗格
var ctpFactory = OfficeUIFactory.CreateCTPFactory(officeCTPFactory);
var ctp = ctpFactory.CreateCTP("MyAddin.UserControl", "我的任务窗格");

// 设置任务窗格属性
ctp.Visible = true;
ctp.Width = 200;

// 显示任务窗格
ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
```

#### 使用功能区控件

```csharp
// 处理功能区控件事件
public void OnRibbonButtonClicked(IRibbonControl control)
{
    switch (control.Id)
    {
        case "buttonNewDocument":
            // 处理按钮点击事件
            break;
        case "buttonOpenDocument":
            // 处理打开文档事件
            break;
    }
}
```

## 许可证

本项目采用双重许可证模式：

- [MIT 许可证](../../LICENSE-MIT)

## 免责声明

本项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。

不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任。