//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word中的OMath方程框接口，继承自IWordOMath接口
/// OMathBox是用于在Office文档中创建和操作数学框的接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathBox : IOfficeObject<IWordOMathBox>, IDisposable
{
    /// <summary>
    /// 获取与此OMathBox关联的Word应用程序对象
    /// </summary>
    /// <value>返回IWordApplication接口的实例或null</value>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此OMathBox的父对象
    /// </summary>
    /// <value>返回OMathBox的父对象或null</value>
    object? Parent { get; }

    /// <summary>
    /// 获取或设置OMathBox中包含的数学表达式对象
    /// </summary>
    /// <value>返回IWordOMath接口的实例或null</value>
    IWordOMath? E { get; }

    /// <summary>
    /// 获取或设置操作数模拟（Operator Emulator）的显示状态
    /// </summary>
    /// <value>true表示启用操作数模拟，false表示禁用</value>
    bool OpEmu { get; set; }

    /// <summary>
    /// 获取或设置是否禁止在OMathBox中换行
    /// </summary>
    /// <value>true表示禁止换行，false表示允许换行</value>
    bool NoBreak { get; set; }

    /// <summary>
    /// 获取或设置差分符号的显示状态
    /// </summary>
    /// <value>true表示启用差分显示，false表示禁用差分显示</value>
    bool Diff { get; set; }
}