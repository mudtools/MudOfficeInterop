//
// MudTools.OfficeInterop 项目的版权、商标、专利和其他相关权利均受相应法律法规的保护。使用本项目应遵守相关法律法规和许可证的要求。
//
// 本项目主要遵循 MIT 许可证和 Apache 许可证（版本 2.0）进行分发和使用。许可证位于源代码树根目录中的 LICENSE-MIT 和 LICENSE-APACHE 文件。
//
// 不得利用本项目从事危害国家安全、扰乱社会秩序、侵犯他人合法权益等法律法规禁止的活动！任何基于本项目二次开发而产生的一切法律纠纷和责任，我们不承担任何责任！

namespace MudTools.OfficeInterop.Word;

/// <summary>
/// 表示Word文档中数学对象的换行符功能接口
/// </summary>
[ComObjectWrap(ComNamespace = "MsWord")]
public interface IWordOMathBreak : IOfficeObject<IWordOMathBreak>, IDisposable
{
    /// <summary>
    /// 获取与此数学换行符关联的Word应用程序实例
    /// </summary>
    /// <value>返回IWordApplication类型的对象，可能为null</value>
    [ComPropertyWrap(NeedDispose = false)]
    IWordApplication? Application { get; }

    /// <summary>
    /// 获取此数学换行符的父对象
    /// </summary>
    /// <value>返回父对象，可能为null</value>
    object? Parent { get; }

    /// <summary>
    /// 获取与此数学换行符关联的Word文档范围
    /// </summary>
    /// <value>返回IWordRange类型的对象</value>
    IWordRange Range { get; }

    /// <summary>
    /// 获取或设置数学公式中换行符的对齐位置
    /// </summary>
    /// <value>返回或设置一个整数值表示对齐位置</value>
    int AlignAt { get; set; }

    /// <summary>
    /// 删除此数学换行符
    /// </summary>
    void Delete();
}